import { type AuthStrategy, createGraphRequest } from "@sapientsai/ms-graph-core"
import dotenv from "dotenv"
import { createServer, getRequestHeader, type SomaServerInstance } from "somamcp"
import { z } from "zod"

import { authorizesWithApiKey } from "./auth/api-key-gate"
import { createAppOnlyAuthStrategy } from "./auth/app-only-strategy"
import { resolveServerRuntimeConfig, type ServerRuntimeConfig } from "./config"
import { buildAiSearchTool } from "./tools/ai-search"
import { buildMicrosoftGraphBatchTool, buildMicrosoftGraphTool } from "./tools/graph-passthrough"
import { buildReadDocumentTool } from "./tools/read-document"
import { buildSharePointSearchTool } from "./tools/sharepoint-search"
import { buildGetUploadConfigTool } from "./tools/upload-config"
import { buildUploadRoute } from "./upload/upload-route"

dotenv.config({ quiet: true })

declare const __VERSION__: string
const VERSION = (typeof __VERSION__ !== "undefined" ? __VERSION__ : "0.0.0-dev") as `${number}.${number}.${number}`

// App-only Microsoft Graph server on the somamcp shell: auth wired through core's
// AuthStrategy, tools for the generic Graph passthrough, $batch, read_document extraction,
// SharePoint/AI Search, and a protected /upload relay route (POST/PUT). The httpStream
// transport and /upload share one `authenticate` gate via somamcp's `protected` routes.
export const buildServer = (config: ServerRuntimeConfig, auth: AuthStrategy): SomaServerInstance => {
  // Capture as a const so the truthy-narrowing to `string` survives into the closure.
  const { apiKey } = config
  const server = createServer({
    name: "ms-graph-mcp-server",
    version: VERSION,
    instructions: "Minimal app-only Microsoft Graph MCP server.",
    ...(apiKey
      ? {
          // Shared gate for the httpStream transport AND the protected /upload route.
          // getRequestHeader hides the http.IncomingMessage (transport) vs Hono Request
          // (route) shape difference. FastMCP signals an HTTP rejection by throwing a
          // Response; surface a 401 on a bad/missing key. Non-async (returns a Promise) to
          // satisfy the no-floating-await rule.
          authenticate: (request: unknown): Promise<{ apiKey: string }> => {
            const bearer = getRequestHeader(request, "authorization")?.replace(/^Bearer\s+/i, "")
            if (!authorizesWithApiKey(bearer, apiKey)) {
              return Promise.reject(new Response("Unauthorized", { status: 401 }))
            }
            return Promise.resolve({ apiKey })
          },
        }
      : {}),
  })

  // Reports whether the server can acquire an app-only Microsoft Graph token (no leak).
  server.addTool({
    name: "get_auth_status",
    description: "Report whether the server can acquire an app-only Microsoft Graph token.",
    parameters: z.object({}),
    execute: async () => {
      const result = await auth.getAccessToken()
      return JSON.stringify(
        result.fold(
          (error) => ({ ok: false, mode: "app-only", error: error.message }),
          () => ({ ok: true, mode: "app-only" }),
        ),
      )
    },
  })

  // Generic Graph plumbing from core, driven by the app-only strategy.
  const graph = createGraphRequest(auth)
  server.addTool(buildMicrosoftGraphTool(graph, process.env.MCP_INSTRUCTIONS))
  server.addTool(buildMicrosoftGraphBatchTool(graph))
  server.addTool(buildReadDocumentTool(auth))
  server.addTool(buildSharePointSearchTool(graph, config.sharePointSearch))

  // Azure AI Search — optional, only when AZURE_AI_SEARCH_* env is configured.
  if (config.aiSearch) server.addTool(buildAiSearchTool(config.aiSearch))

  // Binary upload relay: a first-class somamcp protected route (POST/PUT /upload) that
  // inherits the shared `authenticate` gate, plus get_upload_config which hands back an
  // opaque ticket (never the raw key) resolving through that same gate.
  server.addTool(buildGetUploadConfigTool(config.publicBaseUrl, config.apiKey))
  server.addRoute(buildUploadRoute(auth, config.apiKey))

  return server
}

export const main = async (): Promise<void> => {
  const configResult = resolveServerRuntimeConfig()
  if (configResult.isLeft()) {
    console.error(`[Config] ${configResult.value as string}`)
    process.exit(1)
  }
  const config = configResult.value as ServerRuntimeConfig
  const auth = createAppOnlyAuthStrategy(config.auth)
  const server = buildServer(config, auth)

  if (config.transport === "httpStream") {
    await server.start({
      transportType: "httpStream",
      httpStream: { port: config.port, host: config.host, endpoint: "/mcp" },
    })
    console.error(`[Server] ms-graph-mcp-server v${VERSION} (app-only) on ${config.host}:${config.port}`)
  } else {
    await server.start({ transportType: "stdio" })
    console.error(`[Server] ms-graph-mcp-server v${VERSION} (app-only) on stdio`)
  }
}
