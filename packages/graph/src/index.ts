import { type AuthStrategy, createGraphRequest } from "@sapientsai/ms-graph-core"
import dotenv from "dotenv"
import { createServer, type SomaServerInstance } from "somamcp"
import { z } from "zod"

import { createAppOnlyAuthStrategy } from "./auth/app-only-strategy"
import { resolveServerRuntimeConfig, type ServerRuntimeConfig } from "./config"
import { buildMicrosoftGraphBatchTool, buildMicrosoftGraphTool } from "./tools/graph-passthrough"
import { buildReadDocumentTool } from "./tools/read-document"

dotenv.config({ quiet: true })

declare const __VERSION__: string
const VERSION = (typeof __VERSION__ !== "undefined" ? __VERSION__ : "0.0.0-dev") as `${number}.${number}.${number}`

// Scaffolding-stage server: app-only auth wired through core's AuthStrategy, on the
// somamcp shell. Phase 3 will add the generic Graph passthrough, $batch, read_document
// extraction, AI Search, and the /upload route (with self-applied API-key auth — see the
// somamcp spike note in SAI_PLAN_ms-graph-monorepo).
export const buildServer = (config: ServerRuntimeConfig, auth: AuthStrategy): SomaServerInstance => {
  const server = createServer({
    name: "ms-graph-mcp-server",
    version: VERSION,
    instructions: "Minimal app-only Microsoft Graph MCP server.",
    ...(config.apiKey
      ? {
          // FastMCP signals an HTTP rejection by throwing a Response; surface a 401 on a
          // bad/missing key. Non-async (returns a Promise) to satisfy the no-floating-await rule.
          authenticate: (request: unknown): Promise<{ apiKey: string }> => {
            const token = extractAuthHeader(request)?.replace(/^Bearer\s+/i, "")
            if (!token || token !== config.apiKey) {
              return Promise.reject(new Response("Unauthorized", { status: 401 }))
            }
            return Promise.resolve({ apiKey: token })
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

  return server
}

// somamcp/FastMCP hands `authenticate` an http.IncomingMessage on the MCP path; tolerate
// both that and a Hono Request (the artifact path) per the spike note.
const extractAuthHeader = (request: unknown): string | undefined => {
  const headers = (request as { headers?: unknown })?.headers
  if (headers && typeof (headers as { get?: unknown }).get === "function") {
    return (headers as Headers).get("authorization") ?? undefined
  }
  const h = headers as Record<string, string | string[] | undefined> | undefined
  const raw = h?.authorization ?? h?.Authorization
  return Array.isArray(raw) ? raw[0] : raw
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
