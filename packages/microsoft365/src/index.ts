import dotenv from "dotenv"
import { FastMCP } from "fastmcp"
// AzureSession shape: { accessToken: string; scopes: string[]; refreshToken?: string; upn?: string }
type OAuthSessionContext = { accessToken?: string }
import {
  decodeBase64Upload,
  describeFetchError,
  filenameFromPath,
  MAX_UPLOAD_SIZE,
  resolveUploadContentType,
  sessionUpload,
  SIMPLE_UPLOAD_LIMIT,
  simpleUpload,
} from "@sapientsai/ms-graph-core"
import { type Either, Left, Right } from "functype/either"

import { getAccessToken, initializeAuth } from "./auth"
import { createAzureAuthProvider } from "./auth/oauth-provider"
import { GRAPH_API_BASE } from "./auth/scopes"
import { withToken } from "./auth/token-context"
import { initializeGraphClient } from "./client/graph-client"
import { toolDefinitions } from "./tools/definitions"
import type { ToolDefinition } from "./tools/tool-definitions"
import { filterTools, type ToolFilterConfig } from "./tools/tool-registry"
import type { AuthConfig } from "./types"
import { resolveUploadAccessToken } from "./upload/upload-auth"
import { auditToolCall, auditToolError, auditToolResult } from "./utils/audit"

dotenv.config({ quiet: true })

declare const __VERSION__: string
const VERSION = (typeof __VERSION__ !== "undefined" ? __VERSION__ : "0.0.0-dev") as `${number}.${number}.${number}`

const resolveAuthConfig = (): AuthConfig => {
  const mode = process.env.MS365_AUTH_MODE ?? "interactive"
  const tenantId = process.env.MS365_TENANT_ID ?? "common"
  const clientId = process.env.MS365_CLIENT_ID ?? ""

  switch (mode) {
    case "certificate":
      return {
        mode: "certificate",
        tenantId,
        clientId,
        certPath: process.env.MS365_CERT_PATH ?? "",
        certPassword: process.env.MS365_CERT_PASSWORD,
      }
    case "client-secret":
      return {
        mode: "client-secret",
        tenantId,
        clientId,
        clientSecret: process.env.MS365_CLIENT_SECRET ?? "",
      }
    case "client-token":
      return {
        mode: "client-token",
        accessToken: process.env.MS365_ACCESS_TOKEN,
      }
    case "oauth-proxy":
      return {
        mode: "oauth-proxy",
        tenantId,
        clientId,
        clientSecret: process.env.MS365_CLIENT_SECRET ?? "",
        baseUrl: process.env.MS365_OAUTH_BASE_URL ?? "http://localhost:3000",
      }
    default:
      return {
        mode: "interactive",
        tenantId,
        clientId,
        redirectUri: process.env.MS365_REDIRECT_URI,
      }
  }
}

const setupAuth = async () => {
  const config = resolveAuthConfig()
  const result = await initializeAuth(config)

  result.fold(
    (error) => {
      if (config.mode === "client-token" && !config.accessToken) {
        console.error("[Setup] Client token mode: use set_access_token tool to provide a token")
      } else {
        console.error(`[Error] Authentication failed: ${(error as { message: string }).message}`)
        process.exit(1)
      }
    },
    () => console.error(`[Setup] Authentication initialized (${config.mode} mode)`),
  )

  initializeGraphClient({ getAccessToken })
  console.error("[Setup] Graph client initialized")
}

const resolveFilterConfig = (transport: "stdio" | "httpStream"): ToolFilterConfig => ({
  presets: process.env.MS365_PRESETS?.split(",").map((s) => s.trim()),
  enabledPattern: process.env.MS365_ENABLED_TOOLS,
  readOnly: process.env.MS365_READ_ONLY === "true",
  orgMode: process.env.MS365_ORG_MODE === "true",
  requireDraft: process.env.MS365_REQUIRE_DRAFT === "true",
  transport,
})

type ExecuteContext = { session?: OAuthSessionContext }

const wrapExecute = (tool: ToolDefinition, oauthMode: boolean): never => {
  const baseFn = tool.execute as (p: Record<string, unknown>) => Promise<string>

  // Layer 1: OAuth token injection (wraps the base function)
  const withOAuth = oauthMode
    ? (params: Record<string, unknown>, context: ExecuteContext) =>
        withToken(context.session?.accessToken, () => baseFn(params))
    : (params: Record<string, unknown>) => baseFn(params)

  // Layer 2: Audit logging (wraps the OAuth-aware function)
  const withAudit = async (params: Record<string, unknown>, context: ExecuteContext) => {
    auditToolCall(tool.name, params)
    const start = Date.now()

    try {
      const result = await withOAuth(params, context)
      auditToolResult(tool.name, true, Date.now() - start)
      return result
    } catch (error) {
      auditToolError(tool.name, error instanceof Error ? error.message : String(error))
      auditToolResult(tool.name, false, Date.now() - start)
      throw error
    }
  }

  return withAudit as never
}

const registerTools = (server: FastMCP, allowedTools: Set<string>, oauthMode: boolean) => {
  const toRegister = toolDefinitions.filter((tool) => allowedTools.has(tool.name))

  toRegister.forEach((tool) => {
    server.addTool({
      name: tool.name,
      description: tool.description,
      parameters: tool.parameters,
      execute: wrapExecute(tool, oauthMode),
      annotations: tool.annotations,
    })
  })

  console.error(
    `[Setup] Tools registered: ${toRegister.length}, skipped: ${toolDefinitions.length - toRegister.length}`,
  )
}

const buildUploadWorkflow = (allowedTools: Set<string>): string => {
  const hasFromPath = allowedTools.has("upload_file_from_path")
  const hasUploadConfig = allowedTools.has("get_upload_config")

  const bullets: string[] = ["- Text content → upload_file (inline text, any transport)"]

  if (hasFromPath) {
    bullets.push(
      "- Binary files from the server's local disk → upload_file_from_path (requires absolute path on this machine)",
      "- Binary files generated in a cloud container (e.g., claude.ai) → first save to the user's local filesystem using Desktop Commander's write_file, then call upload_file_from_path with that local path",
    )
  }

  if (hasUploadConfig) {
    bullets.push(
      "- Binary files from HTTP/SSE deployments → get_upload_config returns an authenticated URL + curl command; execute the curl in a shell to upload without routing bytes through the LLM",
    )
  }

  return `\n\nUpload workflows:\n${bullets.join("\n")}`
}

const buildInstructions = (allowedTools: Set<string>): string => {
  const domains = new Set(toolDefinitions.filter((t) => allowedTools.has(t.name)).map((t) => t.domain))
  const domainDescriptions: Record<string, string> = {
    auth: "Authentication: Check auth status and manage tokens",
    mail: "Mail: List, read, send, reply, search, and draft email messages",
    calendar: "Calendar: List, view, create, update, and delete events",
    contacts: "Contacts: List, view, create, and search contacts",
    files:
      "Files: List, view, search, download OneDrive files; create folders; upload files (see Upload workflows below)",
    chats: "Chats: List Teams chats and messages; send chat messages",
    teams: "Teams: List teams, channels, and messages; send channel messages",
    meetings: "Meetings: List Teams meeting transcripts and read their text",
    users: "Users: View profiles and list users",
    groups: "Groups: List groups and group members",
    planner: "Planner: List plans and tasks; create and update tasks",
    onenote: "OneNote: List notebooks, sections, pages; read page content",
    todo: "To Do: List task lists and tasks; create and update tasks",
    query: "Graph Query: Execute arbitrary Microsoft Graph API queries",
  }

  const capabilities = [...domains]
    .map((d) => domainDescriptions[d])
    .filter(Boolean)
    .map((desc) => `- ${desc}`)
    .join("\n")

  const uploadSection = domains.has("files") ? buildUploadWorkflow(allowedTools) : ""

  return `A Microsoft 365 MCP server via Microsoft Graph API.\n\nAvailable capabilities:\n${capabilities}${uploadSection}`
}

type UploadRequestContext = {
  header: (name: string) => string | undefined
  query: (name: string) => string | undefined
  arrayBuffer: () => Promise<ArrayBuffer>
}

const handleUpload = async (
  req: UploadRequestContext,
  oauthMode: boolean,
): Promise<{ status: number; body: unknown }> => {
  const path = req.query("path")
  if (!path) return { status: 400, body: { error: "path query parameter is required" } }
  if (!/:\/content$/i.test(path)) {
    return { status: 400, body: { error: 'path must end with ":/content"' } }
  }

  const apiVersion = req.query("apiVersion") ?? "v1.0"
  const conflictBehavior = req.query("conflictBehavior") ?? "rename"
  const explicitContentType = req.query("contentType")
  const encoding = req.query("encoding")

  const authHeader = req.header("authorization") ?? req.header("Authorization")
  const bearer = authHeader?.replace(/^Bearer\s+/i, "")
  const auth = await resolveUploadAccessToken(oauthMode, bearer)
  if (auth.error || !auth.token) {
    return { status: auth.status ?? 401, body: { error: auth.error ?? "Unauthorized" } }
  }

  const rawBufferResult = await (async (): Promise<Either<string, Buffer>> => {
    try {
      return Right(Buffer.from(await req.arrayBuffer()))
    } catch (error) {
      return Left(`Failed to read request body: ${error instanceof Error ? error.message : String(error)}`)
    }
  })()
  if (rawBufferResult.isLeft()) {
    return { status: 400, body: { error: rawBufferResult.value as string } }
  }
  const rawBuffer = rawBufferResult.value as Buffer
  if (rawBuffer.length === 0) return { status: 400, body: { error: "Empty request body" } }

  const buffer = encoding === "base64" ? decodeBase64Upload(rawBuffer) : rawBuffer
  if (buffer.length === 0) return { status: 400, body: { error: "Invalid base64 content" } }
  if (buffer.length > MAX_UPLOAD_SIZE) {
    return { status: 413, body: { error: `File too large (max ${MAX_UPLOAD_SIZE} bytes)` } }
  }

  const filename = filenameFromPath(path)
  const contentType = resolveUploadContentType(explicitContentType, filename)
  const apiBase = `${GRAPH_API_BASE}/${apiVersion}`

  console.error(
    `[Upload] path=${path} bytes=${buffer.length} contentType=${contentType} mode=${buffer.length <= SIMPLE_UPLOAD_LIMIT ? "simple" : "session"}`,
  )

  const result =
    buffer.length <= SIMPLE_UPLOAD_LIMIT
      ? await simpleUpload(apiBase, path, auth.token, buffer, contentType, conflictBehavior)
      : await sessionUpload(apiBase, path, auth.token, buffer, conflictBehavior)

  if (result.isLeft()) {
    const err = result.value as { message: string; status?: number }
    return { status: err.status ?? 500, body: { error: err.message } }
  }

  return { status: 200, body: result.value }
}

const mountUploadRoute = (server: FastMCP, oauthMode: boolean): void => {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any -- Hono app surface
  const app: any = (server as unknown as { getApp?: () => unknown }).getApp?.()
  if (!app) {
    console.error("[Upload] FastMCP.getApp() unavailable; /upload endpoint not mounted")
    return
  }

  /* eslint-disable @typescript-eslint/no-explicit-any */
  const handler = async (c: any) => {
    try {
      const result = await handleUpload(c.req as UploadRequestContext, oauthMode)
      return c.json(result.body, result.status)
    } catch (err) {
      const { message } = describeFetchError(err)
      console.error("[Upload] unhandled error:", message)
      return c.json({ error: message }, 500)
    }
  }

  app.post("/upload", handler)
  app.put("/upload", handler)
  /* eslint-enable @typescript-eslint/no-explicit-any */
  console.error("[Setup] /upload endpoint mounted (POST, PUT)")
}

// === Server Startup ===
const main = async () => {
  const authConfig = resolveAuthConfig()
  const oauthMode = authConfig.mode === "oauth-proxy"

  const transport: "stdio" | "httpStream" = oauthMode
    ? "httpStream"
    : process.env.TRANSPORT_TYPE === "httpStream"
      ? "httpStream"
      : "stdio"

  const filterConfig = resolveFilterConfig(transport)
  const allowedTools = filterTools(filterConfig)

  if (oauthMode) {
    // OAuth proxy mode: FastMCP handles auth via AzureProvider
    const provider = createAzureAuthProvider({
      baseUrl: (authConfig as { baseUrl: string }).baseUrl,
      clientId: (authConfig as { clientId: string }).clientId,
      clientSecret: (authConfig as { clientSecret: string }).clientSecret,
      tenantId: (authConfig as { tenantId: string }).tenantId,
    })

    const server = new FastMCP({
      name: "microsoft365-mcp-server",
      version: VERSION,
      instructions: buildInstructions(allowedTools),
      auth: provider,
      health: { enabled: true, path: "/ping", message: "ok" },
    } as never)

    // Initialize graph client without credential-based auth (tokens come from session)
    initializeGraphClient({ getAccessToken })

    registerTools(server, allowedTools, true)
    mountUploadRoute(server, true)

    const port = parseInt(process.env.PORT ?? "3000", 10)
    const host = process.env.HOST ?? process.env.FASTMCP_HOST ?? "127.0.0.1"
    await server.start({ transportType: "httpStream", httpStream: { port, host } })
    console.error(`[Server] MS 365 MCP Server v${VERSION} (OAuth proxy) running on ${host}:${port}`)
  } else {
    // Standard mode: credential-based auth
    await setupAuth()

    const server = new FastMCP({
      name: "microsoft365-mcp-server",
      version: VERSION,
      instructions: buildInstructions(allowedTools),
      health: { enabled: true, path: "/ping", message: "ok" },
    })

    registerTools(server, allowedTools, false)

    const transportType = process.env.TRANSPORT_TYPE ?? "stdio"

    if (transportType === "httpStream") {
      mountUploadRoute(server, false)
      const port = parseInt(process.env.PORT ?? "3000", 10)
      const host = process.env.HOST ?? process.env.FASTMCP_HOST ?? "127.0.0.1"
      await server.start({ transportType: "httpStream", httpStream: { port, host } })
      console.error(`[Server] MS 365 MCP Server v${VERSION} running on ${host}:${port}`)
    } else {
      await server.start({ transportType: "stdio" })
      console.error(`[Server] MS 365 MCP Server v${VERSION} running on stdio`)
    }
  }
}

main().catch((error) => {
  console.error("[Fatal]", error)
  process.exit(1)
})
