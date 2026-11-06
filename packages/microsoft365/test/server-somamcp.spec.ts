// The real SomaMCP server, reached through server.fetch with no port. server-env.spec.ts stubs the
// server to test what gets mounted; these cover what a client actually sees over HTTP.
import { mkdtempSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"

import { Right } from "functype/either"
import type { SomaServerInstance } from "somamcp/node"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

const captured = vi.hoisted(() => ({
  server: undefined as SomaServerInstance | undefined,
  start: vi.fn(async () => undefined),
  azureConfig: undefined as Record<string, unknown> | undefined,
}))

vi.mock("somamcp/node", async (original) => {
  const real = await original<typeof import("somamcp/node")>()
  return {
    ...real,
    createAzureProvider: (config: Parameters<typeof real.createAzureProvider>[0]) => {
      captured.azureConfig = config as Record<string, unknown>
      return real.createAzureProvider(config)
    },
    createServer: (options: Parameters<typeof real.createServer>[0]) => {
      const server = real.createServer(options)
      captured.server = server
      return { ...server, start: captured.start }
    },
  }
})
vi.mock("../src/auth", async (original) => ({
  ...(await original<typeof import("../src/auth")>()),
  initializeAuth: vi.fn(async () => Right(undefined)),
  getAccessToken: vi.fn(async () => Right("TOKEN")),
}))

const BASE = "http://localhost:3000"
let storeDirectory = ""
const startServer = async (mode: string): Promise<SomaServerInstance> => {
  vi.stubEnv("MS365_AUTH_MODE", mode)
  vi.stubEnv("TRANSPORT_TYPE", "httpStream")
  await import("../src/index")
  await vi.waitFor(() => expect(captured.server).toBeDefined())
  return captured.server as SomaServerInstance
}

const rpc = async (server: SomaServerInstance, method: string, params: unknown = {}): Promise<Response> =>
  server.fetch(
    new Request(`${BASE}/mcp`, {
      method: "POST",
      headers: {
        accept: "application/json, text/event-stream",
        "content-type": "application/json",
        "mcp-protocol-version": "2025-06-18",
      },
      body: JSON.stringify({ jsonrpc: "2.0", id: 1, method, params }),
    }),
  )

const rpcResult = async (response: Response) => {
  const text = await response.text()
  const json = text.startsWith("{") ? text : (text.split("\n").find((line) => line.startsWith("data: ")) ?? "").slice(6)
  return JSON.parse(json).result
}

beforeEach(() => {
  vi.resetModules()
  captured.server = undefined
  captured.azureConfig = undefined
  captured.start.mockClear()
  vi.stubEnv("MS365_CLIENT_ID", "00000000-0000-0000-0000-000000000001")
  vi.stubEnv("MS365_CLIENT_SECRET", "secret")
  vi.stubEnv("MS365_ACCESS_TOKEN", "a.b.c")
  vi.stubEnv("MS365_OAUTH_BASE_URL", BASE)
  vi.stubEnv("MS365_JWT_SIGNING_KEY", "signing-key")
  vi.stubEnv("MS365_TOKEN_ENCRYPTION_KEY", "encryption-key")
  storeDirectory = mkdtempSync(join(tmpdir(), "ms365-tokens-spec-"))
  vi.stubEnv("TOKEN_STORAGE_PATH", storeDirectory)
})
afterEach(() => {
  vi.unstubAllEnvs()
  rmSync(storeDirectory, { recursive: true, force: true })
})

describe.each(["client-token", "oauth-proxy"])("%s over HTTP", (mode) => {
  it("answers /ping before any sign-in, as the Docker HEALTHCHECK needs", async () => {
    const server = await startServer(mode)
    const response = await server.fetch(new Request(`${BASE}/ping`))
    expect(response.status).toBe(200)
    expect(await response.text()).toBe("pong")
  })

  it.each(["/health", "/health/detail", "/info", "/dashboard"])("does not serve SomaMCP's %s", async (path) => {
    const server = await startServer(mode)
    expect((await server.fetch(new Request(`${BASE}${path}`))).status).toBe(404)
  })

  it("mounts the relay routes outside the sign-in, for upload tickets", async () => {
    const server = await startServer(mode)
    const response = await server.fetch(new Request(`${BASE}/upload`, { method: "PUT", body: "x" }))
    expect(response.status).toBe(400)
    expect(await response.json()).toEqual({ error: "path query parameter is required" })
  })
})

describe("client-token over MCP", () => {
  it("answers a 2026-07-28 request, which FastMCP's stack refused with a 400", async () => {
    const server = await startServer("client-token")
    const response = await server.fetch(
      new Request(`${BASE}/mcp`, {
        method: "POST",
        headers: {
          accept: "application/json, text/event-stream",
          "content-type": "application/json",
          "mcp-method": "tools/list",
          "mcp-protocol-version": "2026-07-28",
        },
        body: JSON.stringify({
          jsonrpc: "2.0",
          id: 1,
          method: "tools/list",
          params: {
            _meta: {
              "io.modelcontextprotocol/clientCapabilities": {},
              "io.modelcontextprotocol/clientInfo": { name: "spec", version: "1" },
              "io.modelcontextprotocol/protocolVersion": "2026-07-28",
            },
          },
        }),
      }),
    )
    expect(response.status).toBe(200)
    const { tools } = await rpcResult(response)
    expect(tools.map((tool: { name: string }) => tool.name)).toContain("get_me")
  })

  it("keeps CORS on /mcp for browser-based clients", async () => {
    await startServer("client-token")
    await vi.waitFor(() => expect(captured.start).toHaveBeenCalled())
    expect(captured.start).toHaveBeenCalledWith(
      expect.objectContaining({ httpStream: expect.objectContaining({ cors: true, endpoint: "/mcp" }) }),
    )
  })

  it("adds no tool of its own to the list", async () => {
    const server = await startServer("client-token")
    const { tools } = await rpcResult(await rpc(server, "tools/list"))
    expect(tools.map((tool: { name: string }) => tool.name)).not.toContain("info")
    expect(tools.map((tool: { name: string }) => tool.name)).toContain("get_me")
  })

  it("reports a tool's refusal as its bare message, not a JSON envelope", async () => {
    const server = await startServer("client-token")
    const result = await rpcResult(
      await rpc(server, "tools/call", { name: "get_event", arguments: { event_id: "../../users/x" } }),
    )
    expect(result.isError).toBe(true)
    expect(result.content[0].text).toBe(
      "event_id is not an event ID. Use an ID exactly as list_events or list_calendar_view returns it.",
    )
  })

  it("takes a request over SomaMCP's 4 MiB default, up to the 10 MiB it always allowed", async () => {
    const server = await startServer("client-token")
    const call = (bytes: number) =>
      rpc(server, "tools/call", { name: "get_event", arguments: { event_id: "x".repeat(bytes) } })
    expect((await call(5 * 1024 * 1024)).status).toBe(200)
    expect((await call(11 * 1024 * 1024)).status).toBe(413)
  })
})

describe("oauth-proxy sign-in", () => {
  it("answers an unsigned MCP request with the challenge a client needs to start signing in", async () => {
    const server = await startServer("oauth-proxy")
    const response = await rpc(server, "tools/list")
    expect(response.status).toBe(401)
    expect(response.headers.get("www-authenticate")).toContain(
      `resource_metadata="${BASE}/.well-known/oauth-protected-resource/mcp"`,
    )
  })

  it("publishes the OAuth metadata with the server's scopes", async () => {
    const server = await startServer("oauth-proxy")
    const response = await server.fetch(new Request(`${BASE}/.well-known/oauth-authorization-server`))
    const metadata = await response.json()
    expect(metadata.token_endpoint).toBe(`${BASE}/oauth/token`)
    expect(metadata.scopes_supported).toContain("Contacts.ReadWrite")
  })

  // With the store directory, these are what lets sign-ins made under FastMCP survive the move.
  it("keeps the base URL and both keys as configured", async () => {
    await startServer("oauth-proxy")
    expect(captured.azureConfig).toMatchObject({
      baseUrl: BASE,
      jwtSigningKey: "signing-key",
      encryptionKey: "encryption-key",
    })
  })

  it.each([
    [undefined, false],
    ["true", true],
    ["TRUE", true],
    ["false", false],
  ])("widens scopes on refresh only when asked: %s", async (value, expected) => {
    vi.stubEnv("MS365_WIDEN_SCOPES_ON_REFRESH", value as string)
    await startServer("oauth-proxy")
    expect(captured.azureConfig?.widenScopesOnRefresh).toBe(expected)
  })
})
