import { describe, expect, it } from "vitest"

import { resolveAppOnlyConfig, resolvePublicBaseUrl, resolveServerRuntimeConfig } from "../src/config"

const base = {
  MS_GRAPH_TENANT_ID: "11111111-2222-3333-4444-555555555555",
  MS_GRAPH_CLIENT_ID: "app-client-id",
  MS_GRAPH_CLIENT_SECRET: "shhh",
} as NodeJS.ProcessEnv

describe("resolveAppOnlyConfig", () => {
  it("resolves a valid app-only config with the default scope", () => {
    const result = resolveAppOnlyConfig(base)
    expect(result.isRight()).toBe(true)
    const cfg = result.value as { appScopes: string[]; tenantId: string }
    expect(cfg.tenantId).toBe(base.MS_GRAPH_TENANT_ID)
    expect(cfg.appScopes).toEqual(["https://graph.microsoft.com/.default"])
  })

  it("parses comma-separated custom scopes", () => {
    const result = resolveAppOnlyConfig({ ...base, MS_GRAPH_APP_SCOPES: "a/.default, b/.default" })
    expect((result.value as { appScopes: string[] }).appScopes).toEqual(["a/.default", "b/.default"])
  })

  it.each(["common", "organizations", "consumers", "COMMON"])("rejects the multi-tenant alias %s", (tenant) => {
    const result = resolveAppOnlyConfig({ ...base, MS_GRAPH_TENANT_ID: tenant })
    expect(result.isLeft()).toBe(true)
    expect(result.value as string).toContain("concrete tenant")
  })

  it("requires tenant, client id, and secret", () => {
    expect(resolveAppOnlyConfig({ ...base, MS_GRAPH_TENANT_ID: "" }).isLeft()).toBe(true)
    expect(resolveAppOnlyConfig({ ...base, MS_GRAPH_CLIENT_ID: "" }).isLeft()).toBe(true)
    expect(resolveAppOnlyConfig({ ...base, MS_GRAPH_CLIENT_SECRET: "" }).isLeft()).toBe(true)
  })
})

describe("resolveServerRuntimeConfig", () => {
  it("defaults to httpStream on 127.0.0.1:8080 with no api key", () => {
    const cfg = resolveServerRuntimeConfig(base).value as {
      transport: string
      port: number
      host: string
      apiKey?: string
    }
    expect(cfg.transport).toBe("httpStream")
    expect(cfg.port).toBe(8080)
    expect(cfg.host).toBe("127.0.0.1")
    expect(cfg.apiKey).toBeUndefined()
  })

  it("honors stdio, PORT, HOST, and MCP_API_KEY", () => {
    const cfg = resolveServerRuntimeConfig({
      ...base,
      TRANSPORT_TYPE: "stdio",
      PORT: "9090",
      HOST: "0.0.0.0",
      MCP_API_KEY: "key",
    }).value as { transport: string; port: number; host: string; apiKey?: string }
    expect(cfg.transport).toBe("stdio")
    expect(cfg.port).toBe(9090)
    expect(cfg.host).toBe("0.0.0.0")
    expect(cfg.apiKey).toBe("key")
  })

  it("reads FASTMCP_HOST, the older name, only when HOST is unset", () => {
    const host = (env: Record<string, string>) =>
      (resolveServerRuntimeConfig({ ...base, ...env }).value as { host: string }).host
    expect(host({ FASTMCP_HOST: "0.0.0.0" })).toBe("0.0.0.0")
    expect(host({ HOST: "127.0.0.1", FASTMCP_HOST: "0.0.0.0" })).toBe("127.0.0.1")
  })
})

// get_upload_config hands out a curl to this URL. On a wildcard bind (the Docker image's 0.0.0.0)
// the bind address is reachable from nowhere outside the container, so there is no default.
describe("resolvePublicBaseUrl", () => {
  it("uses MCP_PUBLIC_BASE_URL, without a trailing slash, whatever the bind address", () => {
    expect(resolvePublicBaseUrl({ MCP_PUBLIC_BASE_URL: "https://graph.example.com/" }, "0.0.0.0", 8080)).toBe(
      "https://graph.example.com",
    )
  })

  it("falls back to a real bind address, as for a local run", () => {
    expect(resolvePublicBaseUrl({}, "127.0.0.1", 8080)).toBe("http://127.0.0.1:8080")
  })

  it.each(["0.0.0.0", "::", "[::]", "", " ", "0", "::0", "0:0:0:0:0:0:0:0", " 0.0.0.0 "])(
    "has no default for the wildcard bind %j",
    (host) => {
      expect(resolvePublicBaseUrl({}, host, 8080)).toBeUndefined()
      expect(resolvePublicBaseUrl({ MCP_PUBLIC_BASE_URL: "  " }, host, 8080)).toBeUndefined()
    },
  )

  it.each([
    ["::1", "http://[::1]:8080"],
    ["[::1]", "http://[::1]:8080"],
    ["localhost", "http://localhost:8080"],
    ["graph.internal", "http://graph.internal:8080"],
    ["10.0.0.5", "http://10.0.0.5:8080"],
  ])("builds a valid URL for the bind address %s", (host, url) => {
    expect(resolvePublicBaseUrl({}, host, 8080)).toBe(url)
    expect(URL.canParse(url)).toBe(true)
  })

  it("reaches the runtime config", () => {
    const url = (env: Record<string, string>) =>
      (resolveServerRuntimeConfig({ ...base, ...env }).value as { publicBaseUrl?: string }).publicBaseUrl
    expect(url({ HOST: "0.0.0.0" })).toBeUndefined()
    expect(url({ HOST: "0.0.0.0", MCP_PUBLIC_BASE_URL: "https://graph.example.com" })).toBe("https://graph.example.com")
  })
})
