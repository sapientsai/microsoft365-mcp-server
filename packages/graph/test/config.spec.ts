import { describe, expect, it } from "vitest"

import { resolveAppOnlyConfig, resolveServerRuntimeConfig } from "../src/config"

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
})
