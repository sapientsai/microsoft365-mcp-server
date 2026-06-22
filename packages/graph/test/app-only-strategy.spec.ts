import { describe, expect, it, vi } from "vitest"

import { createAppOnlyAuthStrategy } from "../src/auth/app-only-strategy"
import type { AppOnlyConfig } from "../src/config"

const config: AppOnlyConfig = {
  tenantId: "tenant-123",
  clientId: "client-123",
  clientSecret: "secret",
  appScopes: ["https://graph.microsoft.com/.default"],
}

const okResponse = (accessToken: string, expiresIn = 3600) =>
  ({
    ok: true,
    status: 200,
    json: () => Promise.resolve({ access_token: accessToken, expires_in: expiresIn }),
  }) as Response

describe("createAppOnlyAuthStrategy", () => {
  it("acquires a token via client_credentials and posts the right body", async () => {
    const fetchImpl = vi.fn(() => Promise.resolve(okResponse("TOK")))
    const auth = createAppOnlyAuthStrategy(config, fetchImpl as unknown as typeof fetch)

    const result = await auth.getAccessToken()
    expect(result.isRight()).toBe(true)
    expect(result.value).toBe("TOK")

    const [url, init] = fetchImpl.mock.calls[0] as [string, RequestInit]
    expect(url).toBe("https://login.microsoftonline.com/tenant-123/oauth2/v2.0/token")
    const body = String(init.body)
    expect(body).toContain("grant_type=client_credentials")
    expect(body).toContain("client_id=client-123")
    expect(body).toContain("scope=https%3A%2F%2Fgraph.microsoft.com%2F.default")
  })

  it("caches the token and does not refetch within the validity window", async () => {
    const fetchImpl = vi.fn(() => Promise.resolve(okResponse("TOK")))
    const auth = createAppOnlyAuthStrategy(config, fetchImpl as unknown as typeof fetch)

    await auth.getAccessToken()
    await auth.getAccessToken()
    expect(fetchImpl).toHaveBeenCalledOnce()
  })

  it("refetches once the token is within the 5-min refresh buffer", async () => {
    let clock = 1_000_000
    const fetchImpl = vi
      .fn()
      .mockResolvedValueOnce(okResponse("TOK1", 240)) // expires in 4 min → inside the 5-min buffer
      .mockResolvedValueOnce(okResponse("TOK2", 3600))
    const auth = createAppOnlyAuthStrategy(config, fetchImpl as unknown as typeof fetch, () => clock)

    expect((await auth.getAccessToken()).value).toBe("TOK1")
    clock += 1000
    expect((await auth.getAccessToken()).value).toBe("TOK2")
    expect(fetchImpl).toHaveBeenCalledTimes(2)
  })

  it("returns a credential error on a non-OK token response", async () => {
    const fetchImpl = vi.fn(() =>
      Promise.resolve({
        ok: false,
        status: 401,
        text: () => Promise.resolve(JSON.stringify({ error_description: "invalid client secret" })),
      } as Response),
    )
    const auth = createAppOnlyAuthStrategy(config, fetchImpl as unknown as typeof fetch)

    const result = await auth.getAccessToken()
    expect(result.isLeft()).toBe(true)
    const err = result.value as { type: string; message: string }
    expect(err.type).toBe("credential")
    expect(err.message).toContain("invalid client secret")
  })

  it("returns a credential error when fetch itself throws", async () => {
    const fetchImpl = vi.fn(() => Promise.reject(new Error("ENOTFOUND login.microsoftonline.com")))
    const auth = createAppOnlyAuthStrategy(config, fetchImpl as unknown as typeof fetch)

    const result = await auth.getAccessToken()
    expect(result.isLeft()).toBe(true)
    expect((result.value as { message: string }).message).toContain("ENOTFOUND")
  })
})
