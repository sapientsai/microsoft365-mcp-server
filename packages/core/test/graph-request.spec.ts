import { Left, Right } from "functype/either"
import { afterEach, describe, expect, it, vi } from "vitest"

import type { AuthStrategy } from "../src/auth-strategy"
import { createGraphRequest } from "../src/graph-request"

const auth = (token = "TOK"): AuthStrategy => ({ getAccessToken: () => Promise.resolve(Right(token)) })

const res = (body: unknown, init: { ok?: boolean; status?: number; headers?: Headers } = {}) =>
  ({
    ok: init.ok ?? true,
    status: init.status ?? 200,
    headers: init.headers ?? new Headers(),
    text: () => Promise.resolve(typeof body === "string" ? body : JSON.stringify(body)),
    json: () => Promise.resolve(body),
  }) as Response

describe("createGraphRequest", () => {
  afterEach(() => vi.unstubAllGlobals())

  it("builds the URL with the default version and sends the injected Bearer", async () => {
    const fetchSpy = vi.fn(() => Promise.resolve(res({ id: "1" })))
    vi.stubGlobal("fetch", fetchSpy)
    const { request } = createGraphRequest(auth("INJECTED"))

    const result = await request<{ id: string }>("GET", "/me")
    expect(result.isRight()).toBe(true)
    const [url, init] = fetchSpy.mock.calls[0] as [string, RequestInit]
    expect(url).toBe("https://graph.microsoft.com/v1.0/me")
    expect((init.headers as Record<string, string>).Authorization).toBe("Bearer INJECTED")
  })

  it("honors a custom defaultVersion resolver", async () => {
    const fetchSpy = vi.fn(() => Promise.resolve(res({})))
    vi.stubGlobal("fetch", fetchSpy)
    const { request } = createGraphRequest(auth(), { defaultVersion: () => "beta" })

    await request("GET", "/me")
    expect((fetchSpy.mock.calls[0] as [string])[0]).toBe("https://graph.microsoft.com/beta/me")
  })

  it("serializes a JSON body on POST", async () => {
    const fetchSpy = vi.fn(() => Promise.resolve(res({})))
    vi.stubGlobal("fetch", fetchSpy)
    const { request } = createGraphRequest(auth())

    await request("POST", "/me/sendMail", { body: { message: { subject: "hi" } } })
    const init = (fetchSpy.mock.calls[0] as [string, RequestInit])[1]
    expect(init.body).toBe(JSON.stringify({ message: { subject: "hi" } }))
  })

  it("short-circuits with an auth error when the strategy fails (no fetch)", async () => {
    const fetchSpy = vi.fn()
    vi.stubGlobal("fetch", fetchSpy)
    const failing: AuthStrategy = {
      getAccessToken: () => Promise.resolve(Left({ type: "token", message: "no token" })),
    }

    const result = await createGraphRequest(failing).request("GET", "/me")
    expect(result.isLeft()).toBe(true)
    expect((result.value as { type: string }).type).toBe("auth")
    expect(fetchSpy).not.toHaveBeenCalled()
  })

  it("maps a 429 to a throttle error with Retry-After", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn(() =>
        Promise.resolve(
          res(
            { error: { message: "slow down", code: "TooManyRequests" } },
            {
              ok: false,
              status: 429,
              headers: new Headers({ "Retry-After": "30" }),
            },
          ),
        ),
      ),
    )
    const result = await createGraphRequest(auth()).request("GET", "/me")
    expect(result.isLeft()).toBe(true)
    const err = result.value as { type: string; status: number; retryAfter?: number }
    expect(err.type).toBe("throttle")
    expect(err.retryAfter).toBe(30)
  })

  it("returns an empty object on 204 No Content", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn(() => Promise.resolve(res("", { status: 204 }))),
    )
    const result = await createGraphRequest(auth()).request<Record<string, never>>("DELETE", "/me/events/1")
    expect(result.isRight()).toBe(true)
  })
})
