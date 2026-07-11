import type { AuthStrategy } from "@sapientsai/ms-graph-core"
import { Left, Right } from "functype/either"
import { afterEach, describe, expect, it, vi } from "vitest"

import { getGraphClient, initializeGraphClient } from "../src/client/graph-client"

// Phase 2b: graph-client no longer imports the server auth module — it receives an
// AuthStrategy. These tests lock that seam.
describe("graph-client AuthStrategy injection", () => {
  afterEach(() => vi.unstubAllGlobals())

  const stubFetch = (json: unknown, ok = true, status = 200) =>
    vi.stubGlobal(
      "fetch",
      vi.fn(() =>
        Promise.resolve({
          ok,
          status,
          text: () => Promise.resolve(JSON.stringify(json)),
          json: () => Promise.resolve(json),
          headers: new Headers(),
        } as Response),
      ),
    )

  it("uses the injected strategy's token as the Bearer credential", async () => {
    const getAccessToken = vi.fn(() => Promise.resolve(Right("INJECTED.TOKEN")))
    const auth: AuthStrategy = { getAccessToken }
    stubFetch({ value: [] })

    const client = initializeGraphClient(auth)
    await client.listMessages()

    expect(getAccessToken).toHaveBeenCalledOnce()
    const [, init] = vi.mocked(fetch).mock.calls[0] as [string, RequestInit]
    expect((init.headers as Record<string, string>).Authorization).toBe("Bearer INJECTED.TOKEN")
  })

  it("short-circuits to an auth error when the strategy fails (fetch never called)", async () => {
    const auth: AuthStrategy = { getAccessToken: () => Promise.resolve(Left({ type: "token", message: "no token" })) }
    const fetchSpy = vi.fn()
    vi.stubGlobal("fetch", fetchSpy)

    const client = initializeGraphClient(auth)
    const result = await client.getMe()

    expect(result.isLeft()).toBe(true)
    expect((result.value as { type: string }).type).toBe("auth")
    expect(fetchSpy).not.toHaveBeenCalled()
  })

  it("initializeGraphClient registers the client as the active singleton", () => {
    const auth: AuthStrategy = { getAccessToken: () => Promise.resolve(Right("t")) }
    initializeGraphClient(auth)
    expect(getGraphClient().isNone()).toBe(false)
  })

  it("graphQuery forwards caller-supplied headers (e.g. If-Match) onto the request", async () => {
    const auth: AuthStrategy = { getAccessToken: () => Promise.resolve(Right("t")) }
    stubFetch({ ok: true })

    const client = initializeGraphClient(auth)
    await client.graphQuery("PATCH", "/planner/tasks/abc/details", { description: "x" }, undefined, {
      "If-Match": 'W/"etag123"',
    })

    const [, init] = vi.mocked(fetch).mock.calls[0] as [string, RequestInit]
    expect((init.headers as Record<string, string>)["If-Match"]).toBe('W/"etag123"')
  })

  it("updatePlannerTaskDetails sends the details path with the If-Match ETag", async () => {
    const auth: AuthStrategy = { getAccessToken: () => Promise.resolve(Right("t")) }
    stubFetch({ ok: true })

    const client = initializeGraphClient(auth)
    await client.updatePlannerTaskDetails("abc", { description: "hi" }, 'W/"e"')

    const [url, init] = vi.mocked(fetch).mock.calls[0] as [string, RequestInit]
    expect(url).toContain("/planner/tasks/abc/details")
    expect(init.method).toBe("PATCH")
    expect((init.headers as Record<string, string>)["If-Match"]).toBe('W/"e"')
  })
})
