import type { AuthStrategy } from "@sapientsai/ms-graph-core"
import { Left, Right } from "functype/either"
import { afterEach, describe, expect, it, vi } from "vitest"

import { buildUploadRoute, handleUpload, type UploadRequestContext } from "../src/upload/upload-route"

const auth = (token = "APP_ONLY_TOKEN"): AuthStrategy => ({ getAccessToken: () => Promise.resolve(Right(token)) })

const req = (opts: { query?: Record<string, string>; body?: Buffer }): UploadRequestContext => ({
  query: (name: string) => opts.query?.[name],
  header: () => undefined,
  arrayBuffer: () => Promise.resolve((opts.body ?? Buffer.from("")).buffer),
})

const contentPath = "/me/drive/root:/Documents/x.txt:/content"

// handleUpload is upload mechanics only — caller authorization now lives in the somamcp
// `protected` gate (see api-key-gate.spec + the onUnauthorized tests below).
describe("handleUpload mechanics", () => {
  afterEach(() => vi.unstubAllGlobals())

  it("requires a path ending in :/content", async () => {
    const res = await handleUpload(req({ query: { path: "/me/drive/root:/x.txt" } }), auth())
    expect(res.status).toBe(400)
  })

  it("uploads with the server's app-only token", async () => {
    const fetchSpy = vi.fn(() =>
      Promise.resolve({ ok: true, json: () => Promise.resolve({ id: "drive-item-1", name: "x.txt" }) } as Response),
    )
    vi.stubGlobal("fetch", fetchSpy)

    const res = await handleUpload(
      req({ query: { path: contentPath }, body: Buffer.from("hello") }),
      auth("APP_ONLY_TOKEN"),
    )

    expect(res.status).toBe(200)
    expect((res.body as { id: string }).id).toBe("drive-item-1")
    const [, init] = fetchSpy.mock.calls[0] as [string, RequestInit]
    expect((init.headers as Record<string, string>).Authorization).toBe("Bearer APP_ONLY_TOKEN")
  })

  it("returns 401 when the app-only token cannot be acquired", async () => {
    const failing: AuthStrategy = {
      getAccessToken: () => Promise.resolve(Left({ type: "credential", message: "bad secret" })),
    }
    const res = await handleUpload(req({ query: { path: contentPath }, body: Buffer.from("x") }), failing)
    expect(res.status).toBe(401)
    expect((res.body as { error: string }).error).toContain("bad secret")
  })

  it("rejects an empty body with 400", async () => {
    const res = await handleUpload(req({ query: { path: contentPath }, body: Buffer.from("") }), auth())
    expect(res.status).toBe(400)
  })
})

describe("buildUploadRoute", () => {
  it("registers a protected POST/PUT /upload route", () => {
    const route = buildUploadRoute(auth(), "SECRET")
    expect(route.path).toBe("/upload")
    expect(route.protected).toBe(true)
    expect(route.method).toEqual(["POST", "PUT"])
  })

  it("onUnauthorized returns 401 when an api key IS configured", async () => {
    const route = buildUploadRoute(auth(), "SECRET")
    const res = (await route.onUnauthorized?.({} as never)) as Response
    expect(res.status).toBe(401)
    expect(await res.json()).toEqual({ error: "Unauthorized" })
  })

  it("onUnauthorized returns 503 when NO api key is configured", async () => {
    const route = buildUploadRoute(auth(), undefined)
    const res = (await route.onUnauthorized?.({} as never)) as Response
    expect(res.status).toBe(503)
    expect((await res.json()).error).toContain("MCP_API_KEY")
  })

  it("handler maps a handleUpload result to a JSON Response", async () => {
    const route = buildUploadRoute(auth(), "SECRET")
    // Bad path → 400 without needing a token or network.
    const c = {
      req: {
        header: () => undefined,
        query: () => "/me/drive/root:/x.txt",
        arrayBuffer: () => Promise.resolve(Buffer.from("x").buffer),
      },
    }
    const res = (await route.handler(c as never)) as Response
    expect(res.status).toBe(400)
    expect(res.headers.get("content-type")).toContain("application/json")
  })
})
