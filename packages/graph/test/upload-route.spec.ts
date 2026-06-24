import { type AuthStrategy, mintUploadTicket } from "@sapientsai/ms-graph-core"
import { Left, Right } from "functype/either"
import { afterEach, describe, expect, it, vi } from "vitest"

import { handleUpload, type UploadRequestContext } from "../src/upload/upload-route"

const auth = (token = "APP_ONLY_TOKEN"): AuthStrategy => ({ getAccessToken: () => Promise.resolve(Right(token)) })

const req = (opts: { query?: Record<string, string>; bearer?: string; body?: Buffer }): UploadRequestContext => ({
  query: (name: string) => opts.query?.[name],
  header: (name: string) =>
    name.toLowerCase() === "authorization" && opts.bearer ? `Bearer ${opts.bearer}` : undefined,
  arrayBuffer: () => Promise.resolve((opts.body ?? Buffer.from("")).buffer),
})

const contentPath = "/me/drive/root:/Documents/x.txt:/content"

describe("handleUpload caller authorization", () => {
  afterEach(() => vi.unstubAllGlobals())

  it("refuses with 503 when no MCP_API_KEY is configured", async () => {
    const res = await handleUpload(
      req({ query: { path: contentPath }, bearer: "anything", body: Buffer.from("x") }),
      auth(),
      undefined,
    )
    expect(res.status).toBe(503)
  })

  it("rejects a wrong key with 401", async () => {
    const res = await handleUpload(
      req({ query: { path: contentPath }, bearer: "wrong", body: Buffer.from("x") }),
      auth(),
      "SECRET",
    )
    expect(res.status).toBe(401)
  })

  it("requires a path ending in :/content", async () => {
    const res = await handleUpload(
      req({ query: { path: "/me/drive/root:/x.txt" }, bearer: "SECRET" }),
      auth(),
      "SECRET",
    )
    expect(res.status).toBe(400)
  })

  it("accepts the raw api key and uploads with the server's app-only token", async () => {
    const fetchSpy = vi.fn(() =>
      Promise.resolve({ ok: true, json: () => Promise.resolve({ id: "drive-item-1", name: "x.txt" }) } as Response),
    )
    vi.stubGlobal("fetch", fetchSpy)

    const res = await handleUpload(
      req({ query: { path: contentPath }, bearer: "SECRET", body: Buffer.from("hello") }),
      auth("APP_ONLY_TOKEN"),
      "SECRET",
    )

    expect(res.status).toBe(200)
    expect((res.body as { id: string }).id).toBe("drive-item-1")
    // simpleUpload used the app-only token, not the caller's key
    const [, init] = fetchSpy.mock.calls[0] as [string, RequestInit]
    expect((init.headers as Record<string, string>).Authorization).toBe("Bearer APP_ONLY_TOKEN")
  })

  it("accepts an opaque ticket that resolves to the api key", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn(() => Promise.resolve({ ok: true, json: () => Promise.resolve({ id: "1" }) } as Response)),
    )
    const ticket = mintUploadTicket("SECRET")
    const res = await handleUpload(
      req({ query: { path: contentPath }, bearer: ticket, body: Buffer.from("hi") }),
      auth(),
      "SECRET",
    )
    expect(res.status).toBe(200)
  })

  it("returns 401 when the app-only token cannot be acquired", async () => {
    const failing: AuthStrategy = {
      getAccessToken: () => Promise.resolve(Left({ type: "credential", message: "bad secret" })),
    }
    const res = await handleUpload(
      req({ query: { path: contentPath }, bearer: "SECRET", body: Buffer.from("x") }),
      failing,
      "SECRET",
    )
    expect(res.status).toBe(401)
    expect((res.body as { error: string }).error).toContain("bad secret")
  })

  it("rejects an empty body with 400", async () => {
    const res = await handleUpload(
      req({ query: { path: contentPath }, bearer: "SECRET", body: Buffer.from("") }),
      auth(),
      "SECRET",
    )
    expect(res.status).toBe(400)
  })
})
