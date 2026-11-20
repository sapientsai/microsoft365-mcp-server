// A write that fails on a 503, 504, timeout or network error isn't retried, because it may already
// have run, so its error says "The change may have gone through". The note belongs exactly where
// that is true: not on reads, not when nothing was sent, and not when no change can exist yet.
import { Right } from "functype/either"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

import type { AuthStrategy } from "../src/auth-strategy"
import { createGraphRequest } from "../src/graph-request"
import { getSharedBreaker, MAY_HAVE_LANDED, resetSharedBreakerForTests } from "../src/resilience"
import { attachFile, requireDraftMessage } from "../src/upload/attachment-upload"
import { uploadDriveContent } from "../src/upload/upload"

const apiBase = "https://graph.microsoft.com/v1.0"
const path = "/me/drive/root:/a.txt:/content"
const status = (code: number) => () => Promise.resolve(new Response("", { status: code }))
const json = (value: unknown, code = 200) => new Response(JSON.stringify(value), { status: code })
const lost = () => Promise.reject(new TypeError("fetch failed"))
const messageOf = (result: { value: unknown }) => (result.value as { message: string }).message
const auth: AuthStrategy = { getAccessToken: () => Promise.resolve(Right("TOKEN")) }

beforeEach(() => {
  vi.stubEnv("MS365_GRAPH_MAX_RETRIES", "2")
  vi.stubEnv("MS365_GRAPH_BASE_BACKOFF_MS", "0")
  vi.stubEnv("MS365_GRAPH_CIRCUIT_DISABLED", "true")
  resetSharedBreakerForTests()
})
afterEach(() => {
  vi.unstubAllGlobals()
  vi.unstubAllEnvs()
  resetSharedBreakerForTests()
})

describe("a simple upload", () => {
  it.each([
    ["a 503", status(503)],
    ["a network error", lost],
  ])("that got %s says it may have gone through", async (_, failure) => {
    vi.stubGlobal("fetch", vi.fn(failure))
    const result = await uploadDriveContent(apiBase, path, "t", Buffer.from("x"), "text/plain", "rename")
    expect(messageOf(result)).toContain(MAY_HAVE_LANDED)
  })

  it("that ran out of 429 retries doesn't: nothing ran", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn(() => Promise.resolve(new Response("", { status: 429, headers: { "Retry-After": "0" } }))),
    )
    const result = await uploadDriveContent(apiBase, path, "t", Buffer.from("x"), "text/plain", "rename")
    expect(result.isLeft()).toBe(true)
    expect(messageOf(result)).not.toContain(MAY_HAVE_LANDED)
  })

  it("refused by an open breaker doesn't: nothing was sent", async () => {
    vi.stubEnv("MS365_GRAPH_CIRCUIT_DISABLED", "false")
    vi.stubEnv("MS365_GRAPH_CIRCUIT_THRESHOLD", "1")
    resetSharedBreakerForTests()
    getSharedBreaker().recordFailure()
    const fetchSpy = vi.fn()
    vi.stubGlobal("fetch", fetchSpy)
    const result = await uploadDriveContent(apiBase, path, "t", Buffer.from("x"), "text/plain", "rename")
    expect(fetchSpy).not.toHaveBeenCalled()
    expect(messageOf(result)).not.toContain(MAY_HAVE_LANDED)
  })
})

describe("a session upload", () => {
  const session = () => Promise.resolve(json({ uploadUrl: "https://upload.example/session" }))

  it("whose session wasn't created doesn't say it: no file exists", async () => {
    vi.stubGlobal("fetch", vi.fn(status(503)))
    const result = await uploadDriveContent(apiBase, path, "t", Buffer.from("x"), "text/plain", "replace", {
      ifMatch: '"{A},1"',
    })
    expect(messageOf(result)).not.toContain(MAY_HAVE_LANDED)
  })

  it("whose last chunk got a 503 says it: the file may be there", async () => {
    vi.stubGlobal("fetch", vi.fn().mockImplementationOnce(session).mockImplementation(status(503)))
    const result = await uploadDriveContent(apiBase, path, "t", Buffer.from("x"), "text/plain", "replace", {
      ifMatch: '"{A},1"',
    })
    expect(messageOf(result)).toContain(MAY_HAVE_LANDED)
  })

  it("whose middle chunk failed doesn't say it: the file appears only with the last chunk", async () => {
    const big = Buffer.alloc(11 * 1024 * 1024, 1)
    for (const failure of [status(503), lost]) {
      vi.stubGlobal("fetch", vi.fn().mockImplementationOnce(session).mockImplementation(failure))
      const result = await uploadDriveContent(apiBase, path, "t", big, "application/octet-stream", "rename")
      expect(messageOf(result)).toContain("at byte 0")
      expect(messageOf(result)).not.toContain(MAY_HAVE_LANDED)
    }
  })
})

describe("attachments", () => {
  it("an inline attach that got a network error says it may have gone through", async () => {
    vi.stubGlobal("fetch", vi.fn(lost))
    const result = await attachFile(apiBase, "m1", "t", Buffer.from("x"), "a.txt", "text/plain")
    expect(messageOf(result)).toContain(MAY_HAVE_LANDED)
  })

  it("the draft check is a read, and doesn't say it", async () => {
    vi.stubGlobal("fetch", vi.fn(status(503)))
    const result = await requireDraftMessage(apiBase, "m1", "t")
    expect(messageOf(result)).not.toContain(MAY_HAVE_LANDED)
  })
})

describe("the request layer", () => {
  it("a POST that only reads doesn't say it", async () => {
    vi.stubGlobal("fetch", vi.fn(status(503)))
    const result = await createGraphRequest(auth).request("POST", "/me/findMeetingTimes", { body: {}, readOnly: true })
    expect(result.isLeft()).toBe(true)
    expect(messageOf(result)).not.toContain(MAY_HAVE_LANDED)
  })

  it("a write that got a 503 says it once", async () => {
    vi.stubGlobal("fetch", vi.fn(status(503)))
    const result = await createGraphRequest(auth).request("POST", "/me/events", { body: {} })
    expect(messageOf(result).split(MAY_HAVE_LANDED)).toHaveLength(2)
  })

  it("a body that can't be sent as JSON is refused before anything is sent", async () => {
    const fetchSpy = vi.fn()
    vi.stubGlobal("fetch", fetchSpy)
    const circular: Record<string, unknown> = {}
    circular.self = circular
    const result = await createGraphRequest(auth).request("POST", "/me/events", { body: circular })
    expect(result.value).toMatchObject({ type: "api", status: 400 })
    expect(messageOf(result)).not.toContain(MAY_HAVE_LANDED)
    expect(fetchSpy).not.toHaveBeenCalled()
  })
})
