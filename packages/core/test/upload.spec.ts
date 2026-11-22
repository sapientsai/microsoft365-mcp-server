import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

import { describeFetchError, sessionUpload, simpleUpload } from "../src/upload/upload"
import { isUploadTicket, mintUploadTicket, redeemUploadTicket } from "../src/upload/upload-ticket"

describe("upload-ticket", () => {
  const drive = { relay: "drive", path: "/me/drive/root:/a.txt:/content" } as const
  const draft = { relay: "attachment", messageId: "AAMkDraft", mailbox: "shared@example.com" } as const

  it("mints an opaque prefixed ticket that redeems to the token for its upload", () => {
    const ticket = mintUploadTicket("DELEGATED.JWT.VALUE", drive)
    expect(isUploadTicket(ticket)).toBe(true)
    expect(ticket).not.toContain("DELEGATED.JWT.VALUE")
    expect(redeemUploadTicket(ticket, { relay: "drive", path: "/me/drive/root:/a.txt:/content" })).toBe(
      "DELEGATED.JWT.VALUE",
    )
  })

  it("is good once: a second use is refused", () => {
    const ticket = mintUploadTicket("tok", drive)
    expect(redeemUploadTicket(ticket, drive)).toBe("tok")
    expect(redeemUploadTicket(ticket, drive)).toBeUndefined()
  })

  it.each([
    ["another path", { relay: "drive", path: "/me/drive/root:/other.txt:/content" } as const],
    ["the attachment relay", { relay: "attachment", messageId: "AAMkDraft" } as const],
  ])("refuses a drive ticket presented for %s, and spends it", (_, other) => {
    const ticket = mintUploadTicket("tok", drive)
    expect(redeemUploadTicket(ticket, other)).toBeUndefined()
    expect(redeemUploadTicket(ticket, drive)).toBeUndefined()
  })

  it.each([
    ["another draft", { relay: "attachment", messageId: "AAMkOther", mailbox: "shared@example.com" } as const],
    ["another mailbox", { relay: "attachment", messageId: "AAMkDraft", mailbox: "other@example.com" } as const],
    ["no mailbox", { relay: "attachment", messageId: "AAMkDraft" } as const],
    ["the drive relay", drive],
  ])("refuses an attachment ticket presented for %s", (_, other) => {
    expect(redeemUploadTicket(mintUploadTicket("tok", draft), other)).toBeUndefined()
  })

  it("returns undefined for an unknown ticket", () => {
    expect(redeemUploadTicket("upl_does-not-exist", drive)).toBeUndefined()
  })

  it("expires tickets past their TTL", () => {
    const t0 = 1_000_000
    expect(redeemUploadTicket(mintUploadTicket("tok", drive, 1000, t0), drive, t0 + 500)).toBe("tok")
    const late = mintUploadTicket("tok", drive, 1000, t0)
    expect(redeemUploadTicket(late, drive, t0 + 1001)).toBeUndefined()
    // once presented it is gone, so even an earlier clock can't redeem it
    expect(redeemUploadTicket(late, drive, t0 + 500)).toBeUndefined()
  })

  it("recognizes non-tickets", () => {
    expect(isUploadTicket("eyJraWQ.real.jwt")).toBe(false)
  })
})

describe("describeFetchError", () => {
  it("surfaces an Error cause and its code", () => {
    const err = new TypeError("fetch failed")
    ;(err as { cause?: unknown }).cause = Object.assign(new Error("getaddrinfo ENOTFOUND graph.microsoft.com"), {
      code: "ENOTFOUND",
    })
    const { message, code } = describeFetchError(err)
    expect(code).toBe("ENOTFOUND")
    expect(message).toContain("fetch failed")
    expect(message).toContain("ENOTFOUND")
    expect(message).toContain("graph.microsoft.com")
  })

  it("handles a non-Error cause", () => {
    const err = new TypeError("fetch failed")
    ;(err as { cause?: unknown }).cause = "boom"
    expect(describeFetchError(err).message).toBe("fetch failed (cause: boom)")
  })

  it("handles an error with no cause", () => {
    expect(describeFetchError(new Error("plain")).message).toBe("plain")
  })

  it("handles non-Error throwables", () => {
    expect(describeFetchError("nope").message).toBe("nope")
  })
})

describe("simpleUpload", () => {
  afterEach(() => vi.unstubAllGlobals())

  it("surfaces the fetch cause instead of an opaque 'fetch failed'", async () => {
    const thrown = new TypeError("fetch failed")
    ;(thrown as { cause?: unknown }).cause = Object.assign(new Error("connect ECONNREFUSED 20.190.1.1:443"), {
      code: "ECONNREFUSED",
    })
    vi.stubGlobal(
      "fetch",
      vi.fn(() => Promise.reject(thrown)),
    )

    const result = await simpleUpload(
      "https://graph.microsoft.com/v1.0",
      "/me/drive/root:/x.bin:/content",
      "tok",
      Buffer.from("data"),
      "application/octet-stream",
      "rename",
    )

    expect(result.isLeft()).toBe(true)
    const err = result.value as { type: string; message: string }
    expect(err.type).toBe("network")
    expect(err.message).toContain("ECONNREFUSED")
    expect(err.message).not.toBe("Network error during upload: fetch failed")
  })
})

describe("upload resilience environment controls", () => {
  const apiBase = "https://graph.microsoft.com/beta"
  const path = "/me/drive/root:/x.bin:/content"
  const data = Buffer.from("data")
  const item = { id: "uploaded", name: "x.bin" }
  const json = (value: unknown, status = 200) => new Response(JSON.stringify(value), { status })

  beforeEach(() => {
    vi.stubEnv("MS365_GRAPH_BASE_BACKOFF_MS", "0")
    vi.stubEnv("MS365_GRAPH_CIRCUIT_DISABLED", "true")
  })
  afterEach(() => {
    vi.unstubAllGlobals()
    vi.unstubAllEnvs()
  })

  it("honours retry settings for a simple upload", async () => {
    vi.stubEnv("MS365_GRAPH_MAX_RETRIES", "1")
    const fetchSpy = vi
      .fn()
      .mockResolvedValueOnce(new Response("throttled", { status: 429 }))
      .mockResolvedValueOnce(json(item))
    vi.stubGlobal("fetch", fetchSpy)
    expect((await simpleUpload(apiBase, path, "token", data, "application/octet-stream", "rename")).isRight()).toBe(
      true,
    )
    expect(fetchSpy).toHaveBeenCalledTimes(2)
  })

  it("honours retry settings for session creation and content chunks", async () => {
    vi.stubEnv("MS365_GRAPH_MAX_RETRIES", "1")
    const fetchSpy = vi
      .fn()
      .mockResolvedValueOnce(new Response("throttled", { status: 429 }))
      .mockResolvedValueOnce(json({ uploadUrl: "https://upload.example/session" }))
      .mockResolvedValueOnce(new Response("throttled", { status: 429 }))
      .mockResolvedValueOnce(json(item, 201))
    vi.stubGlobal("fetch", fetchSpy)
    expect((await sessionUpload(apiBase, path, "token", data, "rename")).isRight()).toBe(true)
    expect(fetchSpy.mock.calls.map(([, init]) => init.method)).toEqual(["POST", "POST", "PUT", "PUT"])
  })

  it("does not retry an upload when retries are disabled", async () => {
    vi.stubEnv("MS365_GRAPH_MAX_RETRIES", "0")
    const fetchSpy = vi.fn().mockResolvedValue(new Response("throttled", { status: 429 }))
    vi.stubGlobal("fetch", fetchSpy)
    expect((await simpleUpload(apiBase, path, "token", data, "application/octet-stream", "rename")).isLeft()).toBe(true)
    expect(fetchSpy).toHaveBeenCalledTimes(1)
  })

  it("never duplicates a renamed upload after an ambiguous server failure", async () => {
    vi.stubEnv("MS365_GRAPH_MAX_RETRIES", "3")
    const fetchSpy = vi.fn().mockResolvedValue(new Response("busy", { status: 503 }))
    vi.stubGlobal("fetch", fetchSpy)
    expect((await simpleUpload(apiBase, path, "token", data, "application/octet-stream", "rename")).isLeft()).toBe(true)
    expect(fetchSpy).toHaveBeenCalledTimes(1)
  })

  it("does not resend an upload chunk after an ambiguous network failure", async () => {
    vi.stubEnv("MS365_GRAPH_MAX_RETRIES", "3")
    const fetchSpy = vi
      .fn()
      .mockResolvedValueOnce(json({ uploadUrl: "https://upload.example/session" }))
      .mockRejectedValueOnce(new Error("response lost"))
    vi.stubGlobal("fetch", fetchSpy)
    expect((await sessionUpload(apiBase, path, "token", data, "rename")).isLeft()).toBe(true)
    expect(fetchSpy.mock.calls.map(([, init]) => init.method)).toEqual(["POST", "PUT"])
  })

  it("aborts an upload at the configured timeout", async () => {
    vi.stubEnv("MS365_GRAPH_MAX_RETRIES", "0")
    vi.stubEnv("MS365_GRAPH_TIMEOUT_MS", "5")
    const fetchSpy = vi.fn(
      (_url: string, init: RequestInit) =>
        new Promise<Response>((_resolve, reject) => {
          init.signal?.addEventListener("abort", () => reject(new Error("configured timeout")), { once: true })
        }),
    )
    vi.stubGlobal("fetch", fetchSpy)
    const result = await simpleUpload(apiBase, path, "token", data, "application/octet-stream", "rename")
    expect(result.isLeft()).toBe(true)
    expect((result.value as { message: string }).message).toContain("configured timeout")
    expect(fetchSpy.mock.calls[0]?.[1].signal?.aborted).toBe(true)
  })
})
