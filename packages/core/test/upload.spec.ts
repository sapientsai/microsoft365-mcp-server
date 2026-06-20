import { afterEach, describe, expect, it, vi } from "vitest"

import { describeFetchError, simpleUpload } from "../src/upload/upload"
import { isUploadTicket, mintUploadTicket, resolveUploadTicket } from "../src/upload/upload-ticket"

describe("upload-ticket", () => {
  it("mints an opaque prefixed ticket that resolves back to the token", () => {
    const ticket = mintUploadTicket("DELEGATED.JWT.VALUE")
    expect(isUploadTicket(ticket)).toBe(true)
    expect(ticket).not.toContain("DELEGATED.JWT.VALUE")
    expect(resolveUploadTicket(ticket)).toBe("DELEGATED.JWT.VALUE")
  })

  it("returns undefined for an unknown ticket", () => {
    expect(resolveUploadTicket("upl_does-not-exist")).toBeUndefined()
  })

  it("expires tickets past their TTL", () => {
    const t0 = 1_000_000
    const ticket = mintUploadTicket("tok", 1000, t0)
    expect(resolveUploadTicket(ticket, t0 + 500)).toBe("tok")
    expect(resolveUploadTicket(ticket, t0 + 1001)).toBeUndefined()
    // once expired it is purged, so even an earlier clock can't see it again
    expect(resolveUploadTicket(ticket, t0 + 500)).toBeUndefined()
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
