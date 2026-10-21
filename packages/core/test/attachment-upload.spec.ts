import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

import {
  ATTACHMENT_CHUNK_SIZE,
  ATTACHMENT_INLINE_LIMIT,
  attachFile,
  attachFileToDraft,
  MAX_ATTACHMENT_SIZE,
} from "../src/upload/attachment-upload"

const apiBase = "https://graph.microsoft.com/v1.0"
const uploadUrl = "https://outlook.office365.com/api/gv1.0/users('u')/messages('m')/AttachmentSessions('s')?authtoken=t"
const json = (value: unknown, status = 200) => new Response(JSON.stringify(value), { status })
const draft = () => json({ id: "msg", isDraft: true })

type Call = { url: string; method: string; headers: Record<string, string> }
const callsOf = (spy: ReturnType<typeof vi.fn>): ReadonlyArray<Call> =>
  spy.mock.calls.map(([url, init]: [string, RequestInit | undefined]) => ({
    url,
    method: init?.method ?? "GET",
    headers: (init?.headers ?? {}) as Record<string, string>,
  }))

beforeEach(() => {
  vi.stubEnv("MS365_GRAPH_BASE_BACKOFF_MS", "0")
  vi.stubEnv("MS365_GRAPH_CIRCUIT_DISABLED", "true")
})
afterEach(() => {
  vi.unstubAllGlobals()
  vi.unstubAllEnvs()
})

describe("attachFileToDraft", () => {
  it("attaches a small file inline after confirming the message is a draft", async () => {
    const fetchSpy = vi
      .fn()
      .mockResolvedValueOnce(draft())
      .mockResolvedValueOnce(json({ id: "att-1", size: 4 }, 201))
    vi.stubGlobal("fetch", fetchSpy)

    const result = await attachFileToDraft(apiBase, "a/b=", "tok", Buffer.from("%PDF"), "report.pdf", "application/pdf")

    expect(result.isRight()).toBe(true)
    expect(result.value).toEqual({ id: "att-1", name: "report.pdf", size: 4 })
    const [check, post] = callsOf(fetchSpy)
    expect(check?.url).toBe(`${apiBase}/me/messages/a%2Fb%3D?$select=id,isDraft`)
    expect(post?.url).toBe(`${apiBase}/me/messages/a%2Fb%3D/attachments`)
    expect(post?.method).toBe("POST")
    const body = JSON.parse(fetchSpy.mock.calls[1]?.[1]?.body as string) as Record<string, string>
    expect(body).toEqual({
      "@odata.type": "#microsoft.graph.fileAttachment",
      name: "report.pdf",
      contentType: "application/pdf",
      contentBytes: Buffer.from("%PDF").toString("base64"),
    })
  })

  it("refuses a message that is not a draft without attaching anything", async () => {
    const fetchSpy = vi.fn().mockResolvedValueOnce(json({ id: "msg", isDraft: false }))
    vi.stubGlobal("fetch", fetchSpy)

    const result = await attachFileToDraft(apiBase, "msg", "tok", Buffer.from("x"), "x.txt", "text/plain")

    expect(result.isLeft()).toBe(true)
    expect(result.value).toMatchObject({ status: 409 })
    expect(fetchSpy).toHaveBeenCalledTimes(1)
  })

  it("refuses a file over 150 MB before calling Graph", async () => {
    const fetchSpy = vi.fn()
    vi.stubGlobal("fetch", fetchSpy)

    const result = await attachFile(
      apiBase,
      "msg",
      "tok",
      Buffer.alloc(MAX_ATTACHMENT_SIZE + 1),
      "big.bin",
      "application/octet-stream",
    )

    expect(result.value).toMatchObject({ status: 413 })
    expect(fetchSpy).not.toHaveBeenCalled()
  })

  it("uploads a 3 MB file through a session in ranges under 4 MB, without a bearer on the PUTs", async () => {
    const size = ATTACHMENT_INLINE_LIMIT + ATTACHMENT_CHUNK_SIZE // two ranges
    const fetchSpy = vi
      .fn()
      .mockResolvedValueOnce(draft())
      .mockResolvedValueOnce(json({ uploadUrl, nextExpectedRanges: ["0-"] }, 201))
      .mockResolvedValueOnce(json({ nextExpectedRanges: [String(ATTACHMENT_CHUNK_SIZE)] }))
      .mockResolvedValueOnce(
        new Response(null, {
          status: 201,
          headers: {
            Location: "https://outlook.office365.com/api/gv1.0/users('u')/messages('m')/Attachments('att-9')",
          },
        }),
      )
    vi.stubGlobal("fetch", fetchSpy)

    const result = await attachFileToDraft(apiBase, "msg", "tok", Buffer.alloc(size), "scan.pdf", "application/pdf")

    expect(result.value).toEqual({ id: "att-9", name: "scan.pdf", size })
    const [, create, first, last] = callsOf(fetchSpy)
    expect(create?.url).toBe(`${apiBase}/me/messages/msg/attachments/createUploadSession`)
    expect(JSON.parse(fetchSpy.mock.calls[1]?.[1]?.body as string)).toEqual({
      AttachmentItem: { attachmentType: "file", name: "scan.pdf", size, contentType: "application/pdf" },
    })
    for (const put of [first, last]) {
      expect(put?.url).toBe(uploadUrl)
      expect(put?.method).toBe("PUT")
      expect(put?.headers).not.toHaveProperty("Authorization")
      expect(Number(put?.headers["Content-Length"])).toBeLessThan(4 * 1024 * 1024)
    }
    expect(first?.headers["Content-Range"]).toBe(`bytes 0-${ATTACHMENT_CHUNK_SIZE - 1}/${size}`)
    expect(last?.headers["Content-Range"]).toBe(`bytes ${ATTACHMENT_CHUNK_SIZE}-${size - 1}/${size}`)
  })

  it("treats a 2xx on the last range as success even without the documented 201", async () => {
    const fetchSpy = vi
      .fn()
      .mockResolvedValueOnce(json({ uploadUrl }, 201))
      .mockResolvedValueOnce(json({ nextExpectedRanges: [] }))
    vi.stubGlobal("fetch", fetchSpy)

    const size = ATTACHMENT_INLINE_LIMIT
    const result = await attachFile(apiBase, "msg", "tok", Buffer.alloc(size), "scan.pdf", "application/pdf")

    expect(result.value).toEqual({ id: undefined, name: "scan.pdf", size })
    expect(callsOf(fetchSpy).map((call) => call.method)).toEqual(["POST", "PUT"])
  })

  it("falls back to a session when Graph refuses an inline POST just under the line as too large", async () => {
    const fetchSpy = vi
      .fn()
      .mockResolvedValueOnce(json({ error: { code: "RequestEntityTooLarge", message: "too large" } }, 413))
      .mockResolvedValueOnce(json({ uploadUrl }, 201))
      .mockResolvedValueOnce(new Response(null, { status: 201 }))
    vi.stubGlobal("fetch", fetchSpy)

    const result = await attachFile(
      apiBase,
      "msg",
      "tok",
      Buffer.alloc(ATTACHMENT_INLINE_LIMIT - 1),
      "scan.pdf",
      "application/pdf",
    )

    expect(result.isRight()).toBe(true)
    expect(callsOf(fetchSpy).map((call) => call.url)).toEqual([
      `${apiBase}/me/messages/msg/attachments`,
      `${apiBase}/me/messages/msg/attachments/createUploadSession`,
      uploadUrl,
    ])
  })

  it("falls back to inline when Outlook calls a file at the line too small for a session", async () => {
    const fetchSpy = vi
      .fn()
      .mockResolvedValueOnce(
        json({ error: { code: "ErrorAttachmentSizeShouldNotBeLessThanMinimumSize", message: "too small" } }, 400),
      )
      .mockResolvedValueOnce(json({ id: "att-2" }, 201))
    vi.stubGlobal("fetch", fetchSpy)

    const result = await attachFile(
      apiBase,
      "msg",
      "tok",
      Buffer.alloc(ATTACHMENT_INLINE_LIMIT),
      "scan.pdf",
      "application/pdf",
    )

    expect(result.value).toMatchObject({ id: "att-2" })
    expect(callsOf(fetchSpy).map((call) => call.url)).toEqual([
      `${apiBase}/me/messages/msg/attachments/createUploadSession`,
      `${apiBase}/me/messages/msg/attachments`,
    ])
  })

  it("does not fall back on an unrelated inline failure", async () => {
    const fetchSpy = vi.fn().mockResolvedValueOnce(json({ error: { code: "ErrorAccessDenied", message: "no" } }, 403))
    vi.stubGlobal("fetch", fetchSpy)

    const result = await attachFile(apiBase, "msg", "tok", Buffer.from("x"), "x.txt", "text/plain")

    expect(result.value).toMatchObject({ status: 403 })
    expect(fetchSpy).toHaveBeenCalledTimes(1)
  })

  it("will not send the file to a host other than Outlook", async () => {
    const fetchSpy = vi
      .fn()
      .mockResolvedValueOnce(draft())
      .mockResolvedValueOnce(json({ uploadUrl: "https://attacker.example/session" }, 201))
    vi.stubGlobal("fetch", fetchSpy)

    const result = await attachFileToDraft(
      apiBase,
      "msg",
      "tok",
      Buffer.alloc(ATTACHMENT_INLINE_LIMIT),
      "scan.pdf",
      "application/pdf",
    )

    expect(result.isLeft()).toBe(true)
    expect((result.value as { message: string }).message).toContain("attacker.example")
    expect(fetchSpy).toHaveBeenCalledTimes(2)
  })

  it("cancels the session and does not replay a range when a PUT fails", async () => {
    const fetchSpy = vi
      .fn()
      .mockResolvedValueOnce(draft())
      .mockResolvedValueOnce(json({ uploadUrl }, 201))
      .mockResolvedValueOnce(new Response("gateway", { status: 502 }))
      .mockResolvedValueOnce(new Response(null, { status: 204 }))
    vi.stubGlobal("fetch", fetchSpy)

    const result = await attachFileToDraft(
      apiBase,
      "msg",
      "tok",
      Buffer.alloc(ATTACHMENT_INLINE_LIMIT),
      "scan.pdf",
      "application/pdf",
    )

    expect(result.isLeft()).toBe(true)
    expect(callsOf(fetchSpy).map((call) => call.method)).toEqual(["GET", "POST", "PUT", "DELETE"])
  })

  it("names the Outlook host when the server cannot reach it", async () => {
    const thrown = new TypeError("fetch failed")
    ;(thrown as { cause?: unknown }).cause = Object.assign(new Error("getaddrinfo ENOTFOUND outlook.office365.com"), {
      code: "ENOTFOUND",
    })
    const fetchSpy = vi
      .fn()
      .mockResolvedValueOnce(draft())
      .mockResolvedValueOnce(json({ uploadUrl }, 201))
      .mockRejectedValueOnce(thrown)
      .mockRejectedValueOnce(thrown)
    vi.stubGlobal("fetch", fetchSpy)

    const result = await attachFileToDraft(
      apiBase,
      "msg",
      "tok",
      Buffer.alloc(ATTACHMENT_INLINE_LIMIT),
      "scan.pdf",
      "application/pdf",
    )

    expect(result.value).toMatchObject({ type: "network" })
    expect((result.value as { message: string }).message).toContain("uploading to outlook.office365.com")
  })
})
