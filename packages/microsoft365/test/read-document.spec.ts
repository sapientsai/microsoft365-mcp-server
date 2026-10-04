import { Some } from "functype"
import { Left, Right } from "functype/either"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

vi.mock("../src/client/graph-client", () => ({
  getGraphClient: vi.fn(),
}))

vi.mock("../src/auth", () => ({
  getAccessToken: vi.fn(),
}))

import { getAccessToken } from "../src/auth"
import { getGraphClient } from "../src/client/graph-client"
import { readDocument } from "../src/tools/read-document-tools"

const PDF = "application/pdf"
const XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
const MB = 1024 * 1024

const mockClient = { graphQuery: vi.fn() }

/** Arrange a drive item for the metadata read. */
const givenItem = (item: { name?: string; size?: number; mimeType?: string }) => {
  mockClient.graphQuery.mockResolvedValue(
    Right({ id: "1", name: item.name ?? "doc.pdf", size: item.size, file: { mimeType: item.mimeType ?? PDF } }),
  )
}

const givenContent = (body: string, contentType = "text/plain") =>
  vi.fn().mockResolvedValue({
    ok: true,
    headers: new Headers({ "content-type": contentType }),
    arrayBuffer: () => Promise.resolve(new TextEncoder().encode(body).buffer),
  })

beforeEach(() => {
  vi.clearAllMocks()
  delete process.env.MS365_MAX_EXTRACT_BYTES
  vi.mocked(getGraphClient).mockReturnValue(Some(mockClient) as never)
  vi.mocked(getAccessToken).mockResolvedValue(Right("token-abc") as never)
})

afterEach(() => {
  vi.unstubAllGlobals()
  delete process.env.MS365_MAX_EXTRACT_BYTES
})

describe("readDocument — size caps", () => {
  // The whole point of the metadata-first ordering: an over-cap file must be rejected without ever
  // transferring the body. Asserting on the error message alone would pass even if the old
  // buffer-then-check ordering came back, so assert that fetch was never called.
  it("rejects an over-cap file from metadata, without fetching the body", async () => {
    givenItem({ name: "huge.xlsx", size: 30 * MB, mimeType: XLSX })
    const fetchSpy = givenContent("never read")
    vi.stubGlobal("fetch", fetchSpy)

    const result = await readDocument({ path: "/me/drive/items/1/content" })

    expect(result.isLeft()).toBe(true)
    expect((result.value as Error).message).toContain("extraction cap")
    expect(fetchSpy).not.toHaveBeenCalled()
  })

  it("allows a 60 MB PDF, which the old flat 10 MB cap would have rejected", async () => {
    givenItem({ name: "report.pdf", size: 60 * MB, mimeType: PDF })
    const fetchSpy = givenContent("plain text stand-in", "text/plain")
    vi.stubGlobal("fetch", fetchSpy)

    const result = await readDocument({ path: "/me/drive/items/1/content" })

    expect(result.isRight()).toBe(true)
    expect(fetchSpy).toHaveBeenCalledOnce()
  })

  it("applies per-format caps: 30 MB passes as PDF but fails as XLSX", async () => {
    vi.stubGlobal("fetch", givenContent("ok"))

    givenItem({ size: 30 * MB, mimeType: PDF })
    expect((await readDocument({ path: "/me/drive/items/1/content" })).isRight()).toBe(true)

    givenItem({ size: 30 * MB, mimeType: XLSX })
    expect((await readDocument({ path: "/me/drive/items/1/content" })).isLeft()).toBe(true)
  })

  it("MS365_MAX_EXTRACT_BYTES tightens every format", async () => {
    process.env.MS365_MAX_EXTRACT_BYTES = String(5 * MB)
    givenItem({ size: 6 * MB, mimeType: PDF })
    const fetchSpy = givenContent("never read")
    vi.stubGlobal("fetch", fetchSpy)

    const result = await readDocument({ path: "/me/drive/items/1/content" })

    expect(result.isLeft()).toBe(true)
    expect(fetchSpy).not.toHaveBeenCalled()
  })

  it("is a ceiling, never a raise — it cannot lift a format above its default", async () => {
    process.env.MS365_MAX_EXTRACT_BYTES = String(500 * MB)
    givenItem({ size: 30 * MB, mimeType: XLSX })
    vi.stubGlobal("fetch", givenContent("never read"))

    expect((await readDocument({ path: "/me/drive/items/1/content" })).isLeft()).toBe(true)
  })
})

describe("readDocument — request shape", () => {
  it("reads metadata from the item path, not the /content endpoint", async () => {
    givenItem({ size: 10 })
    vi.stubGlobal("fetch", givenContent("hi"))

    await readDocument({ path: "/drives/d1/items/i1/content" })

    expect(mockClient.graphQuery).toHaveBeenCalledWith(
      "GET",
      "/drives/d1/items/i1?$select=id,name,size,file",
      undefined,
      "v1.0",
    )
  })

  it("sends the resolved access token as a bearer", async () => {
    givenItem({ size: 10 })
    const fetchSpy = givenContent("hi")
    vi.stubGlobal("fetch", fetchSpy)

    await readDocument({ path: "/me/drive/items/1/content" })

    expect(fetchSpy).toHaveBeenCalledWith(expect.stringContaining("/v1.0/me/drive/items/1/content"), {
      headers: { Authorization: "Bearer token-abc" },
    })
  })

  it("surfaces an auth failure rather than sending 'Bearer [object Object]'", async () => {
    vi.mocked(getAccessToken).mockResolvedValue(Left({ type: "token", message: "no token" }) as never)
    const fetchSpy = givenContent("never read")
    vi.stubGlobal("fetch", fetchSpy)

    const result = await readDocument({ path: "/me/drive/items/1/content" })

    expect(result.isLeft()).toBe(true)
    expect((result.value as Error).message).toBe("no token")
    expect(fetchSpy).not.toHaveBeenCalled()
  })
})

describe("readDocument — output", () => {
  it("returns extracted text with the filename and size header", async () => {
    givenItem({ name: "notes.txt", size: 5, mimeType: "text/plain" })
    vi.stubGlobal("fetch", givenContent("hello", "text/plain"))

    const result = await readDocument({ path: "/me/drive/items/1/content" })

    expect(result.value as string).toContain("File: notes.txt")
    expect(result.value as string).toContain("hello")
  })

  it("truncates at max_chars with an explicit marker", async () => {
    givenItem({ name: "long.txt", size: 100, mimeType: "text/plain" })
    vi.stubGlobal("fetch", givenContent("x".repeat(5000), "text/plain"))

    const result = await readDocument({ path: "/me/drive/items/1/content", max_chars: 1000 })

    expect(result.value as string).toContain("[truncated at 1,000 chars")
    expect(result.value as string).toContain("full document is 5,000 chars")
  })

  // Seen 2026-10-03: OneDrive stored a .log uploaded as text/plain as octet-stream, and this tool
  // refused it. Text is now decided by the bytes, not the MIME type.
  it("reads a .log and a .jsonl stored as octet-stream", async () => {
    givenItem({ name: "app.log", size: 20, mimeType: "application/octet-stream" })
    vi.stubGlobal("fetch", givenContent("INFO started", "application/octet-stream"))
    expect((await readDocument({ path: "/me/drive/items/1/content" })).value as string).toContain("INFO started")

    givenItem({ name: "events.jsonl", size: 20, mimeType: "application/octet-stream" })
    vi.stubGlobal("fetch", givenContent('{"a":1}', "application/octet-stream"))
    expect((await readDocument({ path: "/me/drive/items/1/content" })).value as string).toContain('{"a":1}')
  })

  it("refuses a binary file renamed to .log", async () => {
    givenItem({ name: "image.log", size: 20, mimeType: "application/octet-stream" })
    vi.stubGlobal("fetch", givenContent("\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR", "application/octet-stream"))

    const result = await readDocument({ path: "/me/drive/items/1/content" })

    expect(result.isLeft()).toBe(true)
    expect((result.value as Error).message).toContain("Unsupported content type")
  })

  // "unsupported" is the arm that means download_file is the right fallback, so the message says so.
  it("points at download_file for an unsupported content type", async () => {
    givenItem({ name: "photo.png", size: 10, mimeType: "image/png" })
    vi.stubGlobal("fetch", givenContent("\x89PNG", "image/png"))

    const result = await readDocument({ path: "/me/drive/items/1/content" })

    expect(result.isLeft()).toBe(true)
    expect((result.value as Error).message).toContain("Unsupported content type")
    expect((result.value as Error).message).toContain("download_file")
  })
})

// Regression: mail attachments were completely unreadable in production. `itemPathFrom` stripped
// only "/content", so an attachment path kept its "/$value" and the $select landed on an Edm.Stream
// resource. Graph answered "The type 'Edm.Stream' is not valid for $select or $expand" and the tool
// surfaced it as "Failed to get file info" — twice in one hour of real logs.
describe("readDocument — mail attachments", () => {
  const ATTACHMENT = "/me/messages/msg-1/attachments/att-1/$value"

  it("probes the attachment resource, not the $value stream", async () => {
    mockClient.graphQuery.mockResolvedValue(Right({ id: "att-1", name: "report.pdf", contentType: PDF, size: 1024 }))
    vi.stubGlobal("fetch", givenContent("hello", PDF))

    await readDocument({ path: ATTACHMENT })

    const probedPath = mockClient.graphQuery.mock.calls[0]?.[1] as string
    expect(probedPath).not.toContain("$value")
    expect(probedPath).toBe("/me/messages/msg-1/attachments/att-1?$select=id,name,contentType,size")
  })

  it("reads the cap from contentType, since an attachment has no file.mimeType", async () => {
    // 30 MB PDF: under the 100 MB PDF cap, over the 25 MB default. So this passes only if
    // contentType is actually consulted — ignoring it falls back to the default and rejects.
    mockClient.graphQuery.mockResolvedValue(Right({ id: "att-1", name: "big.pdf", contentType: PDF, size: 30 * MB }))
    // Metadata contentType drives the cap; the response content-type drives extraction. Keeping the
    // body text/plain isolates the cap decision from parser behaviour.
    const fetchSpy = givenContent("extracted")
    vi.stubGlobal("fetch", fetchSpy)

    const result = await readDocument({ path: ATTACHMENT })

    expect(result.isRight()).toBe(true)
    expect(fetchSpy).toHaveBeenCalled()
  })

  it("still probes drive items the old way", async () => {
    givenItem({ name: "doc.pdf", size: 1024, mimeType: PDF })
    vi.stubGlobal("fetch", givenContent("hello", PDF))

    await readDocument({ path: "/me/drive/items/1/content" })

    expect(mockClient.graphQuery.mock.calls[0]?.[1]).toBe("/me/drive/items/1?$select=id,name,size,file")
  })

  // The other /$value shape. Its metadata resource is a message, which has no `name` or `file`, so
  // probing it would 400 as opaquely as the bug above.
  it("rejects whole-message MIME with a usable message instead of a Graph 400", async () => {
    const fetchSpy = givenContent("never read")
    vi.stubGlobal("fetch", fetchSpy)

    const result = await readDocument({ path: "/me/messages/msg-1/$value" })

    expect(result.isLeft()).toBe(true)
    expect((result.value as Error).message).toContain("get_message")
    expect(mockClient.graphQuery).not.toHaveBeenCalled()
    expect(fetchSpy).not.toHaveBeenCalled()
  })
})
