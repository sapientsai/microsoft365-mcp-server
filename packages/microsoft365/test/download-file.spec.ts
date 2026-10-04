import { Some } from "functype"
import { Right } from "functype/either"
import { beforeEach, describe, expect, it, vi } from "vitest"

vi.mock("../src/client/graph-client", () => ({
  getGraphClient: vi.fn(),
}))

import { getGraphClient } from "../src/client/graph-client"
import { downloadFile } from "../src/tools/files-tools"

const mockClient = { downloadFile: vi.fn(), downloadFileContent: vi.fn() }

const givenItem = (item: { name: string; size?: number; mimeType?: string; folder?: boolean }) =>
  mockClient.downloadFile.mockResolvedValue(
    Right({
      id: "1",
      name: item.name,
      size: item.size ?? 20,
      ...(item.folder ? { folder: { childCount: 0 } } : { file: { mimeType: item.mimeType } }),
    }),
  )

const givenBytes = (bytes: Uint8Array) => mockClient.downloadFileContent.mockResolvedValue(Right(bytes))

const PNG_HEADER = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d])

beforeEach(() => {
  vi.clearAllMocks()
  vi.mocked(getGraphClient).mockReturnValue(Some(mockClient) as never)
})

describe("downloadFile — inline content", () => {
  it("shows a text/plain file inline", async () => {
    givenItem({ name: "notes.txt", mimeType: "text/plain" })
    givenBytes(new TextEncoder().encode("hello"))

    const result = await downloadFile({ item_id: "1" })

    expect(result.value).toContain("## Content")
    expect(result.value).toContain("hello")
  })

  // Seen 2026-10-03: a .log uploaded as text/plain was stored as octet-stream and shown without content.
  it("shows a .log stored as octet-stream inline, after sniffing it", async () => {
    givenItem({ name: "app.log", mimeType: "application/octet-stream" })
    givenBytes(new TextEncoder().encode("INFO started"))

    const result = await downloadFile({ item_id: "1" })

    expect(result.value).toContain("INFO started")
  })

  it("shows a .jsonl stored as octet-stream inline", async () => {
    givenItem({ name: "events.jsonl", mimeType: "application/octet-stream" })
    givenBytes(new TextEncoder().encode('{"a":1}'))

    expect((await downloadFile({ item_id: "1" })).value).toContain('{"a":1}')
  })

  it("does not inline a binary file renamed to .log", async () => {
    givenItem({ name: "image.log", mimeType: "application/octet-stream" })
    givenBytes(PNG_HEADER)

    const result = await downloadFile({ item_id: "1" })

    expect(result.isRight()).toBe(true)
    expect(result.value).not.toContain("## Content")
  })

  it("never downloads a known binary type", async () => {
    givenItem({ name: "photo.png", mimeType: "image/png" })

    await downloadFile({ item_id: "1" })

    expect(mockClient.downloadFileContent).not.toHaveBeenCalled()
  })

  it("never downloads a file over the inline size cap", async () => {
    givenItem({ name: "big.log", mimeType: "application/octet-stream", size: 200 * 1024 })

    await downloadFile({ item_id: "1" })

    expect(mockClient.downloadFileContent).not.toHaveBeenCalled()
  })

  // A folder has no `file` facet, so its MIME type reads as unknown; it must not be fetched.
  it("never downloads a folder", async () => {
    givenItem({ name: "Reports", folder: true })

    await downloadFile({ item_id: "1" })

    expect(mockClient.downloadFileContent).not.toHaveBeenCalled()
  })
})
