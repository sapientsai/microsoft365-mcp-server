import { readFile, rm } from "node:fs/promises"
import { tmpdir } from "node:os"
import { join } from "node:path"

import { Some } from "functype"
import { Left, Right } from "functype/either"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

vi.mock("../src/client/graph-client", () => ({ getGraphClient: vi.fn() }))
vi.mock("../src/auth", () => ({ getAccessToken: vi.fn() }))
vi.mock("../src/utils/message-refs", () => ({ resolveMessageIdOrRef: vi.fn() }))

import { getAccessToken } from "../src/auth"
import { getGraphClient } from "../src/client/graph-client"
import { saveAttachment } from "../src/tools/save-attachment-tools"
import { resolveMessageIdOrRef } from "../src/utils/message-refs"

const mockClient = { listAttachments: vi.fn() }

const PDF = { id: "att-1", name: "scan.pdf", contentType: "application/pdf", size: 2048 }
const JPG = { id: "att-2", name: "photo.jpg", contentType: "image/jpeg", size: 1024 }
const REF = {
  id: "att-3",
  name: "Tax Folder",
  "@odata.type": "#microsoft.graph.referenceAttachment",
  sourceUrl: "https://www.icloud.com/iclouddrive/EXAMPLE",
  providerType: "other",
  isFolder: true,
  size: 0,
}

let outDir: string

beforeEach(() => {
  vi.clearAllMocks()
  vi.mocked(getGraphClient).mockReturnValue(Some(mockClient as never))
  vi.mocked(getAccessToken).mockResolvedValue(Right("TOKEN") as never)
  vi.mocked(resolveMessageIdOrRef).mockImplementation((id: string) => (id === "bad-ref" ? undefined : "MSG-ID"))
  outDir = join(tmpdir(), `save-attachment-test-${Date.now()}-${Math.random().toString(36).slice(2)}`)
  vi.stubGlobal(
    "fetch",
    vi.fn(async () => new Response(Buffer.from("PDFBYTES"), { status: 200 })),
  )
})

afterEach(async () => {
  vi.unstubAllGlobals()
  await rm(outDir, { recursive: true, force: true })
})

describe("saveAttachment", () => {
  it("writes the single attachment and returns its path", async () => {
    mockClient.listAttachments.mockResolvedValue(Right({ value: [PDF] }))

    const result = await saveAttachment({ message_id: "1", out_dir: outDir })

    expect(result.isRight()).toBe(true)
    const target = join(outDir, "scan.pdf")
    expect(result.value).toContain(target)
    expect(await readFile(target, "utf-8")).toBe("PDFBYTES")
  })

  // The whole point of the tool: read_document returns nothing useful for these, so the message must
  // not suggest text extraction as the next step.
  it("tells the caller to read the file directly rather than extract text", async () => {
    mockClient.listAttachments.mockResolvedValue(Right({ value: [PDF] }))

    const result = await saveAttachment({ message_id: "1", out_dir: outDir })

    expect(result.value).toContain("Read the file directly")
  })

  it("refuses to guess when several attachments exist, and lists them", async () => {
    mockClient.listAttachments.mockResolvedValue(Right({ value: [PDF, JPG] }))

    const result = await saveAttachment({ message_id: "1", out_dir: outDir })

    expect(result.isLeft()).toBe(true)
    const message = (result.value as { message: string }).message
    expect(message).toContain("att-1")
    expect(message).toContain("att-2")
    expect(message).toContain("scan.pdf")
  })

  it("selects by attachment_id when given", async () => {
    mockClient.listAttachments.mockResolvedValue(Right({ value: [PDF, JPG] }))

    const result = await saveAttachment({ message_id: "1", attachment_id: "att-2", out_dir: outDir })

    expect(result.isRight()).toBe(true)
    expect(result.value).toContain(join(outDir, "photo.jpg"))
  })

  it("rejects an attachment_id that is not on the message", async () => {
    mockClient.listAttachments.mockResolvedValue(Right({ value: [PDF] }))

    const result = await saveAttachment({ message_id: "1", attachment_id: "nope", out_dir: outDir })

    expect(result.isLeft()).toBe(true)
    expect((result.value as { message: string }).message).toContain("list_attachments")
  })

  // An attachment name is chosen by whoever sent the mail. Traversal must not reach the filesystem.
  it("neutralises path traversal in the attachment name", async () => {
    mockClient.listAttachments.mockResolvedValue(
      Right({ value: [{ ...PDF, name: "../../../etc/authorized_keys" }] }),
    )

    const result = await saveAttachment({ message_id: "1", out_dir: outDir })

    expect(result.isRight()).toBe(true)
    const path = (result.value as string).split("\n")[0].replace("Saved: ", "")
    // Dots surviving inside the filename are harmless; what matters is that no separator did, so the
    // write stays inside out_dir.
    expect(path.startsWith(`${outDir}/`)).toBe(true)
    expect(path.slice(outDir.length + 1)).not.toContain("/")
  })

  it("strips control characters from the attachment name", async () => {
    mockClient.listAttachments.mockResolvedValue(Right({ value: [{ ...PDF, name: "in\u0000voice\u001f.pdf" }] }))

    const result = await saveAttachment({ message_id: "1", out_dir: outDir })

    expect(result.isRight()).toBe(true)
    const path = (result.value as string).split("\n")[0].replace("Saved: ", "")
    expect(path).toBe(join(outDir, "in_voice_.pdf"))
  })

  // Ordinary names must survive intact — a sanitiser that mangles everything is as bad as none.
  it("leaves a normal filename untouched", async () => {
    mockClient.listAttachments.mockResolvedValue(
      Right({ value: [{ ...PDF, name: "Rates Notice 2024-25 (final).pdf" }] }),
    )

    const result = await saveAttachment({ message_id: "1", out_dir: outDir })

    expect(result.value).toContain(join(outDir, "Rates Notice 2024-25 (final).pdf"))
  })

  it("gives an unnamed attachment an extension from its content type", async () => {
    mockClient.listAttachments.mockResolvedValue(Right({ value: [{ ...PDF, name: undefined }] }))

    const result = await saveAttachment({ message_id: "1", out_dir: outDir })

    expect(result.isRight()).toBe(true)
    expect(result.value).toMatch(/attachment-att-1\.pdf/)
  })

  // A referenceAttachment is a link, not bytes. It cannot be saved — but it must never be dropped
  // silently: the document exists, and hiding it lets a sweep claim coverage it does not have.
  it("reports the URL of a cloud link rather than hiding it", async () => {
    mockClient.listAttachments.mockResolvedValue(Right({ value: [REF] }))

    const result = await saveAttachment({ message_id: "1", out_dir: outDir })

    expect(result.isLeft()).toBe(true)
    const message = (result.value as { message: string }).message
    expect(message).toContain("Tax Folder")
    expect(message).toContain("https://www.icloud.com/iclouddrive/EXAMPLE")
    expect(message).toContain("folder")
  })

  it("distinguishes no attachments at all from only-cloud-links", async () => {
    mockClient.listAttachments.mockResolvedValue(Right({ value: [] }))

    const result = await saveAttachment({ message_id: "1", out_dir: outDir })

    expect((result.value as { message: string }).message).toBe("This message has no attachments.")
  })

  it("explains why a cloud link cannot be saved when its id is passed", async () => {
    mockClient.listAttachments.mockResolvedValue(Right({ value: [PDF, REF] }))

    const result = await saveAttachment({ message_id: "1", attachment_id: "att-3", out_dir: outDir })

    expect(result.isLeft()).toBe(true)
    const message = (result.value as { message: string }).message
    expect(message).toContain("cloud link")
    expect(message).toContain("https://www.icloud.com/iclouddrive/EXAMPLE")
  })

  it("mentions cloud links alongside the choice list when several attachments exist", async () => {
    mockClient.listAttachments.mockResolvedValue(Right({ value: [PDF, JPG, REF] }))

    const result = await saveAttachment({ message_id: "1", out_dir: outDir })

    const message = (result.value as { message: string }).message
    expect(message).toContain("att-1")
    expect(message).toContain("cloud link")
    expect(message).toContain("https://www.icloud.com/iclouddrive/EXAMPLE")
  })

  it("still saves a real file when a cloud link sits alongside it", async () => {
    mockClient.listAttachments.mockResolvedValue(Right({ value: [PDF, REF] }))

    const result = await saveAttachment({ message_id: "1", out_dir: outDir })

    expect(result.isRight()).toBe(true)
    expect(result.value).toContain(join(outDir, "scan.pdf"))
  })

  it("rejects an attachment over the size cap before fetching it", async () => {
    mockClient.listAttachments.mockResolvedValue(Right({ value: [{ ...PDF, size: 200 * 1024 * 1024 }] }))

    const result = await saveAttachment({ message_id: "1", out_dir: outDir })

    expect(result.isLeft()).toBe(true)
    expect(globalThis.fetch).not.toHaveBeenCalled()
  })

  it("explains an unknown scan ref rather than failing opaquely", async () => {
    const result = await saveAttachment({ message_id: "bad-ref", out_dir: outDir })

    expect(result.isLeft()).toBe(true)
    expect((result.value as { message: string }).message).toContain("scan_messages")
  })

  it("surfaces a Graph error body", async () => {
    mockClient.listAttachments.mockResolvedValue(Right({ value: [PDF] }))
    vi.stubGlobal(
      "fetch",
      vi.fn(
        async () =>
          new Response(JSON.stringify({ error: { message: "Access denied" } }), {
            status: 403,
            headers: { "content-type": "application/json" },
          }),
      ),
    )

    const result = await saveAttachment({ message_id: "1", out_dir: outDir })

    expect(result.isLeft()).toBe(true)
    expect((result.value as { message: string }).message).toContain("Access denied")
  })

  it("reports a token failure", async () => {
    mockClient.listAttachments.mockResolvedValue(Right({ value: [PDF] }))
    vi.mocked(getAccessToken).mockResolvedValue(Left({ type: "token", message: "no token" }) as never)

    const result = await saveAttachment({ message_id: "1", out_dir: outDir })

    expect(result.isLeft()).toBe(true)
    expect((result.value as { message: string }).message).toContain("no token")
  })
})
