import { readFile, rm, stat } from "node:fs/promises"
import { tmpdir } from "node:os"
import { basename, join } from "node:path"

import { Some } from "functype"
import { Left, Right } from "functype/either"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

vi.mock("../src/client/graph-client", () => ({ getGraphClient: vi.fn() }))
vi.mock("../src/auth", () => ({ getAccessToken: vi.fn() }))
vi.mock("../src/utils/message-refs", async (importOriginal) => ({
  ...(await importOriginal<typeof import("../src/utils/message-refs")>()),
  resolveMessageIdOrRef: vi.fn(),
}))

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
  vi.mocked(resolveMessageIdOrRef).mockImplementation((id: string) =>
    id === "bad-ref" ? { kind: "unknown", ref: id } : { kind: "id", id: "MSG-ID" },
  )
  outDir = join(tmpdir(), `save-attachment-test-${Date.now()}-${Math.random().toString(36).slice(2)}`)
  vi.stubGlobal(
    "fetch",
    vi.fn(async () => new Response(Buffer.from("PDFBYTES"), { status: 200 })),
  )
})

afterEach(async () => {
  vi.unstubAllGlobals()
  vi.unstubAllEnvs()
  await rm(outDir, { recursive: true, force: true })
})

describe("saveAttachment", () => {
  it("uses the configured Graph API version for the download", async () => {
    vi.stubEnv("MS365_GRAPH_VERSION", "beta")
    mockClient.listAttachments.mockResolvedValue(Right({ value: [PDF] }))
    const result = await saveAttachment({ message_id: "1", out_dir: outDir })
    expect(result.isRight()).toBe(true)
    expect(vi.mocked(fetch).mock.calls[0]?.[0]).toContain("/beta/me/messages/")
  })

  it("honours the Graph retry settings when downloading attachment bytes", async () => {
    vi.stubEnv("MS365_GRAPH_MAX_RETRIES", "1")
    vi.stubEnv("MS365_GRAPH_BASE_BACKOFF_MS", "0")
    mockClient.listAttachments.mockResolvedValue(Right({ value: [PDF] }))
    vi.mocked(fetch).mockResolvedValueOnce(new Response("busy", { status: 503 }))
    const result = await saveAttachment({ message_id: "1", out_dir: outDir })
    expect(result.isRight()).toBe(true)
    expect(fetch).toHaveBeenCalledTimes(2)
  })

  it("does not retry when MS365_GRAPH_MAX_RETRIES is zero", async () => {
    vi.stubEnv("MS365_GRAPH_MAX_RETRIES", "0")
    mockClient.listAttachments.mockResolvedValue(Right({ value: [PDF] }))
    vi.mocked(fetch).mockResolvedValueOnce(new Response("busy", { status: 503 }))
    const result = await saveAttachment({ message_id: "1", out_dir: outDir })
    expect(result.isLeft()).toBe(true)
    expect(fetch).toHaveBeenCalledTimes(1)
  })

  it("writes the single attachment and returns its path", async () => {
    mockClient.listAttachments.mockResolvedValue(Right({ value: [PDF] }))

    const result = await saveAttachment({ message_id: "1", out_dir: outDir })

    expect(result.isRight()).toBe(true)
    const target = join(outDir, "scan.pdf")
    expect(result.value).toContain(target)
    expect(await readFile(target, "utf-8")).toBe("PDFBYTES")
  })

  it.skipIf(process.platform === "win32")("makes the saved file readable only by its owner", async () => {
    mockClient.listAttachments.mockResolvedValue(Right({ value: [PDF] }))

    await saveAttachment({ message_id: "1", out_dir: outDir })

    expect((await stat(join(outDir, "scan.pdf"))).mode & 0o777).toBe(0o600)
  })

  // Advertised as additive, so a second save must never replace the first — two messages can carry
  // an attachment with the same name.
  it("saves under a numbered name instead of overwriting an existing file", async () => {
    mockClient.listAttachments.mockResolvedValue(Right({ value: [PDF] }))
    await saveAttachment({ message_id: "1", out_dir: outDir })
    vi.mocked(fetch).mockResolvedValueOnce(new Response(Buffer.from("SECOND"), { status: 200 }))

    const result = await saveAttachment({ message_id: "1", out_dir: outDir })

    const second = join(outDir, "scan (1).pdf")
    expect(result.value).toContain(`Saved: ${second}`)
    expect(await readFile(join(outDir, "scan.pdf"), "utf-8")).toBe("PDFBYTES")
    expect(await readFile(second, "utf-8")).toBe("SECOND")
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
    mockClient.listAttachments.mockResolvedValue(Right({ value: [{ ...PDF, name: "../../../etc/authorized_keys" }] }))

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

  it.each([
    ["CON.pdf", "_CON.pdf"],
    ["nul", "_nul.pdf"],
    ["com1.tar.gz", "_com1.tar.gz"],
    ["Lpt9 .txt", "_Lpt9 .txt"],
  ])("prefixes the Windows reserved name %s", async (name, expected) => {
    mockClient.listAttachments.mockResolvedValue(Right({ value: [{ ...PDF, name }] }))

    const result = await saveAttachment({ message_id: "1", out_dir: outDir })

    expect(result.value).toContain(`Saved: ${join(outDir, expected)}`)
  })

  it("keeps names that only start like a reserved name", async () => {
    mockClient.listAttachments.mockResolvedValue(Right({ value: [{ ...PDF, name: "Console log.pdf" }] }))

    const result = await saveAttachment({ message_id: "1", out_dir: outDir })

    expect(result.value).toContain(`Saved: ${join(outDir, "Console log.pdf")}`)
  })

  // An attached email's name is its subject and its $value is MIME, with no content type.
  it("saves an attached email as .eml even when its subject contains a dot", async () => {
    mockClient.listAttachments.mockResolvedValue(
      Right({
        value: [
          {
            id: "att-3",
            name: "Re: Q3 report v1.2",
            contentType: undefined,
            "@odata.type": "#microsoft.graph.itemAttachment",
          },
        ],
      }),
    )

    const result = await saveAttachment({ message_id: "1", out_dir: outDir })

    expect(result.value).toContain(`Saved: ${join(outDir, "Re_ Q3 report v1.2.eml")}`)
    expect(result.value).toContain("Type: message/rfc822")
  })

  it("does not double the extension of an attached email already named .eml", async () => {
    mockClient.listAttachments.mockResolvedValue(
      Right({ value: [{ id: "att-3", name: "forwarded.EML", "@odata.type": "#microsoft.graph.itemAttachment" }] }),
    )

    const result = await saveAttachment({ message_id: "1", out_dir: outDir })

    expect(result.value).toContain(`Saved: ${join(outDir, "forwarded.EML")}`)
  })

  // Filesystems limit names in bytes, so multi-byte characters hit the limit long before 255 of them.
  it("shortens a long name to fit the filesystem and keeps its extension", async () => {
    mockClient.listAttachments.mockResolvedValue(Right({ value: [{ ...PDF, name: `${"請求書".repeat(60)}.pdf` }] }))

    const result = await saveAttachment({ message_id: "1", out_dir: outDir })

    expect(result.isRight()).toBe(true)
    const saved = basename((result.value as string).split("\n")[0].replace("Saved: ", ""))
    expect(saved.endsWith(".pdf")).toBe(true)
    expect(Buffer.byteLength(saved)).toBeLessThanOrEqual(200)
  })

  it("treats an overlong extension as part of the name", async () => {
    mockClient.listAttachments.mockResolvedValue(Right({ value: [{ ...PDF, name: `report.${"x".repeat(300)}` }] }))

    const result = await saveAttachment({ message_id: "1", out_dir: outDir })

    expect(result.isRight()).toBe(true)
    const saved = basename((result.value as string).split("\n")[0].replace("Saved: ", ""))
    expect(saved.startsWith("report.xxx")).toBe(true)
    expect(saved.endsWith(".pdf")).toBe(true)
    expect(Buffer.byteLength(saved)).toBeLessThanOrEqual(200)
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

  describe("mailbox", () => {
    it("lists and downloads from the named mailbox when it is allowed", async () => {
      vi.stubEnv("MS365_ALLOWED_MAILBOXES", "bel@example.com")
      mockClient.listAttachments.mockResolvedValue(Right({ value: [PDF] }))

      const result = await saveAttachment({ message_id: "1", out_dir: outDir, mailbox: "Bel@example.com" })

      expect(result.isRight()).toBe(true)
      expect(mockClient.listAttachments).toHaveBeenCalledWith("MSG-ID", "/users/bel%40example.com")
      expect(resolveMessageIdOrRef).toHaveBeenCalledWith("1", "bel@example.com")
      expect(vi.mocked(fetch).mock.calls[0]?.[0]).toBe(
        "https://graph.microsoft.com/v1.0/users/bel%40example.com/messages/MSG-ID/attachments/att-1/$value",
      )
    })

    it("stays on /me when no mailbox is given", async () => {
      mockClient.listAttachments.mockResolvedValue(Right({ value: [PDF] }))

      await saveAttachment({ message_id: "1", out_dir: outDir })

      expect(mockClient.listAttachments).toHaveBeenCalledWith("MSG-ID", "/me")
      expect(vi.mocked(fetch).mock.calls[0]?.[0]).toContain("/v1.0/me/messages/MSG-ID/")
    })

    it("refuses a mailbox the deployment does not allow, before calling Graph", async () => {
      vi.stubEnv("MS365_ALLOWED_MAILBOXES", "bel@example.com")

      const result = await saveAttachment({ message_id: "1", out_dir: outDir, mailbox: "other@example.com" })

      expect(result.isLeft()).toBe(true)
      expect((result.value as { message: string }).message).toContain("not permitted")
      expect(mockClient.listAttachments).not.toHaveBeenCalled()
    })

    // A ref from a scan of one mailbox names an unrelated message in another; it must fail, not resolve.
    it("rejects a ref minted for a different mailbox", async () => {
      vi.mocked(resolveMessageIdOrRef).mockReturnValue({ kind: "wrong-mailbox", ref: "k7q2-1", mintedFor: undefined })
      vi.stubEnv("MS365_ALLOWED_MAILBOXES", "bel@example.com")

      const result = await saveAttachment({ message_id: "k7q2-1", out_dir: outDir, mailbox: "bel@example.com" })

      expect(result.isLeft()).toBe(true)
      expect(mockClient.listAttachments).not.toHaveBeenCalled()
    })
  })
})
