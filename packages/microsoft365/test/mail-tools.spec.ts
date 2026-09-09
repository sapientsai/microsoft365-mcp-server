import { Some } from "functype"
import { Right } from "functype/either"
import { beforeEach, describe, expect, it, vi } from "vitest"

import type { GraphMessage } from "../src/types"

vi.mock("../src/client/graph-client", () => ({
  getGraphClient: vi.fn(),
}))

import { getGraphClient } from "../src/client/graph-client"
import {
  createDraft,
  createForwardDraft,
  createReplyAllDraft,
  batchMoveMessages,
  createReplyDraft,
  listAttachments,
  getMessage,
  scanMessages,
  listAttachments as listAttachmentsTool,
  listMailFolders,
  moveMessage,
  moveMessagesMatching,
  orderedFilter,
  sendDraft,
  sendForward,
  sendMessage,
  sendReply,
  sendReplyAll,
  summarizeSenders,
} from "../src/tools/mail-tools"

const mockClient = {
  sendMessage: vi.fn(),
  createDraft: vi.fn(),
  sendDraft: vi.fn(),
  sendReply: vi.fn(),
  sendReplyAll: vi.fn(),
  sendForward: vi.fn(),
  createReplyDraft: vi.fn(),
  createReplyAllDraft: vi.fn(),
  createForwardDraft: vi.fn(),
  getMessage: vi.fn(),
  listAttachments: vi.fn(),
  listMailFolders: vi.fn(),
  listFolderMessages: vi.fn(),
  listMessages: vi.fn(),
  moveMessage: vi.fn(),
  requestPaginated: vi.fn(),
  listFolderMessagesAll: vi.fn(),
  batchRequest: vi.fn(),
}

beforeEach(() => {
  vi.clearAllMocks()
  vi.mocked(getGraphClient).mockReturnValue(Some(mockClient as never))
})

describe("mail-tools", () => {
  describe("sendMessage", () => {
    it("should send a message with default content type", async () => {
      mockClient.sendMessage.mockResolvedValue(Right({}))
      const result = await sendMessage({ to: "alice@example.com", subject: "Hi", body: "Hello" })
      expect(result.isRight()).toBe(true)
      expect(result.value).toContain("alice@example.com")
      expect(mockClient.sendMessage).toHaveBeenCalledWith(
        {
          message: {
            subject: "Hi",
            body: { contentType: "Text", content: "Hello" },
            toRecipients: [{ emailAddress: { address: "alice@example.com" } }],
          },
        },
        "/me",
      )
    })

    it("should send a message with HTML content type", async () => {
      mockClient.sendMessage.mockResolvedValue(Right({}))
      await sendMessage({ to: "bob@example.com", subject: "Hi", body: "<b>Bold</b>", content_type: "HTML" })
      expect(mockClient.sendMessage).toHaveBeenCalledWith(
        {
          message: {
            subject: "Hi",
            body: { contentType: "HTML", content: "<b>Bold</b>" },
            toRecipients: [{ emailAddress: { address: "bob@example.com" } }],
          },
        },
        "/me",
      )
    })

    it("should split comma-separated 'to' into multiple toRecipients", async () => {
      mockClient.sendMessage.mockResolvedValue(Right({}))
      await sendMessage({
        to: "alice@example.com, bob@example.com,carol@example.com",
        subject: "Hi",
        body: "Hello",
      })
      const callArg = mockClient.sendMessage.mock.calls[0][0] as { message: Record<string, unknown> }
      expect(callArg.message.toRecipients).toEqual([
        { emailAddress: { address: "alice@example.com" } },
        { emailAddress: { address: "bob@example.com" } },
        { emailAddress: { address: "carol@example.com" } },
      ])
    })

    it("should reject empty 'to' field", async () => {
      const result = await sendMessage({ to: "", subject: "Hi", body: "Hello" })
      expect(result.isLeft()).toBe(true)
      expect((result.value as Error).message).toContain("recipient is required")
      expect(mockClient.sendMessage).not.toHaveBeenCalled()
    })

    it("should reject 'to' containing only whitespace and commas", async () => {
      const result = await sendMessage({ to: " , , ", subject: "Hi", body: "Hello" })
      expect(result.isLeft()).toBe(true)
      expect(mockClient.sendMessage).not.toHaveBeenCalled()
    })
  })

  describe("createDraft", () => {
    const draftResponse: Partial<GraphMessage> = { id: "draft-123", subject: "Test Draft" }

    it("should create a draft with basic params", async () => {
      mockClient.createDraft.mockResolvedValue(Right(draftResponse))
      const result = await createDraft({ to: "alice@example.com", subject: "Draft", body: "Content" })
      expect(result.isRight()).toBe(true)
      expect(result.value).toContain("draft-123")
      expect(mockClient.createDraft).toHaveBeenCalledWith(
        {
          subject: "Draft",
          body: { contentType: "Text", content: "Content" },
          toRecipients: [{ emailAddress: { address: "alice@example.com" } }],
        },
        "/me",
      )
    })

    it("should create a draft with HTML content type", async () => {
      mockClient.createDraft.mockResolvedValue(Right(draftResponse))
      await createDraft({ to: "alice@example.com", subject: "Draft", body: "<p>Hi</p>", content_type: "HTML" })
      expect(mockClient.createDraft).toHaveBeenCalledWith(
        expect.objectContaining({
          body: { contentType: "HTML", content: "<p>Hi</p>" },
        }),
        "/me",
      )
    })

    it("should create a draft with cc recipients", async () => {
      mockClient.createDraft.mockResolvedValue(Right(draftResponse))
      await createDraft({
        to: "alice@example.com",
        subject: "Draft",
        body: "Hi",
        cc: "bob@example.com,carol@example.com",
      })
      expect(mockClient.createDraft).toHaveBeenCalledWith(
        expect.objectContaining({
          ccRecipients: [
            { emailAddress: { address: "bob@example.com" } },
            { emailAddress: { address: "carol@example.com" } },
          ],
        }),
        "/me",
      )
    })

    it("should create a draft with bcc recipients", async () => {
      mockClient.createDraft.mockResolvedValue(Right(draftResponse))
      await createDraft({ to: "alice@example.com", subject: "Draft", body: "Hi", bcc: "secret@example.com" })
      expect(mockClient.createDraft).toHaveBeenCalledWith(
        expect.objectContaining({
          bccRecipients: [{ emailAddress: { address: "secret@example.com" } }],
        }),
        "/me",
      )
    })

    it("should handle cc with whitespace and empty entries", async () => {
      mockClient.createDraft.mockResolvedValue(Right(draftResponse))
      await createDraft({
        to: "alice@example.com",
        subject: "Draft",
        body: "Hi",
        cc: " bob@example.com , , carol@example.com ",
      })
      expect(mockClient.createDraft).toHaveBeenCalledWith(
        expect.objectContaining({
          ccRecipients: [
            { emailAddress: { address: "bob@example.com" } },
            { emailAddress: { address: "carol@example.com" } },
          ],
        }),
        "/me",
      )
    })

    it("should omit cc when empty string", async () => {
      mockClient.createDraft.mockResolvedValue(Right(draftResponse))
      await createDraft({ to: "alice@example.com", subject: "Draft", body: "Hi", cc: "" })
      const callArg = mockClient.createDraft.mock.calls[0][0] as Record<string, unknown>
      expect(callArg).not.toHaveProperty("ccRecipients")
    })

    it("should split comma-separated 'to' into multiple toRecipients", async () => {
      mockClient.createDraft.mockResolvedValue(Right(draftResponse))
      await createDraft({
        to: "alice@example.com, bob@example.com, carol@example.com",
        subject: "Draft",
        body: "Hi",
      })
      expect(mockClient.createDraft).toHaveBeenCalledWith(
        expect.objectContaining({
          toRecipients: [
            { emailAddress: { address: "alice@example.com" } },
            { emailAddress: { address: "bob@example.com" } },
            { emailAddress: { address: "carol@example.com" } },
          ],
        }),
        "/me",
      )
    })

    it("should trim whitespace and drop empty entries in 'to'", async () => {
      mockClient.createDraft.mockResolvedValue(Right(draftResponse))
      await createDraft({
        to: " alice@example.com , , bob@example.com ",
        subject: "Draft",
        body: "Hi",
      })
      expect(mockClient.createDraft).toHaveBeenCalledWith(
        expect.objectContaining({
          toRecipients: [
            { emailAddress: { address: "alice@example.com" } },
            { emailAddress: { address: "bob@example.com" } },
          ],
        }),
        "/me",
      )
    })

    it("should reject empty 'to' field", async () => {
      const result = await createDraft({ to: "", subject: "Draft", body: "Hi" })
      expect(result.isLeft()).toBe(true)
      expect((result.value as Error).message).toContain("recipient is required")
      expect(mockClient.createDraft).not.toHaveBeenCalled()
    })
  })

  describe("sendDraft", () => {
    it("should send a draft by ID", async () => {
      mockClient.sendDraft.mockResolvedValue(Right({}))
      const result = await sendDraft({ message_id: "draft-123" })
      expect(result.isRight()).toBe(true)
      expect(result.value).toContain("Draft sent successfully")
      expect(mockClient.sendDraft).toHaveBeenCalledWith("draft-123", "/me")
    })
  })

  describe("sendReply", () => {
    it("should send a reply by message ID", async () => {
      mockClient.sendReply.mockResolvedValue(Right({}))
      const result = await sendReply({ message_id: "msg-1", comment: "Thanks!" })
      expect(result.isRight()).toBe(true)
      expect(result.value).toContain("Reply sent successfully")
      expect(mockClient.sendReply).toHaveBeenCalledWith("msg-1", "Thanks!", "/me")
    })
  })

  describe("sendReplyAll", () => {
    it("should send a reply-all by message ID", async () => {
      mockClient.sendReplyAll.mockResolvedValue(Right({}))
      const result = await sendReplyAll({ message_id: "msg-1", comment: "Thanks all!" })
      expect(result.isRight()).toBe(true)
      expect(result.value).toContain("Reply-all sent successfully")
      expect(mockClient.sendReplyAll).toHaveBeenCalledWith("msg-1", "Thanks all!", "/me")
    })
  })

  describe("sendForward", () => {
    it("should forward with recipients and an optional comment", async () => {
      mockClient.sendForward.mockResolvedValue(Right({}))
      const result = await sendForward({ message_id: "msg-1", to: "alice@example.com", comment: "FYI" })
      expect(result.isRight()).toBe(true)
      expect(result.value).toContain("alice@example.com")
      expect(mockClient.sendForward).toHaveBeenCalledWith(
        "msg-1",
        "FYI",
        [{ emailAddress: { address: "alice@example.com" } }],
        "/me",
      )
    })

    it("should default an omitted comment to an empty string", async () => {
      mockClient.sendForward.mockResolvedValue(Right({}))
      await sendForward({ message_id: "msg-1", to: "alice@example.com" })
      expect(mockClient.sendForward).toHaveBeenCalledWith(
        "msg-1",
        "",
        [{ emailAddress: { address: "alice@example.com" } }],
        "/me",
      )
    })

    it("should reject an empty 'to' field", async () => {
      const result = await sendForward({ message_id: "msg-1", to: "" })
      expect(result.isLeft()).toBe(true)
      expect((result.value as Error).message).toContain("recipient is required")
      expect(mockClient.sendForward).not.toHaveBeenCalled()
    })
  })

  describe("createReplyDraft", () => {
    it("should create a threaded reply draft and return its ID", async () => {
      mockClient.createReplyDraft.mockResolvedValue(Right({ id: "draft-r1" }))
      const result = await createReplyDraft({ message_id: "msg-1", comment: "Will do" })
      expect(result.isRight()).toBe(true)
      expect(result.value).toContain("draft-r1")
      expect(result.value).toContain("send_draft")
      expect(mockClient.createReplyDraft).toHaveBeenCalledWith("msg-1", "Will do", "/me")
    })
  })

  describe("createReplyAllDraft", () => {
    it("should create a threaded reply-all draft and return its ID", async () => {
      mockClient.createReplyAllDraft.mockResolvedValue(Right({ id: "draft-ra1" }))
      const result = await createReplyAllDraft({ message_id: "msg-1", comment: "Will do" })
      expect(result.isRight()).toBe(true)
      expect(result.value).toContain("draft-ra1")
      expect(mockClient.createReplyAllDraft).toHaveBeenCalledWith("msg-1", "Will do", "/me")
    })
  })

  describe("createForwardDraft", () => {
    it("should create a forward draft with recipients and return its ID", async () => {
      mockClient.createForwardDraft.mockResolvedValue(Right({ id: "draft-f1" }))
      const result = await createForwardDraft({ message_id: "msg-1", to: "alice@example.com", comment: "FYI" })
      expect(result.isRight()).toBe(true)
      expect(result.value).toContain("draft-f1")
      expect(mockClient.createForwardDraft).toHaveBeenCalledWith(
        "msg-1",
        "FYI",
        [{ emailAddress: { address: "alice@example.com" } }],
        "/me",
      )
    })

    it("should reject an empty 'to' field", async () => {
      const result = await createForwardDraft({ message_id: "msg-1", to: "" })
      expect(result.isLeft()).toBe(true)
      expect((result.value as Error).message).toContain("recipient is required")
      expect(mockClient.createForwardDraft).not.toHaveBeenCalled()
    })
  })

  describe("listMailFolders", () => {
    it("should list folders with their counts", async () => {
      mockClient.listMailFolders.mockResolvedValue(
        Right({ value: [{ id: "f1", displayName: "Archive", totalItemCount: 12, unreadItemCount: 3 }] }),
      )
      const result = await listMailFolders()
      expect(result.isRight()).toBe(true)
      expect(result.value).toContain("Archive")
      expect(result.value).toContain("12 items, 3 unread")
      expect(mockClient.listMailFolders).toHaveBeenCalledWith({ $top: 100 }, "/me")
    })

    it("should page through all folders when asked", async () => {
      mockClient.requestPaginated.mockResolvedValue(Right([{ id: "f1", displayName: "Archive" }]))
      const result = await listMailFolders({ fetch_all_pages: true })
      expect(result.isRight()).toBe(true)
      expect(mockClient.requestPaginated).toHaveBeenCalledWith("/me/mailFolders")
      expect(mockClient.listMailFolders).not.toHaveBeenCalled()
    })
  })

  describe("moveMessage", () => {
    it("should pass a well-known folder name straight through", async () => {
      mockClient.moveMessage.mockResolvedValue(Right({ id: "msg-1", subject: "Receipt" }))
      const result = await moveMessage({ message_id: "msg-1", destination: "archive" })
      expect(result.isRight()).toBe(true)
      expect(mockClient.moveMessage).toHaveBeenCalledWith("msg-1", "archive", "/me")
      expect(mockClient.listMailFolders).not.toHaveBeenCalled()
    })

    // Triage moves in batches; echoing each body back would flood the caller's context.
    it("should confirm tersely without echoing the message body", async () => {
      mockClient.moveMessage.mockResolvedValue(
        Right({ id: "new-id", subject: "Receipt", body: { content: "a very long message body" } }),
      )
      const result = await moveMessage({ message_id: "msg-1", destination: "archive" })
      expect(result.value).toBe('Moved "Receipt" to the archive folder. New ID: new-id')
      expect(result.value).not.toContain("a very long message body")
    })

    it("should name an untitled message rather than printing undefined", async () => {
      mockClient.moveMessage.mockResolvedValue(Right({ id: "new-id" }))
      const result = await moveMessage({ message_id: "msg-1", destination: "archive" })
      expect(result.value).toContain("(No Subject)")
    })

    it("should map a well-known alias and ignore case", async () => {
      mockClient.moveMessage.mockResolvedValue(Right({ id: "msg-1" }))
      await moveMessage({ message_id: "msg-1", destination: "Deleted Items" })
      expect(mockClient.moveMessage).toHaveBeenCalledWith("msg-1", "deleteditems", "/me")
    })

    it("should resolve a folder display name to its ID", async () => {
      mockClient.listMailFolders.mockResolvedValue(Right({ value: [{ id: "f-receipts", displayName: "Receipts" }] }))
      mockClient.moveMessage.mockResolvedValue(Right({ id: "msg-1" }))
      await moveMessage({ message_id: "msg-1", destination: "Receipts" })
      expect(mockClient.moveMessage).toHaveBeenCalledWith("msg-1", "f-receipts", "/me")
    })

    it("should error rather than guess when a display name is ambiguous", async () => {
      mockClient.listMailFolders.mockResolvedValue(
        Right({
          value: [
            { id: "f-a", displayName: "Receipts" },
            { id: "f-b", displayName: "Receipts" },
          ],
        }),
      )
      const result = await moveMessage({ message_id: "msg-1", destination: "Receipts" })
      expect(result.isLeft()).toBe(true)
      expect((result.value as Error).message).toContain("Multiple folders")
      expect((result.value as Error).message).toContain("f-a")
      expect(mockClient.moveMessage).not.toHaveBeenCalled()
    })

    it("should fall through to Graph when nothing matches, assuming a folder ID", async () => {
      mockClient.listMailFolders.mockResolvedValue(Right({ value: [{ id: "f1", displayName: "Archive" }] }))
      mockClient.moveMessage.mockResolvedValue(Right({ id: "msg-1" }))
      await moveMessage({ message_id: "msg-1", destination: "AAMkAGI0-opaque-id" })
      expect(mockClient.moveMessage).toHaveBeenCalledWith("msg-1", "AAMkAGI0-opaque-id", "/me")
    })
  })
  describe("getMessage body_format", () => {
    it("should request no Prefer header by default", async () => {
      mockClient.getMessage.mockResolvedValue(Right({ id: "m1", subject: "Hi" }))
      await getMessage({ message_id: "m1" })
      expect(mockClient.getMessage).toHaveBeenCalledWith("m1", undefined, "/me")
    })

    // Marketing mail is mostly CSS; asking Graph for text is a large context saving.
    it("should pass the requested body format through", async () => {
      mockClient.getMessage.mockResolvedValue(Right({ id: "m1", subject: "Hi" }))
      await getMessage({ message_id: "m1", body_format: "text" })
      expect(mockClient.getMessage).toHaveBeenCalledWith("m1", "text", "/me")
    })
  })

  describe("batchMoveMessages", () => {
    it("should resolve the destination once, not per message", async () => {
      mockClient.listMailFolders.mockResolvedValue(Right({ value: [{ id: "f1", displayName: "Receipts" }] }))
      mockClient.moveMessage.mockResolvedValue(Right({ id: "new", subject: "s" }))
      const result = await batchMoveMessages({ message_ids: ["a", "b", "c"], destination: "Receipts" })
      expect(result.isRight()).toBe(true)
      expect(mockClient.listMailFolders).toHaveBeenCalledTimes(1)
      expect(mockClient.moveMessage).toHaveBeenCalledTimes(3)
      expect(result.value).toContain("Moved 3/3")
    })

    // A silent partial success is the worst outcome: the caller believes the inbox is
    // filed when some of it is not.
    it("should report partial failure per message", async () => {
      const { Left: L } = await import("functype/either")
      mockClient.moveMessage
        .mockResolvedValueOnce(Right({ id: "n1", subject: "ok" }))
        .mockResolvedValueOnce(L({ message: "ErrorItemNotFound" }))
      const result = await batchMoveMessages({ message_ids: ["a", "bad"], destination: "archive" })
      expect(result.isRight()).toBe(true)
      expect(result.value).toContain("Moved 1/2")
      expect(result.value).toContain("FAILED bad")
      expect(result.value).toContain("ErrorItemNotFound")
    })

    // Sequencing is the reason this is a reduce and not Promise.all: a 429 partway
    // through a parallel batch would leave the caller unsure what landed.
    it("should move sequentially, not in parallel", async () => {
      const inFlight = { current: 0, max: 0 }
      mockClient.moveMessage.mockImplementation(async () => {
        inFlight.current += 1
        inFlight.max = Math.max(inFlight.max, inFlight.current)
        await new Promise((r) => setTimeout(r, 1))
        inFlight.current -= 1
        return Right({ id: "n", subject: "s" })
      })
      await batchMoveMessages({ message_ids: ["a", "b", "c"], destination: "archive" })
      expect(inFlight.max).toBe(1)
    })

    it("should reject an empty list", async () => {
      const result = await batchMoveMessages({ message_ids: [], destination: "archive" })
      expect(result.isLeft()).toBe(true)
      expect(mockClient.moveMessage).not.toHaveBeenCalled()
    })

    it("should refuse an oversized batch rather than half-file it", async () => {
      const ids = Array.from({ length: 51 }, (_, i) => `id-${i}`)
      const result = await batchMoveMessages({ message_ids: ids, destination: "archive" })
      expect(result.isLeft()).toBe(true)
      expect((result.value as Error).message).toContain("at most 50")
      expect(mockClient.moveMessage).not.toHaveBeenCalled()
    })
  })

  describe("listAttachments", () => {
    it("should list attachments with a read_document path for each", async () => {
      mockClient.listAttachments.mockResolvedValue(
        Right({
          value: [{ id: "att-1", name: "invoice.pdf", contentType: "application/pdf", size: 20480 }],
        }),
      )
      const result = await listAttachments({ message_id: "msg-1" })
      expect(result.isRight()).toBe(true)
      expect(result.value).toContain("invoice.pdf")
      expect(result.value).toContain("application/pdf")
      expect(result.value).toContain("20.0 KB")
      expect(result.value).toContain("/me/messages/msg-1/attachments/att-1/$value")
    })

    it("should mark inline attachments so signature images are recognisable", async () => {
      mockClient.listAttachments.mockResolvedValue(
        Right({ value: [{ id: "att-2", name: "logo.png", contentType: "image/png", size: 900, isInline: true }] }),
      )
      const result = await listAttachments({ message_id: "msg-1" })
      expect(result.value).toContain("[inline]")
      expect(result.value).toContain("900 B")
    })

    it("should report no attachments rather than an empty list", async () => {
      mockClient.listAttachments.mockResolvedValue(Right({ value: [] }))
      const result = await listAttachments({ message_id: "msg-1" })
      expect(result.isRight()).toBe(true)
      expect(result.value).toContain("No attachments found")
    })

    it("should surface a failure as a UserError", async () => {
      const { Left: L } = await import("functype/either")
      mockClient.listAttachments.mockResolvedValue(L({ message: "not found" }))
      const result = await listAttachments({ message_id: "bad" })
      expect(result.isLeft()).toBe(true)
      expect((result.value as Error).message).toContain("Failed to list attachments")
    })
  })
})

describe("scanMessages paging safety", () => {
  // Graph silently ignores $skip when $search is set: it returns page one again
  // rather than erroring. A caller paging a search would re-read the same rows while
  // believing it was advancing, then conclude the mailbox held nothing more. Failing
  // loudly is the only way that surfaces.
  it("refuses skip combined with search, and says how to page instead", async () => {
    const result = await scanMessages({ search: "invoice", skip: 100 })

    expect(result.isLeft()).toBe(true)
    const message = (result.value as { message: string }).message
    expect(message).toContain("silently")
    expect(message).toContain("received:")
    expect(mockClient.listMessages).not.toHaveBeenCalled()
  })

  it("allows skip on a filter scan, which Graph does honour", async () => {
    mockClient.listMessages.mockResolvedValue(Right({ value: [] }))

    const result = await scanMessages({ filter: "hasAttachments eq true", skip: 100 })

    expect(result.isRight()).toBe(true)
    expect(mockClient.listMessages).toHaveBeenCalled()
  })

  it("allows search on its own", async () => {
    mockClient.listMessages.mockResolvedValue(Right({ value: [] }))

    const result = await scanMessages({ search: "invoice" })

    expect(result.isRight()).toBe(true)
  })
})

describe("scan refs work across message tools", () => {
  // scan_messages returns short refs, but only get_message resolved them at first.
  // list_attachments — the tool an attachment sweep leans on hardest — rejected them
  // as malformed IDs, breaking the scan-then-open loop at exactly the wrong point.
  it("rejects an unknown ref with guidance instead of a Graph error", async () => {
    const result = await listAttachments({ message_id: "999999" })

    expect(result.isLeft()).toBe(true)
    expect((result.value as { message: string }).message).toContain("scan_messages")
    expect(mockClient.listAttachments).not.toHaveBeenCalled()
  })

  it("still passes a full Graph ID straight through", async () => {
    mockClient.listAttachments.mockResolvedValue(Right({ value: [] }))
    const graphId = "AAMkAGI0YjA3OTNhLWY2MDEtNGZlYy1hNzU2LTE4NDFiODg5ZjliMg=="

    await listAttachments({ message_id: graphId })

    expect(mockClient.listAttachments).toHaveBeenCalledWith(graphId, "/me")
  })
})

const sweepMessage = (id: string, address: string, received: string): GraphMessage => ({
  id,
  subject: `Subject ${id}`,
  from: { emailAddress: { name: "Sender", address } },
  receivedDateTime: `${received}T00:00:00Z`,
  isRead: false,
})

describe("summarizeSenders", () => {
  it("reads the whole folder without an orderby and returns counts per sender", async () => {
    mockClient.listMailFolders.mockResolvedValue(Right({ value: [] }))
    mockClient.listFolderMessagesAll.mockResolvedValue(
      Right([
        sweepMessage("1", "news@x.com", "2026-01-01"),
        sweepMessage("2", "news@x.com", "2026-01-02"),
        sweepMessage("3", "bob@y.com", "2026-01-03"),
      ]),
    )
    const result = await summarizeSenders({ mailbox: undefined })
    expect(result.isRight()).toBe(true)
    expect(result.value).toContain("3 messages in inbox")
    expect(result.value).toContain("2|2|2026-01-01|2026-01-02|news@x.com")
    const [, odata] = mockClient.listFolderMessagesAll.mock.calls[0]!
    expect(odata.$orderby).toBeUndefined()
    expect(odata.$top).toBe(999)
  })
})

describe("moveMessagesMatching", () => {
  beforeEach(() => {
    mockClient.listMailFolders.mockResolvedValue(Right({ value: [] }))
  })

  it("refuses a sweep with neither senders nor filter", async () => {
    const result = await moveMessagesMatching({ folder: "inbox", destination: "deleteditems" })
    expect(result.isLeft()).toBe(true)
    expect(mockClient.listFolderMessagesAll).not.toHaveBeenCalled()
  })

  it("builds an exact-address filter from senders, escaped for OData", async () => {
    mockClient.listFolderMessagesAll.mockResolvedValue(Right([]))
    await moveMessagesMatching({
      folder: "inbox",
      destination: "deleteditems",
      senders: ["News@X.com", "o'brien@y.com"],
      filter: "receivedDateTime lt 2025-01-01T00:00:00Z",
    })
    const [, odata] = mockClient.listFolderMessagesAll.mock.calls[0]!
    expect(odata.$filter).toBe(
      "(from/emailAddress/address eq 'news@x.com' or from/emailAddress/address eq 'o''brien@y.com') and (receivedDateTime lt 2025-01-01T00:00:00Z)",
    )
  })

  // The default must be the safe path: a caller who forgets dry_run gets a report,
  // not a moved folder.
  it("is a dry run by default and moves nothing", async () => {
    mockClient.listFolderMessagesAll.mockResolvedValue(
      Right([sweepMessage("1", "news@x.com", "2026-01-01"), sweepMessage("2", "news@x.com", "2026-01-05")]),
    )
    const result = await moveMessagesMatching({ folder: "inbox", destination: "deleteditems", senders: ["news@x.com"] })
    expect(result.isRight()).toBe(true)
    expect(result.value).toContain("2 messages in inbox match")
    expect(result.value).toContain("Oldest 2026-01-01, newest 2026-01-05")
    expect(result.value).toContain("Nothing was moved")
    expect(mockClient.batchRequest).not.toHaveBeenCalled()
  })

  it("refuses a live run above the limit instead of moving part of it", async () => {
    mockClient.listFolderMessagesAll.mockResolvedValue(
      Right([sweepMessage("1", "a@x.com", "2026-01-01"), sweepMessage("2", "a@x.com", "2026-01-02")]),
    )
    const result = await moveMessagesMatching({
      folder: "inbox",
      destination: "deleteditems",
      senders: ["a@x.com"],
      dry_run: false,
      limit: 1,
    })
    expect(result.isLeft()).toBe(true)
    expect((result.value as { message: string }).message).toContain("2 messages match, above the limit of 1")
    expect(mockClient.batchRequest).not.toHaveBeenCalled()
  })

  it("moves through $batch and reports the count", async () => {
    mockClient.listFolderMessagesAll.mockResolvedValue(
      Right([sweepMessage("1", "a@x.com", "2026-01-01"), sweepMessage("2", "a@x.com", "2026-01-02")]),
    )
    mockClient.batchRequest.mockImplementation(async (requests: ReadonlyArray<{ id: string }>) =>
      Right({ responses: requests.map((r) => ({ id: r.id, status: 201 })) }),
    )
    const result = await moveMessagesMatching({
      folder: "inbox",
      destination: "deleteditems",
      senders: ["a@x.com"],
      dry_run: false,
      mailbox: undefined,
    })
    expect(result.isRight()).toBe(true)
    expect(result.value).toBe("Moved 2/2 message(s) from inbox to deleteditems.")
    const [requests] = mockClient.batchRequest.mock.calls[0]!
    expect(requests[0]).toMatchObject({ url: "/me/messages/1/move", body: { destinationId: "deleteditems" } })
  })

  it("lists failures by subject", async () => {
    mockClient.listFolderMessagesAll.mockResolvedValue(Right([sweepMessage("1", "a@x.com", "2026-01-01")]))
    mockClient.batchRequest.mockResolvedValue(
      Right({ responses: [{ id: "0", status: 404, body: { error: { code: "ErrorItemNotFound", message: "gone" } } }] }),
    )
    const result = await moveMessagesMatching({
      folder: "inbox",
      destination: "deleteditems",
      senders: ["a@x.com"],
      dry_run: false,
    })
    expect(result.value).toContain("Moved 0/1")
    expect(result.value).toContain('FAILED "Subject 1": ErrorItemNotFound gone')
  })

  it("refuses when the destination is the folder being swept", async () => {
    mockClient.listMailFolders.mockResolvedValue(Right({ value: [{ id: "f1", displayName: "Promos" }] }))
    mockClient.listFolderMessagesAll.mockResolvedValue(Right([]))
    const result = await moveMessagesMatching({ folder: "Promos", destination: "Promos", senders: ["a@x.com"] })
    expect(result.isLeft()).toBe(true)
  })
})

describe("orderedFilter", () => {
  it("prefixes a filter that lacks the sort property so Graph accepts it with $orderby", () => {
    expect(orderedFilter("from/emailAddress/address eq 'a@x.com'")).toBe(
      "receivedDateTime ge 1900-01-01T00:00:00Z and (from/emailAddress/address eq 'a@x.com')",
    )
  })

  it("leaves a filter that already leads with receivedDateTime alone", () => {
    expect(orderedFilter("receivedDateTime ge 2026-01-01T00:00:00Z and hasAttachments eq true")).toBe(
      "receivedDateTime ge 2026-01-01T00:00:00Z and hasAttachments eq true",
    )
  })

  it("passes undefined and blank through", () => {
    expect(orderedFilter(undefined)).toBeUndefined()
    expect(orderedFilter("  ")).toBeUndefined()
  })

  it("is applied by scanMessages to a sorted scan", async () => {
    mockClient.listMessages.mockResolvedValue(Right({ value: [] }))
    await scanMessages({ filter: "hasAttachments eq true" })
    const [odata] = mockClient.listMessages.mock.calls[0]!
    expect(odata.$filter).toBe("receivedDateTime ge 1900-01-01T00:00:00Z and (hasAttachments eq true)")
    expect(odata.$orderby).toBe("receivedDateTime desc")
  })
})

describe("listMailFolders subfolder reporting", () => {
  it("should surface subfolders that the listing cannot show", async () => {
    mockClient.listMailFolders.mockResolvedValue(
      Right({ value: [{ id: "f1", displayName: "Inbox", totalItemCount: 40, childFolderCount: 2 }] }),
    )
    const result = await listMailFolders()
    expect(result.value).toContain("2 subfolders")
    expect(result.value).toContain("Top-level folders only")
  })

  it("should not mention subfolders for a folder that has none", async () => {
    mockClient.listMailFolders.mockResolvedValue(
      Right({ value: [{ id: "f2", displayName: "Archive", totalItemCount: 5, childFolderCount: 0 }] }),
    )
    const result = await listMailFolders()
    expect(result.value).not.toContain("subfolders)")
  })
})

describe("moveMessage destination reporting", () => {
  it("should name the well-known folder it resolved, not what the caller typed", async () => {
    mockClient.moveMessage.mockResolvedValue(Right({ id: "new-id", subject: "Newsletter" }))
    const result = await moveMessage({ message_id: "msg-1", destination: "junk" })
    expect(mockClient.moveMessage).toHaveBeenCalledWith("msg-1", "junkemail", "/me")
    expect(result.value).toBe('Moved "Newsletter" to the junkemail folder. New ID: new-id')
  })

  it("should name the matched folder when resolving a display name", async () => {
    mockClient.listMailFolders.mockResolvedValue(Right({ value: [{ id: "f-receipts", displayName: "Receipts" }] }))
    mockClient.moveMessage.mockResolvedValue(Right({ id: "new-id", subject: "Invoice" }))
    const result = await moveMessage({ message_id: "msg-1", destination: "receipts" })
    expect(mockClient.moveMessage).toHaveBeenCalledWith("msg-1", "f-receipts", "/me")
    expect(result.value).toBe('Moved "Invoice" to "Receipts". New ID: new-id')
  })

  it("should say it fell through to a folder ID rather than implying a name match", async () => {
    mockClient.listMailFolders.mockResolvedValue(Right({ value: [{ id: "f1", displayName: "Archive" }] }))
    mockClient.moveMessage.mockResolvedValue(Right({ id: "new-id", subject: "Contract" }))
    const result = await moveMessage({ message_id: "msg-1", destination: "AAMkAGI0-opaque" })
    expect(result.value).toBe('Moved "Contract" to folder ID AAMkAGI0-opaque. New ID: new-id')
  })

  it("should explain a typo'd folder name once Graph rejects it", async () => {
    const { Left: L } = await import("functype/either")
    mockClient.listMailFolders.mockResolvedValue(Right({ value: [{ id: "f1", displayName: "Receipts" }] }))
    mockClient.moveMessage.mockResolvedValue(L({ message: "The specified object was not found in the store." }))
    const result = await moveMessage({ message_id: "msg-1", destination: "Reciepts" })
    expect(result.isLeft()).toBe(true)
    expect((result.value as Error).message).toContain('No top-level folder is named "Reciepts"')
    expect((result.value as Error).message).toContain("list_mail_folders")
  })

  it("should not blame the folder name when a well-known move fails", async () => {
    const { Left: L } = await import("functype/either")
    mockClient.moveMessage.mockResolvedValue(L({ message: "Mailbox is unavailable." }))
    const result = await moveMessage({ message_id: "msg-1", destination: "archive" })
    expect((result.value as Error).message).toBe("Failed to move message: Mailbox is unavailable.")
  })
})

describe("listAttachments read_document paths", () => {
  it("should not offer a read_document path for a cloud link", async () => {
    mockClient.listAttachments.mockResolvedValue(
      Right({
        value: [
          {
            id: "att-3",
            name: "Renovation invoices",
            "@odata.type": "#microsoft.graph.referenceAttachment",
          },
        ],
      }),
    )
    const result = await listAttachments({ message_id: "msg-1" })
    expect(result.value).toContain("Renovation invoices")
    // Our wording is "cloud file link" / "cloud folder link" — it names the kind as well.
    expect(result.value).toMatch(/cloud (file|folder) link/)
    expect(result.value).not.toContain("read_document path")
  })

  it("should not offer a read_document path for an embedded Outlook item", async () => {
    mockClient.listAttachments.mockResolvedValue(
      Right({
        value: [{ id: "att-4", name: "Fwd: contract", "@odata.type": "#microsoft.graph.itemAttachment" }],
      }),
    )
    const result = await listAttachments({ message_id: "msg-1" })
    expect(result.value).toContain("embedded Outlook item")
    expect(result.value).not.toContain("read_document path")
  })

  it("should still offer the path for a file attachment", async () => {
    mockClient.listAttachments.mockResolvedValue(
      Right({
        value: [
          {
            id: "att-5",
            name: "scan.pdf",
            contentType: "application/pdf",
            size: 2048,
            "@odata.type": "#microsoft.graph.fileAttachment",
          },
        ],
      }),
    )
    const result = await listAttachments({ message_id: "msg-1" })
    expect(result.value).toContain("/me/messages/msg-1/attachments/att-5/$value")
  })

  it("should report an unknown size rather than claiming zero bytes", async () => {
    mockClient.listAttachments.mockResolvedValue(Right({ value: [{ id: "att-6", name: "mystery.bin" }] }))
    const result = await listAttachments({ message_id: "msg-1" })
    expect(result.value).toContain("unknown size")
  })
})

describe("batchMoveMessages failure reporting", () => {
  it("should fail the call when no message moved at all", async () => {
    const { Left: L } = await import("functype/either")
    mockClient.moveMessage.mockResolvedValue(L({ type: "api", message: "boom" }))
    const result = await batchMoveMessages({ message_ids: ["a", "b"], destination: "archive" })
    expect(result.isLeft()).toBe(true)
    expect((result.value as Error).message).toContain("Moved 0/2")
  })

  it("should still succeed when some moved, since the caller needs that list", async () => {
    const { Left: L } = await import("functype/either")
    mockClient.moveMessage
      .mockResolvedValueOnce(Right({ id: "n1", subject: "ok" }))
      .mockResolvedValueOnce(L({ type: "api", message: "boom" }))
    const result = await batchMoveMessages({ message_ids: ["a", "b"], destination: "archive" })
    expect(result.isRight()).toBe(true)
    expect(result.value).toContain("Moved 1/2")
  })

  it("should stop at the first throttle instead of burning the rest of the batch", async () => {
    const { Left: L } = await import("functype/either")
    mockClient.moveMessage
      .mockResolvedValueOnce(Right({ id: "n1", subject: "ok" }))
      .mockResolvedValueOnce(L({ type: "throttle", message: "Too many requests", status: 429 }))
    const result = await batchMoveMessages({ message_ids: ["a", "b", "c", "d"], destination: "archive" })
    // Two calls attempted: the success, then the throttle. c and d are never tried.
    expect(mockClient.moveMessage).toHaveBeenCalledTimes(2)
    expect(result.value).toContain("Moved 1/4")
    // The throttled message failed; the two after it were never sent. Counting all three as
    // failures would overstate the damage and hide that c and d are still safe to retry.
    expect(result.value).toContain("1 failed:")
    expect(result.value).toContain("2 not attempted:")
    expect(result.value).toContain("NOT ATTEMPTED c")
    expect(result.value).toContain("NOT ATTEMPTED d")
  })

  it("should name a destination that was only assumed to be a folder ID", async () => {
    mockClient.listMailFolders.mockResolvedValue(Right({ value: [{ id: "f1", displayName: "Receipts" }] }))
    mockClient.moveMessage.mockResolvedValue(Right({ id: "n", subject: "s" }))
    const ok = await batchMoveMessages({ message_ids: ["a"], destination: "Receipts" })
    expect(ok.value).not.toContain("used as a folder ID")

    const { Left: L } = await import("functype/either")
    mockClient.moveMessage.mockResolvedValue(L({ type: "api", message: "not found in store" }))
    const bad = await batchMoveMessages({ message_ids: ["a"], destination: "Reciepts" })
    expect((bad.value as Error).message).toContain('No top-level folder is named "Reciepts"')
  })

  it("should report the resolved folder in the summary, not what was typed", async () => {
    mockClient.moveMessage.mockResolvedValue(Right({ id: "n", subject: "s" }))
    const result = await batchMoveMessages({ message_ids: ["a"], destination: "junk" })
    expect(mockClient.moveMessage).toHaveBeenCalledWith("a", "junkemail", "/me")
    expect(result.value).toContain("to the junkemail folder")
  })
})
