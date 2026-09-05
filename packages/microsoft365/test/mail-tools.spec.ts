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
  createReplyDraft,
  listMailFolders,
  moveMessage,
  listAttachments,
  sendDraft,
  sendForward,
  sendMessage,
  sendReply,
  sendReplyAll,
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
  listMailFolders: vi.fn(),
  moveMessage: vi.fn(),
  requestPaginated: vi.fn(),
  listAttachments: vi.fn(),
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
      expect(mockClient.sendMessage).toHaveBeenCalledWith({
        message: {
          subject: "Hi",
          body: { contentType: "Text", content: "Hello" },
          toRecipients: [{ emailAddress: { address: "alice@example.com" } }],
        },
      })
    })

    it("should send a message with HTML content type", async () => {
      mockClient.sendMessage.mockResolvedValue(Right({}))
      await sendMessage({ to: "bob@example.com", subject: "Hi", body: "<b>Bold</b>", content_type: "HTML" })
      expect(mockClient.sendMessage).toHaveBeenCalledWith({
        message: {
          subject: "Hi",
          body: { contentType: "HTML", content: "<b>Bold</b>" },
          toRecipients: [{ emailAddress: { address: "bob@example.com" } }],
        },
      })
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
      expect(mockClient.createDraft).toHaveBeenCalledWith({
        subject: "Draft",
        body: { contentType: "Text", content: "Content" },
        toRecipients: [{ emailAddress: { address: "alice@example.com" } }],
      })
    })

    it("should create a draft with HTML content type", async () => {
      mockClient.createDraft.mockResolvedValue(Right(draftResponse))
      await createDraft({ to: "alice@example.com", subject: "Draft", body: "<p>Hi</p>", content_type: "HTML" })
      expect(mockClient.createDraft).toHaveBeenCalledWith(
        expect.objectContaining({
          body: { contentType: "HTML", content: "<p>Hi</p>" },
        }),
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
      )
    })

    it("should create a draft with bcc recipients", async () => {
      mockClient.createDraft.mockResolvedValue(Right(draftResponse))
      await createDraft({ to: "alice@example.com", subject: "Draft", body: "Hi", bcc: "secret@example.com" })
      expect(mockClient.createDraft).toHaveBeenCalledWith(
        expect.objectContaining({
          bccRecipients: [{ emailAddress: { address: "secret@example.com" } }],
        }),
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
      expect(mockClient.sendDraft).toHaveBeenCalledWith("draft-123")
    })
  })

  describe("sendReply", () => {
    it("should send a reply by message ID", async () => {
      mockClient.sendReply.mockResolvedValue(Right({}))
      const result = await sendReply({ message_id: "msg-1", comment: "Thanks!" })
      expect(result.isRight()).toBe(true)
      expect(result.value).toContain("Reply sent successfully")
      expect(mockClient.sendReply).toHaveBeenCalledWith("msg-1", "Thanks!")
    })
  })

  describe("sendReplyAll", () => {
    it("should send a reply-all by message ID", async () => {
      mockClient.sendReplyAll.mockResolvedValue(Right({}))
      const result = await sendReplyAll({ message_id: "msg-1", comment: "Thanks all!" })
      expect(result.isRight()).toBe(true)
      expect(result.value).toContain("Reply-all sent successfully")
      expect(mockClient.sendReplyAll).toHaveBeenCalledWith("msg-1", "Thanks all!")
    })
  })

  describe("sendForward", () => {
    it("should forward with recipients and an optional comment", async () => {
      mockClient.sendForward.mockResolvedValue(Right({}))
      const result = await sendForward({ message_id: "msg-1", to: "alice@example.com", comment: "FYI" })
      expect(result.isRight()).toBe(true)
      expect(result.value).toContain("alice@example.com")
      expect(mockClient.sendForward).toHaveBeenCalledWith("msg-1", "FYI", [
        { emailAddress: { address: "alice@example.com" } },
      ])
    })

    it("should default an omitted comment to an empty string", async () => {
      mockClient.sendForward.mockResolvedValue(Right({}))
      await sendForward({ message_id: "msg-1", to: "alice@example.com" })
      expect(mockClient.sendForward).toHaveBeenCalledWith("msg-1", "", [
        { emailAddress: { address: "alice@example.com" } },
      ])
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
      expect(mockClient.createReplyDraft).toHaveBeenCalledWith("msg-1", "Will do")
    })
  })

  describe("createReplyAllDraft", () => {
    it("should create a threaded reply-all draft and return its ID", async () => {
      mockClient.createReplyAllDraft.mockResolvedValue(Right({ id: "draft-ra1" }))
      const result = await createReplyAllDraft({ message_id: "msg-1", comment: "Will do" })
      expect(result.isRight()).toBe(true)
      expect(result.value).toContain("draft-ra1")
      expect(mockClient.createReplyAllDraft).toHaveBeenCalledWith("msg-1", "Will do")
    })
  })

  describe("createForwardDraft", () => {
    it("should create a forward draft with recipients and return its ID", async () => {
      mockClient.createForwardDraft.mockResolvedValue(Right({ id: "draft-f1" }))
      const result = await createForwardDraft({ message_id: "msg-1", to: "alice@example.com", comment: "FYI" })
      expect(result.isRight()).toBe(true)
      expect(result.value).toContain("draft-f1")
      expect(mockClient.createForwardDraft).toHaveBeenCalledWith("msg-1", "FYI", [
        { emailAddress: { address: "alice@example.com" } },
      ])
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
      expect(mockClient.listMailFolders).toHaveBeenCalledWith({ $top: 100 })
    })

    // Verified against live Graph: /me/mailFolders returns immediate children of the root only.
    // A real Inbox reported childFolderCount 2 while neither subfolder appeared in the response.
    // Without the count printed, those folders are invisible and unreachable by name.
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
      expect(mockClient.moveMessage).toHaveBeenCalledWith("msg-1", "archive")
      expect(mockClient.listMailFolders).not.toHaveBeenCalled()
    })

    // The bug this guards: "junk" is both a well-known alias and a legal display name for a
    // custom folder. The alias wins, so a mailbox with a folder called "Junk" has its message
    // filed into Junk Email instead — a different folder. Echoing params.destination back said
    // "to junk" either way and hid which one happened.
    it("should name the well-known folder it resolved, not what the caller typed", async () => {
      mockClient.moveMessage.mockResolvedValue(Right({ id: "new-id", subject: "Newsletter" }))
      const result = await moveMessage({ message_id: "msg-1", destination: "junk" })
      expect(mockClient.moveMessage).toHaveBeenCalledWith("msg-1", "junkemail")
      expect(result.value).toBe('Moved "Newsletter" to the junkemail folder. New ID: new-id')
    })

    it("should name the matched folder when resolving a display name", async () => {
      mockClient.listMailFolders.mockResolvedValue(Right({ value: [{ id: "f-receipts", displayName: "Receipts" }] }))
      mockClient.moveMessage.mockResolvedValue(Right({ id: "new-id", subject: "Invoice" }))
      const result = await moveMessage({ message_id: "msg-1", destination: "receipts" })
      expect(mockClient.moveMessage).toHaveBeenCalledWith("msg-1", "f-receipts")
      expect(result.value).toBe('Moved "Invoice" to "Receipts". New ID: new-id')
    })

    it("should say it fell through to a folder ID rather than implying a name match", async () => {
      mockClient.listMailFolders.mockResolvedValue(Right({ value: [{ id: "f1", displayName: "Archive" }] }))
      mockClient.moveMessage.mockResolvedValue(Right({ id: "new-id", subject: "Contract" }))
      const result = await moveMessage({ message_id: "msg-1", destination: "AAMkAGI0-opaque" })
      expect(result.value).toBe('Moved "Contract" to folder ID AAMkAGI0-opaque. New ID: new-id')
    })

    // A typo'd folder name and a real folder ID are indistinguishable before the call. After Graph
    // has rejected it they are not, so the failure explains itself instead of surfacing an opaque
    // store error.
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
      expect(mockClient.moveMessage).toHaveBeenCalledWith("msg-1", "deleteditems")
    })

    it("should resolve a folder display name to its ID", async () => {
      mockClient.listMailFolders.mockResolvedValue(Right({ value: [{ id: "f-receipts", displayName: "Receipts" }] }))
      mockClient.moveMessage.mockResolvedValue(Right({ id: "msg-1" }))
      await moveMessage({ message_id: "msg-1", destination: "Receipts" })
      expect(mockClient.moveMessage).toHaveBeenCalledWith("msg-1", "f-receipts")
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
      expect(mockClient.moveMessage).toHaveBeenCalledWith("msg-1", "AAMkAGI0-opaque-id")
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
      // formatBytes is the shared formatter read_document and the file tools already use, so an
      // attachment reports the same size in the listing as it does once opened.
      expect(result.value).toContain("900 B")
    })

    it("should report no attachments rather than an empty list", async () => {
      mockClient.listAttachments.mockResolvedValue(Right({ value: [] }))
      const result = await listAttachments({ message_id: "msg-1" })
      expect(result.isRight()).toBe(true)
      expect(result.value).toContain("No attachments found")
    })

    // A referenceAttachment is a cloud link and an itemAttachment is an embedded Outlook item.
    // Neither has bytes on the /$value stream that read_document reads, so advertising the path for
    // them hands the caller a read that fails opaquely. Verified against live Graph: @odata.type is
    // returned on every attachment even when it is not $select-ed, so this is always available.
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
      expect(result.value).toContain("cloud link")
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

    it("should surface a failure as a UserError", async () => {
      const { Left: L } = await import("functype/either")
      mockClient.listAttachments.mockResolvedValue(L({ message: "not found" }))
      const result = await listAttachments({ message_id: "bad" })
      expect(result.isLeft()).toBe(true)
      expect((result.value as Error).message).toContain("Failed to list attachments")
    })
  })
})
