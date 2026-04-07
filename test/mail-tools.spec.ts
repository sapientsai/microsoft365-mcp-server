import { Some } from "functype"
import { Right } from "functype/either"
import { beforeEach, describe, expect, it, vi } from "vitest"

import type { GraphMessage } from "../src/types"

vi.mock("../src/client/graph-client", () => ({
  getGraphClient: vi.fn(),
}))

import { getGraphClient } from "../src/client/graph-client"
import { createDraft, sendDraft, sendMessage } from "../src/tools/mail-tools"

const mockClient = {
  sendMessage: vi.fn(),
  createDraft: vi.fn(),
  sendDraft: vi.fn(),
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
})
