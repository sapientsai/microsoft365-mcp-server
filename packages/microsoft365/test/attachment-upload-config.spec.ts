import { resolveUploadTicket } from "@sapientsai/ms-graph-core"
import { Some } from "functype"
import { Left, Right } from "functype/either"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

vi.mock("../src/client/graph-client", () => ({ getGraphClient: vi.fn() }))
vi.mock("../src/auth", async (original) => ({
  ...(await original<typeof import("../src/auth")>()),
  getAccessToken: vi.fn(async () => Right("DELEGATED.JWT")),
}))

import { getAccessToken } from "../src/auth"
import { getGraphClient } from "../src/client/graph-client"
import { NO_UPLOAD_TICKET_MESSAGE } from "../src/tools/files-tools"
import { getAttachmentUploadConfig } from "../src/tools/mail-tools"
import { rememberMessageId } from "../src/utils/message-refs"

const mockClient = { getDraftState: vi.fn(), request: vi.fn() }

beforeEach(() => {
  vi.clearAllMocks()
  vi.mocked(getGraphClient).mockReturnValue(Some(mockClient as never))
  mockClient.request.mockResolvedValue(Left({ type: "api", message: "own-address lookup not under test" }))
  vi.stubEnv("MS365_PUBLIC_BASE_URL", "https://ms365.example/")
})
afterEach(() => vi.unstubAllEnvs())

describe("getAttachmentUploadConfig", () => {
  it("returns a relay URL for the draft, with an opaque ticket instead of the Graph token", async () => {
    mockClient.getDraftState.mockResolvedValue(Right({ id: "AAMkDraftMessage", isDraft: true }))

    const result = await getAttachmentUploadConfig({
      message_id: "AAMkDraftMessage",
      name: " expenses.pdf ",
      localFile: "/tmp/e.pdf",
    })

    expect(result.isRight()).toBe(true)
    const config = JSON.parse(result.value as string) as { uploadUrl: string; authHeader: string; curl: string }
    const url = new URL(config.uploadUrl)
    expect(`${url.origin}${url.pathname}`).toBe("https://ms365.example/upload/attachment")
    expect(Object.fromEntries(url.searchParams)).toEqual({
      messageId: "AAMkDraftMessage",
      name: "expenses.pdf",
      contentType: "application/pdf",
      encoding: "base64",
    })
    expect(config.authHeader).not.toContain("DELEGATED.JWT")
    expect(resolveUploadTicket(config.authHeader.replace("Authorization: Bearer ", ""))).toBe("DELEGATED.JWT")
    expect(config.curl).toContain('base64 "/tmp/e.pdf"')
  })

  // Without a Graph token no ticket can be minted, and a curl without one can only be refused.
  // The shared upload secret is never handed out in its place: it never expires.
  it("refuses with a clear error, not a curl that can only fail, when no Graph token resolves", async () => {
    mockClient.getDraftState.mockResolvedValue(Right({ id: "AAMkDraftMessage", isDraft: true }))
    vi.stubEnv("MS365_UPLOAD_TOKEN", "made-up-shared-secret")
    vi.mocked(getAccessToken).mockResolvedValueOnce(Left({ type: "auth", message: "not signed in" }))

    const result = await getAttachmentUploadConfig({ message_id: "AAMkDraftMessage", name: "a.pdf" })

    expect(result.isLeft()).toBe(true)
    expect((result.value as Error).message).toBe(NO_UPLOAD_TICKET_MESSAGE)
    expect((result.value as Error).message).not.toContain("made-up-shared-secret")
  })

  it("resolves a scan ref to the real message id", async () => {
    mockClient.getDraftState.mockResolvedValue(Right({ id: "AAMkLongGraphId", isDraft: true }))
    const ref = rememberMessageId("AAMkLongGraphId")

    const result = await getAttachmentUploadConfig({ message_id: ref, name: "a.pdf" })

    expect(mockClient.getDraftState).toHaveBeenCalledWith("AAMkLongGraphId", "/me")
    expect(
      new URL((JSON.parse(result.value as string) as { uploadUrl: string }).uploadUrl).searchParams.get("messageId"),
    ).toBe("AAMkLongGraphId")
  })

  it("refuses a message that is not a draft", async () => {
    mockClient.getDraftState.mockResolvedValue(Right({ id: "AAMkSentMessage", isDraft: false }))

    const result = await getAttachmentUploadConfig({ message_id: "AAMkSentMessage", name: "a.pdf" })

    expect(result.isLeft()).toBe(true)
    const message = (result.value as Error).message
    expect(message).toContain("is not a draft. Attachments can only be added to drafts")
    expect(message).not.toContain("edited or deleted")
  })

  it("requires a file name before reading the message", async () => {
    const result = await getAttachmentUploadConfig({ message_id: "AAMkDraftMessage", name: "  " })

    expect(result.isLeft()).toBe(true)
    expect(mockClient.getDraftState).not.toHaveBeenCalled()
  })

  // A draft created in a shared mailbox lives there, so it must be read and attached there too.
  describe("in another mailbox", () => {
    beforeEach(() => vi.stubEnv("MS365_ALLOWED_MAILBOXES", "shared@example.com"))

    it("checks the draft in that mailbox and carries it into the relay URL", async () => {
      mockClient.getDraftState.mockResolvedValue(Right({ id: "AAMkShared", isDraft: true }))

      const result = await getAttachmentUploadConfig({
        message_id: "AAMkShared",
        name: "a.pdf",
        mailbox: "Shared@Example.com",
      })

      expect(mockClient.getDraftState).toHaveBeenCalledWith("AAMkShared", "/users/shared%40example.com")
      const url = new URL((JSON.parse(result.value as string) as { uploadUrl: string }).uploadUrl)
      expect(url.searchParams.get("mailbox")).toBe("shared@example.com")
    })

    it("leaves mailbox out of the relay URL for your own mailbox", async () => {
      mockClient.getDraftState.mockResolvedValue(Right({ id: "AAMkDraftMessage", isDraft: true }))

      const result = await getAttachmentUploadConfig({ message_id: "AAMkDraftMessage", name: "a.pdf" })

      expect(
        new URL((JSON.parse(result.value as string) as { uploadUrl: string }).uploadUrl).searchParams.has("mailbox"),
      ).toBe(false)
    })

    it("refuses an unpermitted mailbox before reading the message", async () => {
      const result = await getAttachmentUploadConfig({
        message_id: "AAMkX",
        name: "a.pdf",
        mailbox: "other@example.com",
      })

      expect((result.value as Error).message).toContain("is not permitted")
      expect(mockClient.getDraftState).not.toHaveBeenCalled()
    })
  })
})
