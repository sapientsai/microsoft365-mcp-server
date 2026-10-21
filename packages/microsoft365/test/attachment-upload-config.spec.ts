import { resolveUploadTicket } from "@sapientsai/ms-graph-core"
import { Some } from "functype"
import { Right } from "functype/either"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

vi.mock("../src/client/graph-client", () => ({ getGraphClient: vi.fn() }))
vi.mock("../src/auth", async (original) => ({
  ...(await original<typeof import("../src/auth")>()),
  getAccessToken: vi.fn(async () => Right("DELEGATED.JWT")),
}))

import { getGraphClient } from "../src/client/graph-client"
import { getAttachmentUploadConfig } from "../src/tools/mail-tools"
import { rememberMessageId } from "../src/utils/message-refs"

const mockClient = { getDraftState: vi.fn() }

beforeEach(() => {
  vi.clearAllMocks()
  vi.mocked(getGraphClient).mockReturnValue(Some(mockClient as never))
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

  it("resolves a scan ref to the real message id", async () => {
    mockClient.getDraftState.mockResolvedValue(Right({ id: "AAMkLongGraphId", isDraft: true }))
    const ref = rememberMessageId("AAMkLongGraphId")

    const result = await getAttachmentUploadConfig({ message_id: ref, name: "a.pdf" })

    expect(mockClient.getDraftState).toHaveBeenCalledWith("AAMkLongGraphId")
    expect(
      new URL((JSON.parse(result.value as string) as { uploadUrl: string }).uploadUrl).searchParams.get("messageId"),
    ).toBe("AAMkLongGraphId")
  })

  it("refuses a message that is not a draft", async () => {
    mockClient.getDraftState.mockResolvedValue(Right({ id: "AAMkSentMessage", isDraft: false }))

    const result = await getAttachmentUploadConfig({ message_id: "AAMkSentMessage", name: "a.pdf" })

    expect(result.isLeft()).toBe(true)
    expect((result.value as Error).message).toContain("is not a draft")
  })

  it("requires a file name before reading the message", async () => {
    const result = await getAttachmentUploadConfig({ message_id: "AAMkDraftMessage", name: "  " })

    expect(result.isLeft()).toBe(true)
    expect(mockClient.getDraftState).not.toHaveBeenCalled()
  })
})
