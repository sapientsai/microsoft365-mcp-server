import { Some } from "functype"
import { Left, Right } from "functype/either"
import { beforeEach, describe, expect, it, vi } from "vitest"

vi.mock("../src/client/graph-client", () => ({
  getGraphClient: vi.fn(),
}))

import { getGraphClient } from "../src/client/graph-client"
import { respondToEvent } from "../src/tools/calendar-tools"
import { sendChatMessage } from "../src/tools/chat-tools"
import { deleteDraft, updateDraft } from "../src/tools/mail-tools"
import { sendChannelMessage } from "../src/tools/teams-tools"
import { formatChannelMessageSummary } from "../src/utils/formatters"

const mockClient = {
  getDraftState: vi.fn(),
  updateMessage: vi.fn(),
  deleteMessage: vi.fn(),
  respondToEvent: vi.fn(),
  sendChannelMessage: vi.fn(),
  replyToChannelMessage: vi.fn(),
  getMe: vi.fn(),
  createChat: vi.fn(),
  sendChatMessage: vi.fn(),
}

const message = (result: { fold: (l: (e: Error) => string, r: (v: string) => string) => string }): string =>
  result.fold(
    (e) => e.message,
    (v) => v,
  )

const ME = { id: "me-id", mail: "jordan@civala.com", userPrincipalName: "jordan@civala.com" }

beforeEach(() => {
  vi.clearAllMocks()
  vi.mocked(getGraphClient).mockReturnValue(Some(mockClient as never))
})

describe("updateDraft", () => {
  it("should send only the fields passed", async () => {
    mockClient.getDraftState.mockResolvedValue(Right({ id: "d1", isDraft: true }))
    mockClient.updateMessage.mockResolvedValue(Right({ id: "d1", subject: "New" }))

    const result = await updateDraft({ message_id: "d1", subject: "New", body: "Hi", cc: "a@x.com, b@x.com" })

    expect(result.isRight()).toBe(true)
    expect(mockClient.updateMessage).toHaveBeenCalledWith("d1", {
      subject: "New",
      body: { contentType: "Text", content: "Hi" },
      ccRecipients: [{ emailAddress: { address: "a@x.com" } }, { emailAddress: { address: "b@x.com" } }],
    })
  })

  it("should clear a recipient list given an empty string", async () => {
    mockClient.getDraftState.mockResolvedValue(Right({ id: "d1", isDraft: true }))
    mockClient.updateMessage.mockResolvedValue(Right({ id: "d1" }))

    await updateDraft({ message_id: "d1", bcc: "" })

    expect(mockClient.updateMessage).toHaveBeenCalledWith("d1", { bccRecipients: [] })
  })

  // PATCH on a received message can still change some properties, so the draft check is the guard.
  it("should refuse a message that is not a draft, without changing it", async () => {
    mockClient.getDraftState.mockResolvedValue(Right({ id: "m1", isDraft: false }))

    const result = await updateDraft({ message_id: "m1", subject: "Changed" })

    expect(message(result)).toContain("is not a draft")
    expect(mockClient.updateMessage).not.toHaveBeenCalled()
  })

  it("should refuse a call with nothing to change, without reading the message", async () => {
    const result = await updateDraft({ message_id: "d1" })

    expect(message(result)).toContain("Nothing to change")
    expect(mockClient.getDraftState).not.toHaveBeenCalled()
  })

  it("should report a failed read instead of guessing", async () => {
    mockClient.getDraftState.mockResolvedValue(Left({ message: "Not found" }))

    const result = await updateDraft({ message_id: "gone", subject: "x" })

    expect(message(result)).toContain("Failed to read message: Not found")
    expect(mockClient.updateMessage).not.toHaveBeenCalled()
  })
})

describe("deleteDraft", () => {
  it("should delete a draft and name it", async () => {
    mockClient.getDraftState.mockResolvedValue(Right({ id: "d1", isDraft: true, subject: "Quarterly update" }))
    mockClient.deleteMessage.mockResolvedValue(Right({}))

    const result = await deleteDraft({ message_id: "d1" })

    expect(message(result)).toBe('Draft "Quarterly update" deleted.')
    expect(mockClient.deleteMessage).toHaveBeenCalledWith("d1")
  })

  // DELETE works on any message, so without the check this would be a general delete tool.
  it("should refuse a message that is not a draft, without deleting it", async () => {
    mockClient.getDraftState.mockResolvedValue(Right({ id: "m1", isDraft: false }))

    const result = await deleteDraft({ message_id: "m1" })

    expect(message(result)).toContain("is not a draft")
    expect(mockClient.deleteMessage).not.toHaveBeenCalled()
  })

  it("should treat a missing isDraft as not a draft", async () => {
    mockClient.getDraftState.mockResolvedValue(Right({ id: "m1" }))

    const result = await deleteDraft({ message_id: "m1" })

    expect(result.isLeft()).toBe(true)
    expect(mockClient.deleteMessage).not.toHaveBeenCalled()
  })
})

describe("respondToEvent", () => {
  it.each([
    ["accept", "accept", "Accepted"],
    ["decline", "decline", "Declined"],
    ["tentative", "tentativelyAccept", "Tentatively accepted"],
  ] as const)("should map %s to the %s action", async (response, action, word) => {
    mockClient.respondToEvent.mockResolvedValue(Right({}))

    const result = await respondToEvent({ event_id: "e1", response })

    expect(mockClient.respondToEvent).toHaveBeenCalledWith("e1", action, { sendResponse: true })
    expect(message(result)).toBe(`${word}. A response was sent to the organizer.`)
  })

  it("should pass a comment and a suppressed response through", async () => {
    mockClient.respondToEvent.mockResolvedValue(Right({}))

    const result = await respondToEvent({
      event_id: "e1",
      response: "decline",
      comment: "Out that week",
      send_response: false,
    })

    expect(mockClient.respondToEvent).toHaveBeenCalledWith("e1", "decline", {
      comment: "Out that week",
      sendResponse: false,
    })
    expect(message(result)).toContain("No response was sent")
  })
})

describe("sendChannelMessage", () => {
  it("should start a new thread without reply_to_message_id", async () => {
    mockClient.sendChannelMessage.mockResolvedValue(Right({ id: "m1" }))

    await sendChannelMessage({ team_id: "t", channel_id: "c", content: "Hi" })

    expect(mockClient.sendChannelMessage).toHaveBeenCalledWith("t", "c", "Hi")
    expect(mockClient.replyToChannelMessage).not.toHaveBeenCalled()
  })

  it("should reply in the thread with reply_to_message_id", async () => {
    mockClient.replyToChannelMessage.mockResolvedValue(Right({ id: "r1" }))

    const result = await sendChannelMessage({
      team_id: "t",
      channel_id: "c",
      content: "Agreed",
      reply_to_message_id: "m1",
    })

    expect(mockClient.replyToChannelMessage).toHaveBeenCalledWith("t", "c", "m1", "Agreed")
    expect(mockClient.sendChannelMessage).not.toHaveBeenCalled()
    expect(message(result)).toBe("Reply posted in the thread.")
  })
})

describe("formatChannelMessageSummary", () => {
  // Without the ID a caller has nothing to pass as reply_to_message_id.
  it("should include the message ID", () => {
    const line = formatChannelMessageSummary({
      id: "1690000000000",
      createdDateTime: "2026-10-01T09:00:00Z",
      from: { user: { displayName: "Ada" } },
      body: { content: "Kickoff notes" },
    })
    expect(line).toBe("- **Ada** (2026-10-01T09:00:00Z): Kickoff notes (ID: 1690000000000)")
  })
})

describe("sendChatMessage", () => {
  it("should send in an existing chat with chat_id", async () => {
    mockClient.sendChatMessage.mockResolvedValue(Right({ id: "m1" }))

    const result = await sendChatMessage({ chat_id: "19:abc", content: "Hi" })

    expect(mockClient.sendChatMessage).toHaveBeenCalledWith("19:abc", "Hi")
    expect(mockClient.createChat).not.toHaveBeenCalled()
    expect(message(result)).toBe("Chat message sent.")
  })

  it.each([[{ content: "Hi" }], [{ chat_id: "19:abc", to: "ada@civala.com", content: "Hi" }]])(
    "should require exactly one of chat_id or to",
    async (params) => {
      const result = await sendChatMessage(params)

      expect(message(result)).toContain("exactly one of chat_id")
      expect(mockClient.sendChatMessage).not.toHaveBeenCalled()
    },
  )

  it("should start or reuse a one-on-one chat for one person, listing the caller first", async () => {
    mockClient.getMe.mockResolvedValue(Right(ME))
    mockClient.createChat.mockResolvedValue(Right({ id: "19:pair", chatType: "oneOnOne" }))
    mockClient.sendChatMessage.mockResolvedValue(Right({ id: "m1" }))

    const result = await sendChatMessage({ to: "ada@civala.com", content: "Hi" })

    expect(mockClient.createChat).toHaveBeenCalledWith({
      chatType: "oneOnOne",
      members: [
        {
          "@odata.type": "#microsoft.graph.aadUserConversationMember",
          roles: ["owner"],
          "user@odata.bind": "https://graph.microsoft.com/v1.0/users('me-id')",
        },
        {
          "@odata.type": "#microsoft.graph.aadUserConversationMember",
          roles: ["owner"],
          "user@odata.bind": "https://graph.microsoft.com/v1.0/users('ada@civala.com')",
        },
      ],
    })
    expect(mockClient.sendChatMessage).toHaveBeenCalledWith("19:pair", "Hi")
    expect(message(result)).toBe("Chat message sent. Chat ID: 19:pair.")
  })

  // Listing the caller twice fails the create, and a group of "me plus one" should be a 1:1.
  it("should drop the caller's own address and duplicates from to", async () => {
    mockClient.getMe.mockResolvedValue(Right(ME))
    mockClient.createChat.mockResolvedValue(Right({ id: "19:pair", chatType: "oneOnOne" }))
    mockClient.sendChatMessage.mockResolvedValue(Right({ id: "m1" }))

    await sendChatMessage({ to: "Jordan@Civala.com, ada@civala.com, ADA@civala.com", content: "Hi" })

    const chat = mockClient.createChat.mock.calls[0]?.[0] as { chatType: string; members: ReadonlyArray<unknown> }
    expect(chat.chatType).toBe("oneOnOne")
    expect(chat.members).toHaveLength(2)
  })

  it("should point the caller at the self-chat when to names only themselves", async () => {
    mockClient.getMe.mockResolvedValue(Right(ME))

    const result = await sendChatMessage({ to: "jordan@civala.com", content: "note" })

    expect(message(result)).toContain("48:notes")
    expect(mockClient.createChat).not.toHaveBeenCalled()
  })

  it("should start a group chat with a topic and say to reuse its ID", async () => {
    mockClient.getMe.mockResolvedValue(Right(ME))
    mockClient.createChat.mockResolvedValue(Right({ id: "19:group", chatType: "group" }))
    mockClient.sendChatMessage.mockResolvedValue(Right({ id: "m1" }))

    const result = await sendChatMessage({ to: "ada@civala.com, bo@civala.com", topic: "Launch", content: "Hi all" })

    const chat = mockClient.createChat.mock.calls[0]?.[0] as { chatType: string; topic?: string; members: unknown[] }
    expect(chat.chatType).toBe("group")
    expect(chat.topic).toBe("Launch")
    expect(chat.members).toHaveLength(3)
    expect(message(result)).toContain("Started a new group chat (chat ID: 19:group)")
    expect(message(result)).toContain("Use this chat_id for follow-ups")
  })

  it("should refuse a topic on a one-on-one chat", async () => {
    mockClient.getMe.mockResolvedValue(Right(ME))

    const result = await sendChatMessage({ to: "ada@civala.com", topic: "Launch", content: "Hi" })

    expect(message(result)).toContain("only be set on a group chat")
    expect(mockClient.createChat).not.toHaveBeenCalled()
  })

  it("should refuse a topic alongside chat_id", async () => {
    const result = await sendChatMessage({ chat_id: "19:abc", topic: "Launch", content: "Hi" })

    expect(message(result)).toContain("topic only applies")
    expect(mockClient.sendChatMessage).not.toHaveBeenCalled()
  })

  it("should escape a quote in an address for the OData bind", async () => {
    mockClient.getMe.mockResolvedValue(Right(ME))
    mockClient.createChat.mockResolvedValue(Right({ id: "19:pair", chatType: "oneOnOne" }))
    mockClient.sendChatMessage.mockResolvedValue(Right({ id: "m1" }))

    await sendChatMessage({ to: "o'brien@civala.com", content: "Hi" })

    const chat = mockClient.createChat.mock.calls[0]?.[0] as { members: ReadonlyArray<Record<string, unknown>> }
    expect(chat.members[1]?.["user@odata.bind"]).toBe("https://graph.microsoft.com/v1.0/users('o''brien@civala.com')")
  })

  // A retry with 'to' would make a second group chat, so the ID must survive a failed send.
  it("should report the chat ID when the chat was created but the send failed", async () => {
    mockClient.getMe.mockResolvedValue(Right(ME))
    mockClient.createChat.mockResolvedValue(Right({ id: "19:group", chatType: "group" }))
    mockClient.sendChatMessage.mockResolvedValue(Left({ message: "Forbidden" }))

    const result = await sendChatMessage({ to: "ada@civala.com, bo@civala.com", content: "Hi" })

    expect(message(result)).toBe("Chat is ready (chat ID: 19:group), but sending failed: Forbidden")
  })

  it("should pass Graph's error through when the chat can't be created", async () => {
    mockClient.getMe.mockResolvedValue(Right(ME))
    mockClient.createChat.mockResolvedValue(Left({ message: "User not found" }))

    const result = await sendChatMessage({ to: "nobody@civala.com", content: "Hi" })

    expect(message(result)).toBe("Failed to start chat: User not found")
    expect(mockClient.sendChatMessage).not.toHaveBeenCalled()
  })
})
