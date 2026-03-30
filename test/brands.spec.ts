import { describe, expect, it } from "vitest"

import { ChannelId, ContactId, EventId, MessageId, TeamId, TodoListId, TodoTaskId, UserId } from "../src/brands"

describe("Brand types", () => {
  it("should create branded IDs", () => {
    const userId = UserId("user-123")
    const messageId = MessageId("msg-456")
    const eventId = EventId("evt-789")
    const contactId = ContactId("contact-abc")
    const teamId = TeamId("team-def")
    const channelId = ChannelId("channel-ghi")
    const todoListId = TodoListId("list-jkl")
    const todoTaskId = TodoTaskId("task-mno")

    // Branded types are still strings at runtime
    expect(String(userId)).toBe("user-123")
    expect(String(messageId)).toBe("msg-456")
    expect(String(eventId)).toBe("evt-789")
    expect(String(contactId)).toBe("contact-abc")
    expect(String(teamId)).toBe("team-def")
    expect(String(channelId)).toBe("channel-ghi")
    expect(String(todoListId)).toBe("list-jkl")
    expect(String(todoTaskId)).toBe("task-mno")
  })
})
