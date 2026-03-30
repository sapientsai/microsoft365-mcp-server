import { describe, expect, it } from "vitest"

import type { GraphEvent, GraphMessage, GraphTodoTask, GraphUser } from "../src/types"
import {
  formatEventDetail,
  formatEventList,
  formatMessageDetail,
  formatMessageList,
  formatTodoTaskDetail,
  formatTodoTaskList,
  formatUserDetail,
} from "../src/utils/formatters"

describe("formatters", () => {
  describe("mail formatters", () => {
    const message: GraphMessage = {
      id: "msg-1",
      subject: "Test Subject",
      from: { emailAddress: { name: "John", address: "john@example.com" } },
      toRecipients: [{ emailAddress: { name: "Jane", address: "jane@example.com" } }],
      receivedDateTime: "2024-01-15T10:00:00Z",
      isRead: false,
      hasAttachments: true,
      bodyPreview: "Hello...",
      body: { contentType: "Text", content: "Hello World" },
      importance: "high",
    }

    it("should format message list", () => {
      const result = formatMessageList([message])
      expect(result).toContain("# Messages")
      expect(result).toContain("Test Subject")
      expect(result).toContain("[Unread]")
      expect(result).toContain("[Attachments]")
    })

    it("should format empty message list", () => {
      expect(formatMessageList([])).toBe("No messages found.")
    })

    it("should format message detail", () => {
      const result = formatMessageDetail(message)
      expect(result).toContain("# Test Subject")
      expect(result).toContain("john@example.com")
      expect(result).toContain("jane@example.com")
      expect(result).toContain("Hello World")
    })
  })

  describe("calendar formatters", () => {
    const event: GraphEvent = {
      id: "evt-1",
      subject: "Team Meeting",
      start: { dateTime: "2024-01-15T14:00:00", timeZone: "UTC" },
      end: { dateTime: "2024-01-15T15:00:00", timeZone: "UTC" },
      location: { displayName: "Room A" },
      organizer: { emailAddress: { name: "Alice", address: "alice@example.com" } },
      attendees: [{ emailAddress: { name: "Bob", address: "bob@example.com" }, status: { response: "accepted" } }],
      isAllDay: false,
      isCancelled: false,
    }

    it("should format event list", () => {
      const result = formatEventList([event])
      expect(result).toContain("# Events")
      expect(result).toContain("Team Meeting")
      expect(result).toContain("@ Room A")
    })

    it("should format event detail", () => {
      const result = formatEventDetail(event)
      expect(result).toContain("# Team Meeting")
      expect(result).toContain("alice@example.com")
      expect(result).toContain("bob@example.com")
      expect(result).toContain("(accepted)")
    })
  })

  describe("user formatters", () => {
    const user: GraphUser = {
      id: "user-1",
      displayName: "Test User",
      mail: "test@example.com",
      userPrincipalName: "test@example.com",
      jobTitle: "Engineer",
      department: "Engineering",
    }

    it("should format user detail", () => {
      const result = formatUserDetail(user)
      expect(result).toContain("# Test User")
      expect(result).toContain("test@example.com")
      expect(result).toContain("Engineer")
      expect(result).toContain("Engineering")
    })
  })

  describe("todo formatters", () => {
    const task: GraphTodoTask = {
      id: "task-1",
      title: "Buy groceries",
      status: "notStarted",
      importance: "high",
      dueDateTime: { dateTime: "2024-01-20T00:00:00", timeZone: "UTC" },
    }

    it("should format todo task list", () => {
      const result = formatTodoTaskList([task])
      expect(result).toContain("# To Do Tasks")
      expect(result).toContain("Buy groceries")
      expect(result).toContain("[notStarted]")
    })

    it("should format todo task detail", () => {
      const result = formatTodoTaskDetail(task)
      expect(result).toContain("# Buy groceries")
      expect(result).toContain("notStarted")
      expect(result).toContain("high")
    })
  })
})
