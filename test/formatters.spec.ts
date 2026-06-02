import { describe, expect, it } from "vitest"

import type {
  GraphEvent,
  GraphMessage,
  GraphNotebook,
  GraphPage,
  GraphSection,
  GraphTodoList,
  GraphTodoTask,
  GraphUser,
} from "../src/types"
import {
  formatEventDetail,
  formatEventList,
  formatMessageDetail,
  formatMessageList,
  formatNotebookList,
  formatPageList,
  formatSectionList,
  formatTodoListList,
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
      expect(result).toContain("ID: msg-1")
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
      expect(result).toContain("- ID: msg-1")
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
      expect(result).toContain("ID: evt-1")
    })

    it("should format event detail", () => {
      const result = formatEventDetail(event)
      expect(result).toContain("# Team Meeting")
      expect(result).toContain("alice@example.com")
      expect(result).toContain("bob@example.com")
      expect(result).toContain("(accepted)")
      expect(result).toContain("- ID: evt-1")
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
      expect(result).toContain("ID: task-1")
    })

    it("should include the list ID so list_todo_tasks can be called", () => {
      const list: GraphTodoList = { id: "list-1", displayName: "Tasks", wellknownListName: "defaultList" }
      const result = formatTodoListList([list])
      expect(result).toContain("# To Do Lists")
      expect(result).toContain("Tasks")
      expect(result).toContain("[defaultList]")
      expect(result).toContain("ID: list-1")
    })

    it("should format todo task detail", () => {
      const result = formatTodoTaskDetail(task)
      expect(result).toContain("# Buy groceries")
      expect(result).toContain("notStarted")
      expect(result).toContain("high")
      expect(result).toContain("- ID: task-1")
    })
  })

  describe("onenote formatters", () => {
    it("should include the notebook ID so the typed tools can chain", () => {
      const notebook: GraphNotebook = { id: "nb-1", displayName: "Graph API Test", isDefault: true }
      const result = formatNotebookList([notebook])
      expect(result).toContain("# Notebooks")
      expect(result).toContain("Graph API Test")
      expect(result).toContain("[Default]")
      expect(result).toContain("ID: nb-1")
    })

    it("should include the section ID so list_onenote_pages can be called", () => {
      const section: GraphSection = { id: "sec-1", displayName: "Quick Notes" }
      const result = formatSectionList([section])
      expect(result).toContain("# Sections")
      expect(result).toContain("Quick Notes")
      expect(result).toContain("ID: sec-1")
    })

    it("should include the page ID so get_onenote_page_content can be called", () => {
      const page: GraphPage = { id: "pg-1", title: "Meeting Notes", lastModifiedDateTime: "2026-06-02T10:00:00Z" }
      const result = formatPageList([page])
      expect(result).toContain("# Pages")
      expect(result).toContain("Meeting Notes")
      expect(result).toContain("ID: pg-1")
    })
  })
})
