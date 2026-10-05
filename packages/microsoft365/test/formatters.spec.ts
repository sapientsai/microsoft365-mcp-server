import { describe, expect, it } from "vitest"

import type {
  GraphChat,
  GraphChatMessage,
  GraphDriveItem,
  GraphEvent,
  GraphMeetingTimeSuggestionsResult,
  GraphMessage,
  GraphNotebook,
  GraphPage,
  GraphSection,
  GraphTodoList,
  GraphTodoTask,
  GraphUser,
} from "../src/types"
import {
  CHAT_SELF_UNRESOLVED_NOTE,
  chatMessageText,
  formatChatList,
  formatChatMessageList,
  formatDriveItemDetail,
  formatDriveItemList,
  formatEventDetail,
  formatEventList,
  formatMeetingTimeSuggestions,
  formatMessageDetail,
  formatMessageList,
  formatNotebookList,
  formatPageList,
  formatSectionList,
  formatTodoListList,
  formatTodoTaskDetail,
  formatTodoTaskList,
  formatUserDetail,
  MESSAGE_SUMMARY_FIELDS,
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

    // Callers match senders by address; a display name alone ("John") cannot be matched reliably.
    it("shows the sender's name and address together", () => {
      expect(formatMessageList([message])).toContain("from John <john@example.com>")
    })

    it("falls back to whichever of name and address exists", () => {
      const sender = (emailAddress: { name?: string; address?: string }) =>
        formatMessageList([{ id: "m", from: { emailAddress } }])
      expect(sender({ address: "a@example.com" })).toContain("from a@example.com (")
      expect(sender({ name: "Only Name" })).toContain("from Only Name (")
      expect(sender({ name: "a@example.com", address: "a@example.com" })).toContain("from a@example.com (")
      expect(formatMessageList([{ id: "m" }])).toContain("from Unknown (")
    })

    it("shows the internetMessageId when Graph returns it", () => {
      expect(formatMessageList([{ ...message, internetMessageId: "<abc@mail.example.com>" }])).toContain(
        "(Message-ID: <abc@mail.example.com>) (ID: msg-1)",
      )
      expect(formatMessageList([message])).not.toContain("Message-ID")
    })

    it("flags high importance after the other flags, and nothing else", () => {
      const line = (importance?: string) =>
        formatMessageList([{ id: "m", isRead: false, hasAttachments: true, importance, internetMessageId: "<x@y>" }])
      expect(line("high")).toContain("[Unread] [Attachments] [High importance] (Message-ID: <x@y>) (ID: m)")
      expect(line("High")).toContain("[High importance]")
      expect(line("normal")).not.toContain("importance")
      expect(line("low")).not.toContain("importance")
      expect(line(undefined)).not.toContain("importance")
    })

    // Callers parse the Graph ID as the last element of each line; anything appended after it breaks them.
    it("ends every summary line with the Graph ID", () => {
      const line = formatMessageList([{ ...message, internetMessageId: "<abc@mail.example.com>" }])
        .split("\n")
        .pop()
      expect(line).toMatch(/\(ID: msg-1\)$/)
    })

    it("strips angle brackets from a display name, so the address is the only <...>", () => {
      const result = formatMessageList([
        { id: "m", from: { emailAddress: { name: "Jane <Sales>", address: "j@x.com" } } },
      ])
      expect(result).toContain("from Jane Sales <j@x.com> (")
    })

    it("adds the body preview only when asked, on one collapsed line", () => {
      const withPreview = { ...message, bodyPreview: "Hi team,\r\n\r\n  the report   is attached." }
      expect(formatMessageList([withPreview])).not.toContain("the report")
      expect(formatMessageList([withPreview], { preview: true })).toContain("\n  > Hi team, the report is attached.")
    })

    it("adds no preview line for an empty preview", () => {
      expect(formatMessageList([{ ...message, bodyPreview: "  \n " }], { preview: true })).not.toContain("  >")
    })

    // list_messages $selects MESSAGE_SUMMARY_FIELDS. A field the formatter reads but the list lacks
    // would print blank with no error, so record every field it actually touches.
    it("reads no message field outside MESSAGE_SUMMARY_FIELDS", () => {
      const read = new Set<string>()
      const tracked = new Proxy(
        { ...message, internetMessageId: "<abc@mail.example.com>" },
        {
          get: (target, key, receiver) => {
            if (typeof key === "string") read.add(key)
            return Reflect.get(target, key, receiver)
          },
        },
      )
      formatMessageList([tracked], { preview: true })

      expect(read.size).toBeGreaterThan(0)
      expect([...read].filter((key) => !(MESSAGE_SUMMARY_FIELDS as ReadonlyArray<string>).includes(key))).toEqual([])
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

  describe("meeting time suggestions formatter", () => {
    it("should render slots with confidence and attendee availability", () => {
      const result: GraphMeetingTimeSuggestionsResult = {
        emptySuggestionsReason: "",
        meetingTimeSuggestions: [
          {
            confidence: 100,
            organizerAvailability: "free",
            attendeeAvailability: [
              { availability: "free", attendee: { emailAddress: { address: "bob@example.com" } } },
            ],
            meetingTimeSlot: {
              start: { dateTime: "2026-06-04T15:00:00.0000000", timeZone: "UTC" },
              end: { dateTime: "2026-06-04T15:30:00.0000000", timeZone: "UTC" },
            },
          },
        ],
      }
      const output = formatMeetingTimeSuggestions(result)
      expect(output).toContain("# Meeting Time Suggestions")
      expect(output).toContain("2026-06-04T15:00:00.0000000 → 2026-06-04T15:30:00.0000000")
      expect(output).toContain("100% confidence")
      expect(output).toContain("bob@example.com: free")
    })

    it("should render the no-availability message with the empty reason", () => {
      const output = formatMeetingTimeSuggestions({
        emptySuggestionsReason: "AttendeesUnavailable",
        meetingTimeSuggestions: [],
      })
      expect(output).toContain("No common availability found")
      expect(output).toContain("AttendeesUnavailable")
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
  describe("chat formatters", () => {
    const member = (displayName: string) => ({ displayName })

    it("names an untitled chat by its members and shows the last message's time and sender", () => {
      const chat: GraphChat = {
        id: "c1",
        chatType: "oneOnOne",
        lastUpdatedDateTime: "2026-07-01T00:00:00Z",
        members: [member("Jordan Burke"), member("Gregg Smith")],
        lastMessagePreview: { createdDateTime: "2026-10-04T09:00:00Z", from: { user: { displayName: "Gregg Smith" } } },
      }
      expect(formatChatList([chat])).toContain(
        "- **Jordan Burke, Gregg Smith** (last message 2026-10-04T09:00:00Z from Gregg Smith) (oneOnOne, ID: c1)",
      )
    })

    it("prefers the topic, and caps a large member list", () => {
      const members = ["A", "B", "C", "D", "E", "F"].map(member)
      expect(formatChatList([{ id: "g1", chatType: "group", topic: "Launch", members }])).toContain("- **Launch**")
      expect(formatChatList([{ id: "g2", chatType: "group", members }])).toContain("- **A, B, C, D +2 more**")
    })

    it("names a bot sender, and omits the sender when Graph gives none", () => {
      const preview = (from: GraphChat["lastMessagePreview"]) => formatChatList([{ id: "c", lastMessagePreview: from }])
      expect(preview({ createdDateTime: "T1", from: { user: null, application: { displayName: "Bot" } } })).toContain(
        "(last message T1 from Bot)",
      )
      expect(preview({ createdDateTime: "T2", from: null })).toContain("(last message T2) (")
    })

    it("falls back to the update time and the chat type when Graph returns no preview or members", () => {
      expect(formatChatList([{ id: "c", chatType: "group", lastUpdatedDateTime: "T3" }])).toContain(
        "- **group** (updated: T3) (group, ID: c)",
      )
    })
  })

  describe("chat message formatters", () => {
    const ME = "me-1"
    const msg = (overrides: Partial<GraphChatMessage>): GraphChatMessage => ({
      id: "m1",
      messageType: "message",
      createdDateTime: "2026-10-04T09:00:00Z",
      from: { user: { id: "u-gregg", displayName: "Gregg Smith" } },
      body: { contentType: "text", content: "Hello" },
      ...overrides,
    })

    it("reduces HTML to one line of text, keeping mention names and decoding entities", () => {
      const html =
        '<div><at id="0">Jordan</at>&nbsp;can you check Q3?<br>It&#39;s &lt;urgent&gt; &amp; due &#x2014; today</div>'
      expect(chatMessageText({ contentType: "html", content: html })).toBe(
        "Jordan can you check Q3? It's <urgent> & due — today",
      )
    })

    it("marks a message that is only an attachment or only an image", () => {
      expect(chatMessageText({ contentType: "html", content: '<attachment id="a1"></attachment>' })).toBe(
        "[attachment]",
      )
      expect(chatMessageText({ contentType: "html", content: '<p><img src="x" width="67"></p>' })).toBe("[image]")
    })

    it("leaves text bodies alone apart from whitespace", () => {
      expect(chatMessageText({ contentType: "text", content: "  a &amp; b\n\n c " })).toBe("a &amp; b c")
    })

    it("puts the flags in a fixed order and the ID last, with the text on a '  > ' line", () => {
      const all = msg({
        from: { user: { id: ME, displayName: "Jordan Burke" } },
        importance: "urgent",
        mentions: [{ mentioned: { user: { id: ME } } }],
      })
      expect(formatChatMessageList([all], { meId: ME })).toContain(
        "- **Jordan Burke** (2026-10-04T09:00:00Z) [You] [Urgent] [Mentions you] (ID: m1)\n  > Hello",
      )
    })

    it("flags an app sender and high importance", () => {
      const bot = msg({ from: { user: null, application: { displayName: "Planner" } }, importance: "high" })
      expect(formatChatMessageList([bot], { meId: ME })).toContain(
        "- **Planner** (2026-10-04T09:00:00Z) [App] [High importance] (ID: m1)",
      )
    })

    // A mention of the whole chat or a tag has no user, and is not a mention of you.
    it("does not count a chat-wide mention as mentioning you", () => {
      const everyone = msg({ mentions: [{ mentioned: { user: null } }, { mentioned: null }] })
      expect(formatChatMessageList([everyone], { meId: ME })).not.toContain("[Mentions you]")
    })

    it("leaves out system events and deleted messages", () => {
      const result = formatChatMessageList([
        msg({
          id: "sys",
          messageType: "systemEventMessage",
          body: { contentType: "html", content: "<systemEventMessage/>" },
        }),
        msg({ id: "future", messageType: "unknownFutureValue" }),
        msg({ id: "gone", deletedDateTime: "2026-10-04T10:00:00Z" }),
        msg({ id: "kept" }),
      ])
      expect(result).toContain("(ID: kept)")
      expect(result).not.toMatch(/\(ID: (sys|future|gone)\)/)
    })

    it("cuts long text at max_chars and omits the text line when there is none", () => {
      const long = msg({ body: { contentType: "text", content: "x".repeat(400) } })
      expect(formatChatMessageList([long])).toContain(`  > ${"x".repeat(300)}…`)
      expect(formatChatMessageList([long], { maxChars: 10 })).toContain(`  > ${"x".repeat(10)}…`)
      expect(formatChatMessageList([msg({ body: { contentType: "html", content: "<p></p>" } })])).not.toContain("  >")
    })

    it("adds one fixed note at the end when the signed-in user is unknown, and no [You]", () => {
      const result = formatChatMessageList([msg({ from: { user: { id: ME, displayName: "Jordan" } } })], {
        selfUnresolved: true,
      })
      expect(result).not.toContain("[You]")
      expect(result.endsWith(`\n\n${CHAT_SELF_UNRESOLVED_NOTE}`)).toBe(true)
    })
  })

  describe("drive item formatters", () => {
    const item: GraphDriveItem = {
      id: "item-1",
      name: "Plan.docx",
      size: 2048,
      lastModifiedDateTime: "2026-09-30T12:00:00Z",
      file: { mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
      parentReference: { driveId: "drive-1", id: "parent-1", path: "/drive/root:/Work/Oncala" },
    }

    it("shows the parent path, parent ID and drive ID in the detail view", () => {
      const result = formatDriveItemDetail(item)
      expect(result).toContain("- Parent Path: /drive/root:/Work/Oncala")
      expect(result).toContain("- Parent ID: parent-1")
      expect(result).toContain("- Drive ID: drive-1")
    })

    it("omits the parent lines when Graph returns no parentReference", () => {
      const result = formatDriveItemDetail({ id: "item-2", name: "Loose.txt" })
      expect(result).not.toContain("Parent")
      expect(result).not.toContain("Drive ID")
    })

    it("puts the parent path and modified date on each summary line", () => {
      const result = formatDriveItemList([item])
      expect(result).toContain("- in /drive/root:/Work/Oncala")
      expect(result).toContain("- modified 2026-09-30T12:00:00Z")
    })

    it("falls back to the parent ID when the path is missing, as /search returns", () => {
      const result = formatDriveItemList([{ ...item, parentReference: { driveId: "drive-1", id: "parent-1" } }])
      expect(result).toContain("- parent ID: parent-1")
      expect(result).not.toContain("- in ")
    })

    it("keeps the bare summary line when Graph returns neither parent nor date", () => {
      expect(formatDriveItemList([{ id: "item-3", name: "Bare" }])).toBe("# Files\n\n- **Bare** (ID: item-3) - File")
    })

    it("decodes a percent-encoded parent path in both views", () => {
      const encoded = {
        ...item,
        parentReference: { id: "parent-1", path: "/drive/root:/Client%20Files/100%25%20Done" },
      }
      expect(formatDriveItemDetail(encoded)).toContain("- Parent Path: /drive/root:/Client Files/100% Done")
      expect(formatDriveItemList([encoded])).toContain("- in /drive/root:/Client Files/100% Done")
    })

    it("prints a malformed path as Graph sent it instead of throwing", () => {
      const malformed = { ...item, parentReference: { path: "/drive/root:/Bad%ZZ" } }
      expect(formatDriveItemList([malformed])).toContain("- in /drive/root:/Bad%ZZ")
    })

    // Graph's search index reports zero for a folder's count and size; printing them claims an empty folder.
    it("omits a folder's count and size in search results only", () => {
      const folder: GraphDriveItem = { id: "f1", name: "Reports", folder: { childCount: 0 }, size: 0 }
      expect(formatDriveItemList([folder], { fromSearch: true })).toBe("# Files\n\n- **Reports** (ID: f1) - Folder")
      expect(formatDriveItemList([folder])).toContain("- Folder (0 items) (0 B)")
      const file: GraphDriveItem = { id: "x", name: "a.txt", size: 10, file: { mimeType: "text/plain" } }
      expect(formatDriveItemList([file], { fromSearch: true })).toContain("- text/plain (10 B)")
    })

    it("keeps a real count and size if search does return them", () => {
      const folder: GraphDriveItem = { id: "f2", name: "Real", folder: { childCount: 3 }, size: 2048 }
      expect(formatDriveItemList([folder], { fromSearch: true })).toContain("- Folder (3 items) (2.0 KB)")
    })

    it("shows a folder's parent on its summary line", () => {
      const folder: GraphDriveItem = {
        id: "folder-1",
        name: "Reports",
        folder: { childCount: 3 },
        parentReference: { path: "/drive/root:/Work" },
      }
      expect(formatDriveItemList([folder])).toContain(
        "- **Reports** (ID: folder-1) - Folder (3 items) - in /drive/root:/Work",
      )
    })

    it("shows only the parent ID in the detail view when Graph omits the path", () => {
      const result = formatDriveItemDetail({ ...item, parentReference: { id: "parent-1" } })
      expect(result).toContain("- Parent ID: parent-1")
      expect(result).not.toContain("Parent Path")
      expect(result).not.toContain("Drive ID")
    })
  })
})
