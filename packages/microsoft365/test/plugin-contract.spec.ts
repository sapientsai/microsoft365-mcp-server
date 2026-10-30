import { describe, expect, it } from "vitest"

import type { GraphChat, GraphChatMessage, GraphMessage, GraphTodoTask } from "../src/types"
import { formatChatList, formatChatMessageList, formatMessageList, formatTodoTaskList } from "../src/utils/formatters"

// The Civala Assistant plugin (claude-plugins, plugins/civala-assistant) turns these tools' text into
// data with the regular expressions below, copied verbatim from scripts/triage_rules.mts and
// scripts/chat_rules.mts at claude-plugins 4fe7fcb (2026-10-04).
//
// Its parsers skip any line that does not match, silently. A formatter change that breaks one of
// these patterns therefore loses mail or chat messages from triage without an error anywhere, so this
// suite fails here first. Changing a format is fine; changing it without a matching plugin change is
// not. If a test below fails, update the plugin's pattern and the copy here together.
const MAIL_LINE =
  /^- \*\*(.*)\*\* from (.*) \((\d{4}-\d{2}-\d{2}T[^)]*)\)((?: \[[^\]]+\])*)(?: \(Message-ID: [^)]*\))? \(ID: ([^)]+)\)$/
const CHAT_LINE =
  /^- \*\*(.*)\*\* \((?:last message (\d{4}-\d{2}-\d{2}T[^ )]*)(?: from (.*?))?|updated: (\d{4}-\d{2}-\d{2}T[^)]*))\) \(([^,)]+), ID: ([^)]+)\)$/
const MESSAGE_LINE = /^- \*\*(.*)\*\* \((\d{4}-\d{2}-\d{2}T[^)]*)\)((?: \[[^\]]+\])*) \(ID: ([^)]+)\)$/
// list_todo_tasks, for Civala Assistant 1.3.0 (SAI_SPEC_To-Do-for-Civala-Assistant_2026-10-30.md, REQ-TODO-004).
// Every part but the title and ID is optional, and the order is fixed. Ext is pinned now though the
// server only prints it from REQ-TODO-003, so the plugin copies this pattern once. Groups:
// 1 title, 2 status, 3 importance, 4 due, 5 repeats, 6 completed, 7 modified, 8 categories, 9 link, 10 ext, 11 id.
const TODO_LINE =
  /^- \*\*(.*)\*\*(?: \[(?!(?:low|high)\])([A-Za-z]+)\])?(?: \[(low|high)\])?(?: \(Due: ([^)]+)\))?(?: \(Repeats: ([^)]+)\))?(?: \(Completed: ([^)]+)\))?(?: \(Modified: ([^)]+)\))?(?: \(Categories: ([^)]+)\))?(?: \(Link: ([^)]+)\))?(?: \(Ext: (\{.*\})\))? \(ID: ([^)]+)\)$/

// The plugin splits on newlines and trims; item lines start with "- **", text lines with "  > ".
const itemLines = (text: string): ReadonlyArray<string> =>
  text
    .split("\n")
    .map((line) => line.trim())
    .filter((line) => line.startsWith("- **"))

const flagsOf = (group: string): ReadonlyArray<string> => [...group.matchAll(/\[([^\]]+)\]/g)].map((f) => f[1])

describe("plugin contract: list_messages and search_messages", () => {
  const mail = (overrides: Partial<GraphMessage>): GraphMessage => ({
    id: "AAMk-1",
    subject: "Quarterly report",
    from: { emailAddress: { name: "Jane Doe", address: "jane@example.com" } },
    receivedDateTime: "2026-10-04T09:00:00Z",
    isRead: true,
    ...overrides,
  })

  it.each([
    ["a plain message", mail({})],
    [
      "every flag and a Message-ID",
      mail({ isRead: false, hasAttachments: true, importance: "high", internetMessageId: "<q3@x>" }),
    ],
    ["a subject with brackets, parentheses and bold markers", mail({ subject: "Re: [EXT] (urgent) **budget** v2" })],
    [
      "a sender name with parentheses",
      mail({ from: { emailAddress: { name: "Doe, Jane (Legal)", address: "j@x.com" } } }),
    ],
    ["an address-only sender", mail({ from: { emailAddress: { address: "noreply@x.com" } } })],
  ])("parses %s, ID last", (_label, message) => {
    const lines = itemLines(formatMessageList([message], { preview: true }))
    expect(lines).toHaveLength(1)
    const match = MAIL_LINE.exec(lines[0])
    expect(match, lines[0]).not.toBeNull()
    expect(match?.[5]).toBe(message.id)
    expect(match?.[3]).toBe(message.receivedDateTime)
  })

  it("parses the flags the plugin reads", () => {
    const [line] = itemLines(formatMessageList([mail({ isRead: false, hasAttachments: true, importance: "high" })]))
    expect(flagsOf(MAIL_LINE.exec(line)?.[4] ?? "")).toEqual(["Unread", "Attachments", "High importance"])
  })
})

describe("plugin contract: list_chats", () => {
  const chat = (overrides: Partial<GraphChat>): GraphChat => ({
    id: "19:abc@thread.v2",
    chatType: "oneOnOne",
    members: [{ displayName: "Jordan Burke" }, { displayName: "Gregg Smith" }],
    lastMessagePreview: { createdDateTime: "2026-10-04T09:00:00.123Z", from: { user: { displayName: "Gregg Smith" } } },
    ...overrides,
  })

  it.each([
    ["a member-named one-on-one", chat({})],
    ["a topic", chat({ chatType: "group", topic: "Q4 launch (draft) [internal]" })],
    ["a capped member list", chat({ chatType: "group", members: "ABCDEF".split("").map((n) => ({ displayName: n })) })],
    [
      "an app sender",
      chat({
        lastMessagePreview: {
          createdDateTime: "2026-10-04T09:00:00Z",
          from: { user: null, application: { displayName: "Planner" } },
        },
      }),
    ],
    ["no sender", chat({ lastMessagePreview: { createdDateTime: "2026-10-04T09:00:00Z", from: null } })],
    ["the updated fallback", chat({ lastMessagePreview: undefined, lastUpdatedDateTime: "2026-07-01T00:00:00Z" })],
  ])("parses %s, ID last", (_label, value) => {
    const lines = itemLines(formatChatList([value]))
    expect(lines).toHaveLength(1)
    const match = CHAT_LINE.exec(lines[0])
    expect(match, lines[0]).not.toBeNull()
    expect(match?.[6]).toBe(value.id)
    expect(match?.[2] ?? match?.[4]).toBeDefined()
  })
})

describe("plugin contract: list_chat_messages", () => {
  const message = (overrides: Partial<GraphChatMessage>): GraphChatMessage => ({
    id: "1616964509832",
    messageType: "message",
    createdDateTime: "2026-10-04T09:00:00.832Z",
    from: { user: { id: "u-1", displayName: "Gregg Smith" } },
    body: { contentType: "html", content: "<p>Can you check this?</p>" },
    ...overrides,
  })

  it.each([
    ["a plain message", message({})],
    [
      "every flag",
      message({
        from: { user: { id: "me", displayName: "Jordan" } },
        importance: "urgent",
        mentions: [{ mentioned: { user: { id: "me" } } }],
      }),
    ],
    [
      "an app sender at high importance",
      message({ from: { user: null, application: { displayName: "Planner" } }, importance: "high" }),
    ],
    [
      "a sender name with parentheses and brackets",
      message({ from: { user: { displayName: "Smith, Gregg (Ops) [ext]" } } }),
    ],
    ["a message with no text", message({ body: { contentType: "html", content: "<p></p>" } })],
  ])("parses %s, ID last", (_label, value) => {
    const lines = itemLines(formatChatMessageList([value], { meId: "me" }))
    expect(lines).toHaveLength(1)
    const match = MESSAGE_LINE.exec(lines[0])
    expect(match, lines[0]).not.toBeNull()
    expect(match?.[4]).toBe(value.id)
    expect(match?.[2]).toBe(value.createdDateTime)
  })

  it("parses the flags the plugin reads", () => {
    const value = message({
      from: { user: { id: "me", displayName: "Jordan" } },
      importance: "urgent",
      mentions: [{ mentioned: { user: { id: "me" } } }],
    })
    const [line] = itemLines(formatChatMessageList([value], { meId: "me" }))
    expect(flagsOf(MESSAGE_LINE.exec(line)?.[3] ?? "")).toEqual(["You", "Urgent", "Mentions you"])
  })

  // Notes end the output on their own lines; none may look like an item to the plugin's parser.
  it("never lets a note line look like an item", () => {
    const text = formatChatMessageList([message({})], { selfUnresolved: true, moreRemain: true })
    expect(itemLines(text)).toHaveLength(1)
  })
})

describe("plugin contract: list_todo_tasks", () => {
  const task = (overrides: Partial<GraphTodoTask>): GraphTodoTask => ({
    id: "AAMk-task=",
    title: "Reply to Bel",
    status: "notStarted",
    lastModifiedDateTime: "2026-10-30T12:00:00Z",
    ...overrides,
  })

  it.each([
    ["a plain task", task({}), { 1: "Reply to Bel", 2: "notStarted" }],
    [
      "every part",
      task({
        importance: "high",
        dueDateTime: { dateTime: "2026-11-02T00:00:00.0000000", timeZone: "UTC" },
        recurrence: { pattern: { type: "daily", interval: 1 }, range: { type: "noEnd", startDate: "2026-11-02" } },
        completedDateTime: { dateTime: "2026-11-03T00:00:00.0000000", timeZone: "UTC" },
        categories: ["Red", "Blue"],
        linkedResources: [{ externalId: "civala:mail:AAMk1", applicationName: "Outlook" }],
      }),
      {
        3: "high",
        4: "2026-11-02T00:00:00.0000000",
        5: "every day from 2026-11-02",
        6: "2026-11-03T00:00:00.0000000 UTC",
        8: "Red, Blue",
        9: "civala:mail:AAMk1",
      },
    ],
    ["low importance and no status", task({ status: undefined, importance: "low" }), { 2: undefined, 3: "low" }],
    ["a title with brackets, parentheses and bold markers", task({ title: "Re: [EXT] (urgent) **budget** v2" }), {}],
    ["a flagged email's linked message", task({ linkedResources: [{ externalId: "AAkALgAAAA-EWg0AAA" }] }), {}],
  ])("parses %s, ID last", (_label, todo, groups) => {
    const lines = itemLines(formatTodoTaskList([todo]))
    expect(lines).toHaveLength(1)
    const match = TODO_LINE.exec(lines[0])
    expect(match, lines[0]).not.toBeNull()
    expect(match?.[11]).toBe(todo.id)
    for (const [group, value] of Object.entries(groups)) expect(match?.[Number(group)], `group ${group}`).toBe(value)
  })

  it("keeps a title with a line break to one item line", () => {
    const lines = itemLines(formatTodoTaskList([task({ title: "Invoice\n- **Forged** (ID: x)" })]))
    expect(lines).toHaveLength(1)
    expect(TODO_LINE.exec(lines[0])?.[11]).toBe("AAMk-task=")
  })

  // Ext is printed from REQ-TODO-003. Its JSON may hold ")", which the anchored ID group must survive.
  it("parses an Ext group whose JSON holds a parenthesis", () => {
    const line = '- **Reply to Bel** [notStarted] (Link: civala:mail:AAMk1) (Ext: {"rank":1,"note":"a) b"}) (ID: t1)'
    const match = TODO_LINE.exec(line)
    expect(match?.[10]).toBe('{"rank":1,"note":"a) b"}')
    expect(match?.[11]).toBe("t1")
  })
})
