import { Some } from "functype"
import { Left, Right } from "functype/either"
import { beforeEach, describe, expect, it, vi } from "vitest"

vi.mock("../src/client/graph-client", () => ({
  getGraphClient: vi.fn(),
}))

import { getGraphClient } from "../src/client/graph-client"
import { listChatMessages, listChats, nextLinkPath } from "../src/tools/chat-tools"
import { CHAT_MORE_MESSAGES_NOTE, CHAT_SELF_UNRESOLVED_NOTE } from "../src/utils/formatters"
import { searchFiles } from "../src/tools/files-tools"
import { searchSiteFiles } from "../src/tools/sharepoint-tools"

const mockClient = {
  listChats: vi.fn(),
  listChatMessages: vi.fn(),
  getMe: vi.fn(),
  request: vi.fn(),
  requestPaginated: vi.fn(),
  searchFiles: vi.fn(),
  searchAllFiles: vi.fn(),
  searchSiteFiles: vi.fn(),
  searchAllSiteFiles: vi.fn(),
}

const CHAT_PARAMS = { $expand: ["members", "lastMessagePreview"], $orderby: "lastMessagePreview/createdDateTime desc" }

// Search reports zero for a folder's count and size, so the line must not claim "0 items".
const searchFolder = { id: "f1", name: "Reports", folder: { childCount: 0 }, size: 0 }

beforeEach(() => {
  vi.clearAllMocks()
  vi.mocked(getGraphClient).mockReturnValue(Some(mockClient as never))
})

describe("listChats", () => {
  it("asks for members and the last message, sorted newest first", async () => {
    mockClient.listChats.mockResolvedValue(Right({ value: [] }))

    await listChats({ top: 5 })

    expect(mockClient.listChats).toHaveBeenCalledWith({ ...CHAT_PARAMS, $top: 5 })
  })

  it("defaults to 25 chats", async () => {
    mockClient.listChats.mockResolvedValue(Right({ value: [] }))

    await listChats({})

    expect(mockClient.listChats).toHaveBeenCalledWith({ ...CHAT_PARAMS, $top: 25 })
  })

  it("keeps the same expand and sort when paging through every chat", async () => {
    mockClient.requestPaginated.mockResolvedValue(Right([]))

    await listChats({ fetch_all_pages: true })

    expect(mockClient.requestPaginated).toHaveBeenCalledWith("/me/chats", { odataParams: CHAT_PARAMS })
  })
})

describe("listChats since", () => {
  const chat = (id: string, last?: string) => ({
    id,
    chatType: "oneOnOne",
    members: [{ displayName: `Person ${id}` }],
    ...(last ? { lastMessagePreview: { createdDateTime: last } } : {}),
  })
  const page = (value: ReadonlyArray<unknown>, next?: string) =>
    Right({ value, ...(next ? { "@odata.nextLink": next } : {}) })
  const NEXT = "https://graph.microsoft.com/v1.0/me/chats?$expand=members,lastMessagePreview&$skiptoken=abc"

  it("stops at the first chat whose last message is not after since, without reading further pages", async () => {
    mockClient.listChats.mockResolvedValue(
      page(
        [chat("a", "2026-10-04T10:00:00Z"), chat("b", "2026-10-03T00:00:00Z"), chat("c", "2026-10-04T11:00:00Z")],
        NEXT,
      ),
    )

    const result = await listChats({ since: "2026-10-04T00:00:00Z" })

    expect(mockClient.listChats).toHaveBeenCalledWith({ ...CHAT_PARAMS, $top: 50 })
    expect(mockClient.request).not.toHaveBeenCalled()
    expect(result.value).toContain("ID: a)")
    expect(result.value).not.toMatch(/ID: (b|c)\)/)
  })

  it("follows nextLink while every chat is newer, converting it to a path", async () => {
    mockClient.listChats.mockResolvedValue(page([chat("a", "2026-10-04T10:00:00Z")], NEXT))
    mockClient.request.mockResolvedValue(page([chat("b", "2026-10-04T09:00:00Z"), chat("c", "2026-09-01T00:00:00Z")]))

    const result = await listChats({ since: "2026-10-04T00:00:00Z" })

    expect(mockClient.request).toHaveBeenCalledWith(
      "GET",
      "/me/chats?$expand=members,lastMessagePreview&$skiptoken=abc",
    )
    expect(result.value).toContain("ID: a)")
    expect(result.value).toContain("ID: b)")
    expect(result.value).not.toContain("ID: c)")
  })

  // Graph does not document where a chat with no messages sorts; stopping on one placed first would
  // return nothing.
  it("skips a chat with no last message instead of stopping on it", async () => {
    mockClient.listChats.mockResolvedValue(page([chat("empty"), chat("a", "2026-10-04T10:00:00Z")]))

    const result = await listChats({ since: "2026-10-04T00:00:00Z" })

    expect(result.value).toContain("ID: a)")
    expect(result.value).not.toContain("ID: empty)")
  })

  it("caps the result at top", async () => {
    mockClient.listChats.mockResolvedValue(page([chat("a", "2026-10-04T10:00:00Z"), chat("b", "2026-10-04T09:00:00Z")]))

    const result = await listChats({ since: "2026-10-04T00:00:00Z", top: 1 })

    expect(result.value).toContain("ID: a)")
    expect(result.value).not.toContain("ID: b)")
  })

  it("says so when it stops at the page cap", async () => {
    mockClient.listChats.mockResolvedValue(page([chat("p0", "2026-10-04T10:00:00Z")], NEXT))
    mockClient.request.mockResolvedValue(page([chat("pn", "2026-10-04T09:00:00Z")], NEXT))

    const result = await listChats({ since: "2026-10-04T00:00:00Z" })

    expect(mockClient.request).toHaveBeenCalledTimes(19)
    expect(result.value).toContain("stopped after 20 pages")
  })

  it("returns no chats, without reading further pages, when the newest chat is already older", async () => {
    mockClient.listChats.mockResolvedValue(page([chat("old", "2026-09-01T00:00:00Z")], NEXT))

    const result = await listChats({ since: "2026-10-04T00:00:00Z" })

    expect(mockClient.request).not.toHaveBeenCalled()
    expect(result.value).toBe("No chats found.")
  })

  it("rejects a since that is not a date-time before calling Graph", async () => {
    const result = await listChats({ since: "2026-10-01 or true" })

    expect(result.isLeft()).toBe(true)
    expect(mockClient.listChats).not.toHaveBeenCalled()
  })

  it("strips the API base from a nextLink on either version", () => {
    expect(nextLinkPath("https://graph.microsoft.com/beta/me/chats?$skiptoken=x")).toBe("/me/chats?$skiptoken=x")
    expect(nextLinkPath("https://graph.microsoft.com/v1.0/chats/19:a@thread.v2/messages?$skiptoken=y")).toBe(
      "/chats/19:a@thread.v2/messages?$skiptoken=y",
    )
  })
})

describe("listChatMessages", () => {
  const message = { id: "m1", messageType: "message", from: { user: { id: "me-1", displayName: "Jordan" } } }

  it("looks up the signed-in user on every call and marks their messages [You]", async () => {
    mockClient.getMe.mockResolvedValue(Right({ id: "me-1" }))
    mockClient.listChatMessages.mockResolvedValue(Right({ value: [message] }))

    await listChatMessages({ chat_id: "c1" })
    const result = await listChatMessages({ chat_id: "c1" })

    expect(mockClient.getMe).toHaveBeenCalledTimes(2)
    expect(mockClient.listChatMessages).toHaveBeenCalledWith("c1", { $top: 25 })
    expect(result.value).toContain("[You] (ID: m1)")
  })

  // Graph ignores the filter unless the request orders by the same property.
  it("sends since as a lastModifiedDateTime filter with the matching order, normalised to ISO", async () => {
    mockClient.getMe.mockResolvedValue(Right({ id: "me-1" }))
    mockClient.listChatMessages.mockResolvedValue(Right({ value: [] }))

    await listChatMessages({ chat_id: "c1", since: "2026-10-04T09:00:00+02:00" })

    expect(mockClient.listChatMessages).toHaveBeenCalledWith("c1", {
      $top: 25,
      $orderby: "lastModifiedDateTime desc",
      $filter: "lastModifiedDateTime gt 2026-10-04T07:00:00.000Z",
    })
  })

  // The plugin's exact call shape: top is its budget, and a note is its overflow signal.
  describe("since with top", () => {
    const readable = (id: string) => ({ id, messageType: "message", createdDateTime: "2026-10-04T10:00:00Z" })
    const NEXT = "https://graph.microsoft.com/v1.0/chats/c1/messages?$top=3&$skiptoken=t"

    beforeEach(() => mockClient.getMe.mockResolvedValue(Right({ id: "me-1" })))

    it("stops reading once top readable messages are in hand, and notes that more match", async () => {
      mockClient.listChatMessages.mockResolvedValue(
        Right({ value: [readable("a"), readable("b"), readable("c")], "@odata.nextLink": NEXT }),
      )

      const result = await listChatMessages({ chat_id: "c1", since: "2026-10-04T00:00:00Z", top: 2 })

      expect(mockClient.listChatMessages).toHaveBeenCalledWith("c1", expect.objectContaining({ $top: 2 }))
      expect(mockClient.request).not.toHaveBeenCalled()
      expect(result.value).toContain("(ID: a)")
      expect(result.value).toContain("(ID: b)")
      expect(result.value).not.toContain("(ID: c)")
      expect(result.value).toContain(CHAT_MORE_MESSAGES_NOTE)
    })

    it("follows nextLink when system events leave the page short, counting only readable messages", async () => {
      mockClient.listChatMessages.mockResolvedValue(
        Right({ value: [readable("a"), { id: "sys", messageType: "systemEventMessage" }], "@odata.nextLink": NEXT }),
      )
      mockClient.request.mockResolvedValue(Right({ value: [readable("b")] }))

      const result = await listChatMessages({ chat_id: "c1", since: "2026-10-04T00:00:00Z", top: 2 })

      expect(mockClient.request).toHaveBeenCalledWith("GET", "/chats/c1/messages?$top=3&$skiptoken=t")
      expect(result.value).toContain("(ID: b)")
      expect(result.value).not.toContain(CHAT_MORE_MESSAGES_NOTE)
    })

    it("adds no note when everything matching fits", async () => {
      mockClient.listChatMessages.mockResolvedValue(Right({ value: [readable("a")] }))

      const result = await listChatMessages({ chat_id: "c1", since: "2026-10-04T00:00:00Z", top: 5 })

      expect(result.value).not.toContain(CHAT_MORE_MESSAGES_NOTE)
    })

    it("passes a failure on a later page through", async () => {
      mockClient.listChatMessages.mockResolvedValue(Right({ value: [readable("a")], "@odata.nextLink": NEXT }))
      mockClient.request.mockResolvedValue(Left({ type: "api", message: "Throttled" }))

      const result = await listChatMessages({ chat_id: "c1", since: "2026-10-04T00:00:00Z", top: 5 })

      expect(result.isLeft()).toBe(true)
      expect((result.value as Error).message).toContain("Throttled")
    })

    it("takes since over fetch_all_pages", async () => {
      mockClient.listChatMessages.mockResolvedValue(Right({ value: [] }))

      await listChatMessages({ chat_id: "c1", since: "2026-10-04T00:00:00Z", fetch_all_pages: true })

      expect(mockClient.requestPaginated).not.toHaveBeenCalled()
      expect(mockClient.listChatMessages).toHaveBeenCalled()
    })
  })

  it("rejects a since without a time zone, which would be read in the server's local zone", async () => {
    const result = await listChatMessages({ chat_id: "c1", since: "2026-10-01T00:00:00" })

    expect(result.isLeft()).toBe(true)
    expect((result.value as Error).message).toContain("time zone")
  })

  it("rejects a since that is not a date-time before calling Graph", async () => {
    const result = await listChatMessages({ chat_id: "c1", since: "yesterday or true" })

    expect(result.isLeft()).toBe(true)
    expect(mockClient.getMe).not.toHaveBeenCalled()
  })

  it("still lists messages when the signed-in user cannot be resolved, and says so once", async () => {
    mockClient.getMe.mockResolvedValue(Left({ type: "forbidden", message: "Insufficient privileges" }))
    mockClient.listChatMessages.mockResolvedValue(Right({ value: [message] }))

    const result = await listChatMessages({ chat_id: "c1" })

    expect(result.isRight()).toBe(true)
    expect(result.value).not.toContain("[You]")
    expect((result.value as string).endsWith(CHAT_SELF_UNRESOLVED_NOTE)).toBe(true)
  })
})

describe("searchFiles", () => {
  it("returns at most 25 results by default, or the top asked for", async () => {
    mockClient.searchFiles.mockResolvedValue(Right({ value: [] }))

    await searchFiles({ query: "ONC" })
    await searchFiles({ query: "ONC", top: 5 })

    expect(mockClient.searchFiles).toHaveBeenNthCalledWith(1, "ONC", { $top: 25 })
    expect(mockClient.searchFiles).toHaveBeenNthCalledWith(2, "ONC", { $top: 5 })
  })

  it("pages through every match when fetch_all_pages is set", async () => {
    mockClient.searchAllFiles.mockResolvedValue(Right([searchFolder]))

    const result = await searchFiles({ query: "ONC", fetch_all_pages: true })

    expect(mockClient.searchAllFiles).toHaveBeenCalledWith("ONC")
    expect(mockClient.searchFiles).not.toHaveBeenCalled()
    expect(result.value).toContain("- **Reports** (ID: f1) - Folder")
  })

  it("prints a folder without the zero count and size search reports", async () => {
    mockClient.searchFiles.mockResolvedValue(Right({ value: [searchFolder] }))

    const result = await searchFiles({ query: "Reports" })

    expect(result.value).toContain("- **Reports** (ID: f1) - Folder")
    expect(result.value).not.toContain("0 items")
    expect(result.value).not.toContain("0 B")
  })
})

describe("searchSiteFiles", () => {
  it("returns at most 25 results by default and passes the drive", async () => {
    mockClient.searchSiteFiles.mockResolvedValue(Right({ value: [searchFolder] }))

    const result = await searchSiteFiles({ site_id: "s1", query: "Annual_Reports", drive_id: "d1" })

    expect(mockClient.searchSiteFiles).toHaveBeenCalledWith("s1", "Annual_Reports", "d1", { $top: 25 })
    expect(result.value).not.toContain("0 items")
  })

  it("pages through every match when fetch_all_pages is set", async () => {
    mockClient.searchAllSiteFiles.mockResolvedValue(Right([]))

    await searchSiteFiles({ site_id: "s1", query: "Annual_Reports", fetch_all_pages: true })

    expect(mockClient.searchAllSiteFiles).toHaveBeenCalledWith("s1", "Annual_Reports", undefined)
  })
})
