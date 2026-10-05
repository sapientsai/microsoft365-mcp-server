import { Some } from "functype"
import { Right } from "functype/either"
import { beforeEach, describe, expect, it, vi } from "vitest"

vi.mock("../src/client/graph-client", () => ({
  getGraphClient: vi.fn(),
}))

import { getGraphClient } from "../src/client/graph-client"
import { listChats } from "../src/tools/chat-tools"
import { searchFiles } from "../src/tools/files-tools"
import { searchSiteFiles } from "../src/tools/sharepoint-tools"

const mockClient = {
  listChats: vi.fn(),
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
