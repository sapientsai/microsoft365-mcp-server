// Graph returns a pre-authenticated download URL with a drive item. Its token lets anyone holding the
// URL fetch the file without signing in until it expires, and tool output lands in transcripts. So
// only download_file, which exists to hand it over, prints it. Every other tool leaves it out.
import { Right } from "functype/either"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

vi.mock("../src/auth/token-context", () => ({ getContextToken: vi.fn(() => "GRAPH.TOKEN") }))

import { initializeGraphClient } from "../src/client/graph-client"
import { moveDriveItem, renameDriveItem } from "../src/tools/drive-item-tools"
import { createFolder, downloadFile, getDriveItem, uploadFile } from "../src/tools/files-tools"
import { formatDriveItemDetail } from "../src/utils/formatters"

const SECRET = "made-up-tempauth-token"
const DOWNLOAD_URL = `https://example.sharepoint.com/_layouts/15/download.aspx?UniqueId=abc&tempauth=${SECRET}`
const item = {
  id: "ITEM1",
  name: "report.pdf",
  eTag: '"{5A1F},4"',
  size: 10,
  file: { mimeType: "application/pdf" },
  parentReference: { id: "PARENT", driveId: "b!drive" },
  "@microsoft.graph.downloadUrl": DOWNLOAD_URL,
}
const json = (value: unknown) =>
  new Response(JSON.stringify(value), { status: 200, headers: { "Content-Type": "application/json" } })
const text = (result: { value: unknown }) =>
  typeof result.value === "string" ? result.value : (result.value as Error).message

beforeEach(() => {
  vi.stubEnv("MS365_GRAPH_MAX_RETRIES", "0")
  vi.stubEnv("MS365_GRAPH_CIRCUIT_DISABLED", "true")
  vi.stubGlobal(
    "fetch",
    vi.fn(() => Promise.resolve(json(item))),
  )
  initializeGraphClient({ getAccessToken: async () => Right("TOKEN") })
})
afterEach(() => {
  vi.unstubAllGlobals()
  vi.unstubAllEnvs()
})

describe("the drive item formatter", () => {
  it("leaves the download URL out unless asked", () => {
    expect(formatDriveItemDetail(item)).not.toContain(SECRET)
    expect(formatDriveItemDetail(item, { includeDownloadUrl: true })).toContain(`- Download URL: ${DOWNLOAD_URL}`)
  })
})

describe("tools whose job isn't the download", () => {
  it.each([
    ["get_drive_item", () => getDriveItem({ item_id: "ITEM1" })],
    ["rename_drive_item", () => renameDriveItem({ item_id: "ITEM1", new_name: "renamed.pdf" })],
    ["move_drive_item", () => moveDriveItem({ item_id: "ITEM1", destination_folder_id: "FOLDER2" })],
    ["create_folder", () => createFolder({ parent_id: "PARENT", name: "New" })],
    ["upload_file", () => uploadFile({ path: "/me/drive/root:/a.txt:/content", content: "hello" })],
  ])("%s doesn't print the download URL or its token", async (_, run) => {
    const result = await run()
    expect(result.isRight()).toBe(true)
    expect(text(result)).not.toContain(SECRET)
    expect(text(result)).not.toContain("Download URL")
  })
})

describe("download_file", () => {
  it("still returns the download URL, which is what it's for", async () => {
    const result = await downloadFile({ item_id: "ITEM1" })
    expect(text(result)).toContain(`- Download URL: ${DOWNLOAD_URL}`)
  })
})
