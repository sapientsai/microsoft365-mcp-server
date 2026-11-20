import { Right } from "functype/either"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

import { initializeGraphClient } from "../src/client/graph-client"
import {
  copyDriveItem,
  deleteDriveItem,
  ITEM_PRECONDITION_FAILED_MESSAGE,
  moveDriveItem,
  renameDriveItem,
} from "../src/tools/drive-item-tools"
import { annotationsFor, filterTools } from "../src/tools/tool-registry"

const ETAG = '"{5A1F},3"'
const json = (value: unknown, status = 200) =>
  new Response(JSON.stringify(value), { status, headers: { "Content-Type": "application/json" } })
const item = { id: "ITEM1", name: "report.docx", eTag: '"{5A1F},4"', file: { mimeType: "application/msword" } }
const DRIVE_TOOLS = ["move_drive_item", "rename_drive_item", "copy_drive_item", "delete_drive_item"]

const stubFetch = (...responses: ReadonlyArray<Response>) => {
  const spy = vi.fn()
  responses.forEach((response) => spy.mockResolvedValueOnce(response))
  vi.stubGlobal("fetch", spy)
  return spy
}
const call = (spy: ReturnType<typeof vi.fn>, index = 0) => {
  const [url, init] = spy.mock.calls[index] as [string, RequestInit]
  return { url, method: init.method, headers: init.headers as Record<string, string>, body: init.body }
}

beforeEach(() => {
  vi.stubEnv("MS365_GRAPH_MAX_RETRIES", "0")
  vi.stubEnv("MS365_GRAPH_CIRCUIT_DISABLED", "true")
  initializeGraphClient({ getAccessToken: async () => Right("TOKEN") })
})
afterEach(() => {
  vi.unstubAllGlobals()
  vi.unstubAllEnvs()
})

describe("move_drive_item", () => {
  it("PATCHes the item's parent and name, with If-Match", async () => {
    const spy = stubFetch(json(item))
    const result = await moveDriveItem({
      item_id: "ITEM1",
      destination_folder_id: "FOLDER2",
      new_name: "final.docx",
      if_match: ETAG,
    })
    expect(result.isRight()).toBe(true)
    expect(result.value).toContain("Item moved.")
    const sent = call(spy)
    expect(sent.method).toBe("PATCH")
    expect(sent.url).toMatch(/\/me\/drive\/items\/ITEM1$/)
    expect(sent.headers).toMatchObject({ "If-Match": ETAG })
    expect(JSON.parse(sent.body as string)).toEqual({ parentReference: { id: "FOLDER2" }, name: "final.docx" })
  })

  it("looks up the real root ID of a SharePoint drive for destination root", async () => {
    const spy = stubFetch(json({ id: "ROOTID" }), json(item))
    const result = await moveDriveItem({ item_id: "ITEM1", destination_folder_id: "root", drive_id: "b!drive_1" })
    expect(result.isRight()).toBe(true)
    expect(call(spy, 0).url).toContain("/drives/b!drive_1/root")
    const patch = call(spy, 1)
    expect(patch.url).toMatch(/\/drives\/b!drive_1\/items\/ITEM1$/)
    expect(JSON.parse(patch.body as string)).toEqual({ parentReference: { id: "ROOTID" } })
    expect(patch.headers).not.toHaveProperty("If-Match")
  })

  it("reports a stale eTag as precondition-failed", async () => {
    stubFetch(json({ error: { code: "notAllowed", message: "ETag mismatch" } }, 412))
    const result = await moveDriveItem({ item_id: "ITEM1", destination_folder_id: "FOLDER2", if_match: ETAG })
    expect(result.isLeft()).toBe(true)
    expect((result.value as Error).message).toContain(ITEM_PRECONDITION_FAILED_MESSAGE)
  })
})

describe("rename_drive_item", () => {
  it("PATCHes only the name, in the given drive", async () => {
    const spy = stubFetch(json(item))
    const result = await renameDriveItem({ item_id: "ITEM1", new_name: "final.docx", drive_id: "b!drive_1" })
    expect(result.isRight()).toBe(true)
    const sent = call(spy)
    expect(sent.url).toMatch(/\/drives\/b!drive_1\/items\/ITEM1$/)
    expect(JSON.parse(sent.body as string)).toEqual({ name: "final.docx" })
  })

  it("refuses a blank name or one with a slash, without calling Graph", async () => {
    const spy = stubFetch()
    for (const new_name of ["  ", "a/b.docx", "a\\b.docx"]) {
      const result = await renameDriveItem({ item_id: "ITEM1", new_name })
      expect(result.isLeft()).toBe(true)
    }
    expect(spy).not.toHaveBeenCalled()
  })
})

describe("copy_drive_item", () => {
  it("copies into another drive and reports that Graph finishes in the background", async () => {
    const spy = stubFetch(new Response(null, { status: 202, headers: { Location: "https://monitor.example/x" } }))
    const result = await copyDriveItem({
      item_id: "ITEM1",
      drive_id: "b!source",
      destination_folder_id: "FOLDER2",
      destination_drive_id: "b!target",
      conflict_behavior: "rename",
    })
    expect(result.isRight()).toBe(true)
    expect(result.value).toContain("Copy started")
    expect(result.value).not.toContain("monitor.example")
    const sent = call(spy)
    expect(sent.method).toBe("POST")
    expect(sent.url).toMatch(/\/drives\/b!source\/items\/ITEM1\/copy\?@microsoft\.graph\.conflictBehavior=rename$/)
    expect(JSON.parse(sent.body as string)).toEqual({ parentReference: { id: "FOLDER2", driveId: "b!target" } })
  })

  it("sends no conflict setting by default, and stays in the source drive", async () => {
    const spy = stubFetch(new Response(null, { status: 202 }))
    await copyDriveItem({
      item_id: "ITEM1",
      drive_id: "b!source",
      destination_folder_id: "FOLDER2",
      new_name: "c.docx",
    })
    const sent = call(spy)
    expect(sent.url).toMatch(/\/items\/ITEM1\/copy$/)
    expect(JSON.parse(sent.body as string)).toEqual({
      parentReference: { id: "FOLDER2", driveId: "b!source" },
      name: "c.docx",
    })
  })
})

describe("delete_drive_item", () => {
  it("DELETEs with If-Match and says the item is in the recycle bin", async () => {
    const spy = stubFetch(new Response(null, { status: 204 }))
    const result = await deleteDriveItem({ item_id: "ITEM1", drive_id: "b!drive_1", if_match: ETAG })
    expect(result.isRight()).toBe(true)
    expect(result.value).toContain("recycle bin")
    const sent = call(spy)
    expect(sent.method).toBe("DELETE")
    expect(sent.url).toMatch(/\/drives\/b!drive_1\/items\/ITEM1$/)
    expect(sent.headers).toMatchObject({ "If-Match": ETAG })
  })

  it("returns an error when Graph refuses, so nothing reads as deleted", async () => {
    stubFetch(json({ error: { code: "itemNotFound", message: "Item not found" } }, 404))
    const result = await deleteDriveItem({ item_id: "ITEM1" })
    expect(result.isLeft()).toBe(true)
    expect((result.value as Error).message).toContain("Failed to delete item")
  })

  it("refuses a blank if_match rather than deleting unconditionally", async () => {
    const spy = stubFetch()
    const result = await deleteDriveItem({ item_id: "ITEM1", if_match: " " })
    expect(result.isLeft()).toBe(true)
    expect(spy).not.toHaveBeenCalled()
  })
})

describe("IDs that could reach another Graph resource", () => {
  const BAD = ["ITEM1/permissions", "..", "ITEM1?x=1", "ITEM1#x", "ITEM1%2Fchildren", " ITEM1", "a\\b"]

  it.each(BAD)("refuses item_id %j on every tool without calling Graph", async (bad) => {
    const spy = stubFetch()
    const results = await Promise.all([
      moveDriveItem({ item_id: bad, destination_folder_id: "FOLDER2" }),
      renameDriveItem({ item_id: bad, new_name: "x.txt" }),
      copyDriveItem({ item_id: bad, destination_folder_id: "FOLDER2" }),
      deleteDriveItem({ item_id: bad }),
    ])
    results.forEach((result) => expect(result.isLeft()).toBe(true))
    expect(spy).not.toHaveBeenCalled()
  })

  it.each(BAD)("refuses drive and folder IDs %j too", async (bad) => {
    const spy = stubFetch()
    const results = await Promise.all([
      deleteDriveItem({ item_id: "ITEM1", drive_id: bad }),
      moveDriveItem({ item_id: "ITEM1", destination_folder_id: bad }),
      copyDriveItem({ item_id: "ITEM1", destination_folder_id: "FOLDER2", destination_drive_id: bad }),
    ])
    results.forEach((result) => expect(result.isLeft()).toBe(true))
    expect(spy).not.toHaveBeenCalled()
  })
})

describe("registration", () => {
  it("hides all four in read-only mode", () => {
    const allowed = filterTools({ readOnly: true, orgMode: true })
    DRIVE_TOOLS.forEach((name) => expect(allowed.has(name)).toBe(false))
  })

  it("offers all four without org mode, in the personal preset", () => {
    const allowed = filterTools({ presets: ["personal"], orgMode: false })
    DRIVE_TOOLS.forEach((name) => expect(allowed.has(name)).toBe(true))
  })

  it("leaves them out of a preset without files, and out of a tool filter that doesn't match", () => {
    const productivity = filterTools({ presets: ["productivity"], orgMode: true })
    const filtered = filterTools({ enabledPattern: "^list_", orgMode: true })
    DRIVE_TOOLS.forEach((name) => {
      expect(productivity.has(name)).toBe(false)
      expect(filtered.has(name)).toBe(false)
    })
  })

  // A copy can land in a shared library, so even copy asks first.
  it("marks all four destructive, so a client asks before running them", () => {
    DRIVE_TOOLS.forEach((name) => expect(annotationsFor(name)?.destructiveHint).toBe(true))
  })
})
