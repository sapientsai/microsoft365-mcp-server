import { mkdir, readdir, readFile, rm, stat, writeFile } from "node:fs/promises"
import { tmpdir } from "node:os"
import { join } from "node:path"

import { Some } from "functype"
import { Left, Right } from "functype/either"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

vi.mock("../src/client/graph-client", () => ({ getGraphClient: vi.fn() }))
vi.mock("node:fs/promises", async (importOriginal) => {
  const actual = await importOriginal<typeof import("node:fs/promises")>()
  return { ...actual, writeFile: vi.fn(actual.writeFile) }
})

import { getGraphClient } from "../src/client/graph-client"
import { createContact, deleteContact, exportContacts, updateContact } from "../src/tools/contacts-tools"

const mockClient = {
  requestPaginated: vi.fn(),
  createContact: vi.fn(),
  updateContact: vi.fn(),
  getContact: vi.fn(),
  deleteContact: vi.fn(),
}

const pages: Record<string, ReadonlyArray<Record<string, unknown>>> = {
  "/me/contactFolders": [{ id: "F1", displayName: "Family" }],
  "/me/contactFolders/F1/childFolders": [{ id: "F2", displayName: "Cousins" }],
  "/me/contactFolders/F2/childFolders": [],
  "/me/contacts": [{ id: "C1", displayName: "Default Person", categories: ["Friends"] }],
  "/me/contactFolders/F1/contacts": [{ id: "C2", displayName: "Family Person" }],
  "/me/contactFolders/F2/contacts": [{ id: "C3", displayName: "Cousin Person" }],
}

let outDir: string

beforeEach(() => {
  vi.clearAllMocks()
  vi.mocked(getGraphClient).mockReturnValue(Some(mockClient as never))
  mockClient.requestPaginated.mockImplementation(async (path: string) => Right(pages[path] ?? []))
  outDir = join(tmpdir(), `export-contacts-test-${Date.now()}-${Math.random().toString(36).slice(2)}`)
})

afterEach(async () => {
  vi.useRealTimers()
  await rm(outDir, { recursive: true, force: true })
})

const filesIn = async (dir: string) => readdir(dir).catch(() => [] as string[])

const readExport = async (message: string) => {
  const path = /Saved: (.+)/.exec(message)?.[1]
  expect(path).toBeDefined()
  return JSON.parse(await readFile(path as string, "utf-8")) as {
    folders: ReadonlyArray<{ path: string; count: number }>
    contacts: ReadonlyArray<{ id: string; folderPath: string }>
  }
}

describe("exportContacts", () => {
  it("exports the default folder and every nested contact folder", async () => {
    const result = await exportContacts({ out_dir: outDir })

    expect(result.isRight()).toBe(true)
    expect(result.value).toContain("Contacts: 3")
    const exported = await readExport(result.value as string)
    expect(exported.contacts.map((c) => [c.id, c.folderPath])).toEqual([
      ["C1", "Contacts"],
      ["C2", "Contacts/Family"],
      ["C3", "Contacts/Family/Cousins"],
    ])
    expect(exported.folders.map((f) => f.count)).toEqual([1, 1, 1])
  })

  it("writes the export readable only by its owner", async () => {
    const result = await exportContacts({ out_dir: outDir })
    const path = (result.value as string).match(/^Saved: (.+)$/m)?.[1]

    expect(((await stat(path as string)).mode & 0o777).toString(8)).toBe("600")
  })

  // The default page size would hit requestPaginated's 50-page cap at 500 contacts and
  // truncate without an error, so every contact listing must ask for full pages.
  it("requests full pages with the complete field set", async () => {
    await exportContacts({ out_dir: outDir })

    const contactCalls = mockClient.requestPaginated.mock.calls.filter(([path]) => String(path).endsWith("/contacts"))
    expect(contactCalls).toHaveLength(3)
    for (const [, options] of contactCalls) {
      expect(options.odataParams.$top).toBe(999)
      expect(options.odataParams.$select).toContain("categories")
      expect(options.odataParams.$select).toContain("emailAddresses")
    }
  })

  // requestPaginated returns what it has at its 50-page cap with no error, so a folder that
  // fills every page has to be flagged here or the export looks complete.
  it("flags a folder that hit the paging limit", async () => {
    const full = Array.from({ length: 999 * 50 }, (_, i) => ({ id: `D${i}` }))
    mockClient.requestPaginated.mockImplementation(async (path: string) =>
      Right(path === "/me/contacts" ? full : (pages[path] ?? [])),
    )

    const result = await exportContacts({ out_dir: outDir })

    expect(result.isRight()).toBe(true)
    expect(result.value).toContain("  Contacts: 49950 (may be incomplete: hit the paging limit)")
    expect(result.value).toContain("WARNING: stopped at the 49950-contact paging limit in Contacts.")
    const exported = (await readExport(result.value as string)) as unknown as {
      folders: ReadonlyArray<{ path: string; truncated?: boolean }>
    }
    expect(exported.folders.map((f) => [f.path, f.truncated])).toEqual([
      ["Contacts", true],
      ["Contacts/Family", undefined],
      ["Contacts/Family/Cousins", undefined],
    ])
  })

  it("does not warn when no folder hit the paging limit", async () => {
    const result = await exportContacts({ out_dir: outDir })

    expect(result.value).not.toContain("WARNING")
    expect(result.value).not.toContain("may be incomplete")
  })

  it("exports a folder reached twice only once", async () => {
    mockClient.requestPaginated.mockImplementation(async (path: string) =>
      Right(
        path === "/me/contactFolders"
          ? [
              { id: "F1", displayName: "Family" },
              { id: "F2", displayName: "Cousins" },
            ]
          : (pages[path] ?? []),
      ),
    )

    const result = await exportContacts({ out_dir: outDir })

    const exported = await readExport(result.value as string)
    expect(exported.contacts.filter((c) => c.id === "C3")).toHaveLength(1)
  })

  it("stops at a folder cycle instead of recursing forever", async () => {
    mockClient.requestPaginated.mockImplementation(async (path: string) =>
      Right(
        path === "/me/contactFolders/F2/childFolders" ? [{ id: "F1", displayName: "Family" }] : (pages[path] ?? []),
      ),
    )

    const result = await exportContacts({ out_dir: outDir })

    expect(result.isRight()).toBe(true)
    const exported = await readExport(result.value as string)
    expect(exported.contacts.map((c) => [c.id, c.folderPath])).toEqual([
      ["C1", "Contacts"],
      ["C2", "Contacts/Family"],
      ["C3", "Contacts/Family/Cousins"],
    ])
    const childFolderCalls = mockClient.requestPaginated.mock.calls.filter(([path]) =>
      String(path).endsWith("/childFolders"),
    )
    expect(childFolderCalls.map(([path]) => path)).toEqual([
      "/me/contactFolders/F1/childFolders",
      "/me/contactFolders/F2/childFolders",
    ])
  })

  it("fails rather than writing a partial export when a folder cannot be read", async () => {
    mockClient.requestPaginated.mockImplementation(async (path: string) =>
      path === "/me/contactFolders/F1/contacts"
        ? Left({ type: "http", message: "Access is denied" })
        : Right(pages[path] ?? []),
    )

    const result = await exportContacts({ out_dir: outDir })

    expect(result.isLeft()).toBe(true)
    expect((result.value as Error).message).toContain("Contacts/Family")
    expect(await filesIn(outDir)).toEqual([])
  })

  it("leaves no file behind when the write fails part-way", async () => {
    vi.mocked(writeFile).mockImplementationOnce(async (path, _data, options) => {
      const actual = await vi.importActual<typeof import("node:fs/promises")>("node:fs/promises")
      await actual.writeFile(path, '{"exportedAt":', options)
      throw Object.assign(new Error("ENOSPC: no space left on device"), { code: "ENOSPC" })
    })

    const result = await exportContacts({ out_dir: outDir })

    expect(result.isLeft()).toBe(true)
    expect((result.value as Error).message).toContain("no space left on device")
    expect(await filesIn(outDir)).toEqual([])
  })

  it("never overwrites an existing export", async () => {
    vi.useFakeTimers({ toFake: ["Date"] })
    vi.setSystemTime(new Date("2026-01-02T03:04:05.006Z"))
    const existing = join(outDir, "contacts-2026-01-02T03-04-05-006Z.json")
    await mkdir(outDir, { recursive: true })
    await writeFile(existing, "earlier export")

    const result = await exportContacts({ out_dir: outDir })

    expect(result.isLeft()).toBe(true)
    expect(await readFile(existing, "utf-8")).toBe("earlier export")
    expect(await filesIn(outDir)).toEqual(["contacts-2026-01-02T03-04-05-006Z.json"])
  })

  it("leaves only the finished export in the directory", async () => {
    const result = await exportContacts({ out_dir: outDir })

    const path = /Saved: (.+)/.exec(result.value as string)?.[1] as string
    expect(await filesIn(outDir)).toEqual([path.slice(outDir.length + 1)])
  })
})

describe("updateContact", () => {
  beforeEach(() => {
    mockClient.updateContact.mockImplementation(async (id: string, body: Record<string, unknown>) =>
      Right({ id, displayName: "Someone", ...body }),
    )
  })

  it("sends only the fields that were passed", async () => {
    await updateContact({ contact_id: "C1", categories: ["Friends"], mobile_phone: "0400 000 000" })

    expect(mockClient.updateContact).toHaveBeenCalledWith("C1", {
      categories: ["Friends"],
      mobilePhone: "0400 000 000",
    })
  })

  // Graph replaces list properties wholesale, so an empty list is how a caller clears one;
  // dropping it as "nothing to send" would make clearing impossible.
  it("sends an empty list to clear a list property", async () => {
    await updateContact({ contact_id: "C1", categories: [] })

    expect(mockClient.updateContact).toHaveBeenCalledWith("C1", { categories: [] })
  })

  it("wraps email addresses in the shape Graph expects", async () => {
    await updateContact({ contact_id: "C1", email_addresses: ["a@example.com", "b@example.com"] })

    expect(mockClient.updateContact).toHaveBeenCalledWith("C1", {
      emailAddresses: [{ address: "a@example.com" }, { address: "b@example.com" }],
    })
  })

  it("sends a birthday given as a date", async () => {
    await updateContact({ contact_id: "C1", birthday: "1974-07-22" })

    expect(mockClient.updateContact).toHaveBeenCalledWith("C1", { birthday: "1974-07-22" })
  })

  // Merging a duplicate means copying its birthday exactly as Graph returned it; trimming the
  // timestamp to a date first could move it a day.
  it("passes a birthday timestamp through unchanged", async () => {
    await updateContact({ contact_id: "C1", birthday: "1974-07-22T11:59:00Z" })

    expect(mockClient.updateContact).toHaveBeenCalledWith("C1", { birthday: "1974-07-22T11:59:00Z" })
  })

  it.each(["22/07/1974", "1974-02-30", "1974-13-01", "1974-07-22T25:00:00Z", "1974-07-22T11:59:00", "soon"])(
    "refuses the invalid birthday %s without calling Graph",
    async (birthday) => {
      const result = await updateContact({ contact_id: "C1", birthday })

      expect(result.isLeft()).toBe(true)
      expect((result.value as Error).message).toContain("Invalid birthday")
      expect(mockClient.updateContact).not.toHaveBeenCalled()
    },
  )

  it("refuses an update with no fields", async () => {
    const result = await updateContact({ contact_id: "C1" })

    expect(result.isLeft()).toBe(true)
    expect(mockClient.updateContact).not.toHaveBeenCalled()
  })
})

describe("deleteContact", () => {
  it("returns the deleted contact's details", async () => {
    mockClient.getContact.mockResolvedValue(
      Right({ id: "C1", displayName: "Dup Person", emailAddresses: [{ address: "dup@example.com" }] }),
    )
    mockClient.deleteContact.mockResolvedValue(Right({}))

    const result = await deleteContact({ contact_id: "C1" })

    expect(mockClient.deleteContact).toHaveBeenCalledWith("C1")
    expect(result.value).toContain("Dup Person")
    expect(result.value).toContain("dup@example.com")
  })

  // delete_contact tells the caller to carry the birthday over, so it must be in what comes back.
  it("includes the deleted contact's birthday", async () => {
    mockClient.getContact.mockResolvedValue(
      Right({ id: "C1", displayName: "Dup Person", birthday: "1974-07-22T11:59:00Z" }),
    )
    mockClient.deleteContact.mockResolvedValue(Right({}))

    const result = await deleteContact({ contact_id: "C1" })

    expect(result.value).toContain("Birthday: 1974-07-22T11:59:00Z")
  })

  it("does not delete when the contact cannot be read", async () => {
    mockClient.getContact.mockResolvedValue(Left({ type: "http", message: "Not found" }))

    const result = await deleteContact({ contact_id: "C1" })

    expect(result.isLeft()).toBe(true)
    expect(mockClient.deleteContact).not.toHaveBeenCalled()
  })
})

describe("contact writes without Contacts.ReadWrite", () => {
  // Graph answers a write the sign-in has no permission for with 403 "Access is denied".
  const forbidden = Left({
    type: "forbidden",
    status: 403,
    message: "Access is denied. Check credentials and try again.",
  })

  it("names the missing permission when Graph refuses a create", async () => {
    mockClient.createContact.mockResolvedValue(forbidden)
    const result = await createContact({ given_name: "Ada" })
    expect(result.isLeft()).toBe(true)
    expect((result.value as Error).message).toContain("needs the Contacts.ReadWrite permission")
  })

  it("names the missing permission when Graph refuses an update or a delete", async () => {
    mockClient.updateContact.mockResolvedValue(forbidden)
    const update = await updateContact({ contact_id: "C1", job_title: "x" })
    expect((update.value as Error).message).toContain("Contacts.ReadWrite")

    mockClient.getContact.mockResolvedValue(Right({ id: "C1", displayName: "Ada" }))
    mockClient.deleteContact.mockResolvedValue(forbidden)
    const removal = await deleteContact({ contact_id: "C1" })
    expect((removal.value as Error).message).toContain("Contacts.ReadWrite")
  })

  it("keeps Graph's own message for other failures", async () => {
    mockClient.createContact.mockResolvedValue(Left({ type: "api", status: 400, message: "Bad name" }))
    const result = await createContact({ given_name: "Ada" })
    expect((result.value as Error).message).toBe("Failed to create contact: Bad name")
  })
})
