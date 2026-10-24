import { readFile, rm, stat } from "node:fs/promises"
import { tmpdir } from "node:os"
import { join } from "node:path"

import { Some } from "functype"
import { Left, Right } from "functype/either"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

vi.mock("../src/client/graph-client", () => ({ getGraphClient: vi.fn() }))

import { getGraphClient } from "../src/client/graph-client"
import { deleteContact, exportContacts, updateContact } from "../src/tools/contacts-tools"

const mockClient = { requestPaginated: vi.fn(), updateContact: vi.fn(), getContact: vi.fn(), deleteContact: vi.fn() }

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
  await rm(outDir, { recursive: true, force: true })
})

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

  it("fails rather than writing a partial export when a folder cannot be read", async () => {
    mockClient.requestPaginated.mockImplementation(async (path: string) =>
      path === "/me/contactFolders/F1/contacts"
        ? Left({ type: "http", message: "Access is denied" })
        : Right(pages[path] ?? []),
    )

    const result = await exportContacts({ out_dir: outDir })

    expect(result.isLeft()).toBe(true)
    expect((result.value as Error).message).toContain("Contacts/Family")
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

  it("does not delete when the contact cannot be read", async () => {
    mockClient.getContact.mockResolvedValue(Left({ type: "http", message: "Not found" }))

    const result = await deleteContact({ contact_id: "C1" })

    expect(result.isLeft()).toBe(true)
    expect(mockClient.deleteContact).not.toHaveBeenCalled()
  })
})
