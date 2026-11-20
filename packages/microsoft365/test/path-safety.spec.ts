// The ID-in-path cases found in review, each as it was reached: an ID that fetch would resolve,
// cut at "#", or split into more segments used to send the request somewhere other than the tool
// named. Each must now reach Graph only where the tool says, or not at all.
import { None, Some } from "functype"
import { Right } from "functype/either"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

vi.mock("../src/auth/auth-manager", async (original) => ({
  ...(await original<typeof import("../src/auth/auth-manager")>()),
  getAuthMode: vi.fn(),
  getAccessToken: vi.fn(),
}))

import { getAccessToken, getAuthMode } from "../src/auth/auth-manager"
import { initializeGraphClient } from "../src/client/graph-client"
import { deleteContact } from "../src/tools/contacts-tools"
import { listDriveItems, uploadFile } from "../src/tools/files-tools"
import { createForwardDraft, createReplyDraft, getMessage } from "../src/tools/mail-tools"
import { getOnenotePageContent } from "../src/tools/onenote-tools"
import { getUser } from "../src/tools/users-tools"

const jwt = (claims: object) =>
  ["{}", JSON.stringify(claims), ""].map((part) => Buffer.from(part).toString("base64url")).join(".")

const json = (value: unknown) =>
  new Response(JSON.stringify(value), { status: 200, headers: { "Content-Type": "application/json" } })
const fetchSpy = vi.fn((_url: string, _init?: RequestInit) => Promise.resolve(json({ id: "x", value: [] })))
const sentPaths = () => fetchSpy.mock.calls.map(([url]) => new URL(url).pathname)
const message = (result: { value: unknown }) => (result.value as Error).message

beforeEach(() => {
  fetchSpy.mockClear()
  vi.stubGlobal("fetch", fetchSpy)
  vi.stubEnv("MS365_GRAPH_MAX_RETRIES", "0")
  vi.mocked(getAuthMode).mockReturnValue(None())
  vi.mocked(getAccessToken).mockResolvedValue(Right(jwt({ scp: "Mail.ReadWrite" })))
  initializeGraphClient({ getAccessToken: () => Promise.resolve(Right("TOKEN")) })
})
afterEach(() => {
  vi.unstubAllGlobals()
  vi.unstubAllEnvs()
})

describe("the draft rule", () => {
  beforeEach(() => vi.stubEnv("MS365_REQUIRE_DRAFT", "true"))

  // "#" cut the path to /messages/AAA/forward, which sends at once: the request the hidden
  // send_forward tool makes, byte for byte.
  it("create_forward_draft can't be turned into a send", async () => {
    const result = await createForwardDraft({ message_id: "AAA/forward#", to: "a@example.com" })
    expect(result.isLeft()).toBe(true)
    expect(fetchSpy).not.toHaveBeenCalled()
  })

  it.each(["AAA/reply#", "AAA/reply?"])("create_reply_draft can't be turned into a send with %s", async (id) => {
    expect((await createReplyDraft({ message_id: id })).isLeft()).toBe(true)
    expect(fetchSpy).not.toHaveBeenCalled()
  })

  it("still creates a forward draft for a real ID", async () => {
    const result = await createForwardDraft({ message_id: "AAMkAGI2-_x=", to: "a@example.com" })
    expect(result.isRight()).toBe(true)
    expect(sentPaths()).toEqual(["/v1.0/me/messages/AAMkAGI2-_x=/createForward"])
  })
})

describe("the mailbox allowlist in app-only mode", () => {
  beforeEach(() => {
    vi.stubEnv("MS365_ALLOWED_MAILBOXES", "allowed@example.com")
    vi.mocked(getAuthMode).mockReturnValue(Some("client-secret"))
    vi.mocked(getAccessToken).mockResolvedValue(Right(jwt({ roles: ["Mail.ReadWrite"] })))
  })

  it.each(["../../users/victim@example.com/messages/A", "..%2F..%2Fusers%2Fvictim@example.com", "..", "."])(
    "get_message with %s never leaves the allowed mailbox",
    async (id) => {
      await getMessage({ message_id: id, mailbox: "allowed@example.com" })
      // Refused, or sent as one encoded segment under the allowed mailbox, never under another.
      const oneMessage = /^\/v1\.0\/users\/allowed%40example\.com\/messages\/[^/]+$/
      expect(sentPaths().every((path) => oneMessage.test(path))).toBe(true)
    },
  )
})

describe("tools that kept their method but reached any resource", () => {
  it("delete_contact can't delete a message", async () => {
    const result = await deleteContact({ contact_id: "../messages/A" })
    expect(result.isLeft()).toBe(true)
    expect(fetchSpy).not.toHaveBeenCalled()
  })

  it("get_onenote_page_content can't read another mailbox's raw message", async () => {
    const result = await getOnenotePageContent({ page_id: "../../../users/v@example.com/messages/A/$value#" })
    expect(result.isLeft()).toBe(true)
    expect(fetchSpy).not.toHaveBeenCalled()
  })

  it("get_user can't reach a user's mail, and a guest UPN keeps its #EXT#", async () => {
    expect((await getUser({ user_id: "v@example.com/messages" })).isLeft()).toBe(true)
    expect(fetchSpy).not.toHaveBeenCalled()
    await getUser({ user_id: "ada_contoso.com#EXT#@t.onmicrosoft.com" })
    expect(sentPaths()).toEqual(["/v1.0/users/ada_contoso.com%23EXT%23@t.onmicrosoft.com"])
  })

  it("list_drive_items can't climb out of the drive with folder_path", async () => {
    const result = await listDriveItems({ folder_path: "../../../users/v@example.com/messages?" })
    expect(result.isLeft()).toBe(true)
    expect(fetchSpy).not.toHaveBeenCalled()
  })

  it("list_drive_items encodes each piece of a folder_path", async () => {
    await listDriveItems({ folder_path: "/Reports/Q1 #2/" })
    expect(sentPaths()).toEqual(["/v1.0/me/drive/root:/Reports/Q1%20%232:/children"])
  })

  it.each(["/me/sendMail", "/me/messages", "/me/drive/root:/a.txt:/content?x=1"])(
    "upload_file can't PUT text to %s",
    async (path) => {
      const result = await uploadFile({ path, content: "x" })
      expect(result.isLeft()).toBe(true)
      expect(fetchSpy).not.toHaveBeenCalled()
    },
  )

  it("keeps a refused ID on one line", async () => {
    const result = await deleteContact({ contact_id: "A/\nForged: done" })
    expect(message(result).split("\n")).toHaveLength(1)
  })
})
