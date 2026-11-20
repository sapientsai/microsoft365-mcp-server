// The settings tools take the same mailbox parameter as the other mail tools. These check it reaches
// Graph, and that a mailbox outside MS365_ALLOWED_MAILBOXES never does.

import { None, Some } from "functype"
import { Left, Right } from "functype/either"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

vi.mock("../src/client/graph-client", () => ({ getGraphClient: vi.fn() }))
vi.mock("../src/auth/auth-manager", () => ({ getAuthMode: vi.fn(), getAccessToken: vi.fn() }))

import { getAccessToken, getAuthMode } from "../src/auth/auth-manager"
import { getGraphClient } from "../src/client/graph-client"
import {
  createMailCategory,
  createMailRule,
  deleteMailRule,
  getAutomaticReplies,
  listMailCategories,
  listMailRules,
  setAutomaticReplies,
  setMessageCategories,
} from "../src/tools/mail-settings-tools"

const BEL = "bel@example.com"
const BEL_PREFIX = "/users/bel%40example.com"
const jwt = (claims: object) =>
  ["{}", JSON.stringify(claims), ""].map((part) => Buffer.from(part).toString("base64url")).join(".")

const mockClient = {
  getAutomaticReplies: vi.fn(),
  setAutomaticReplies: vi.fn(),
  listMasterCategories: vi.fn(),
  createMasterCategory: vi.fn(),
  getMessageCategories: vi.fn(),
  updateMessage: vi.fn(),
  listMessageRules: vi.fn(),
  createMessageRule: vi.fn(),
  deleteMessageRule: vi.fn(),
  listMailFolders: vi.fn(),
  getMailFolder: vi.fn(),
  request: vi.fn(),
}

beforeEach(() => {
  vi.clearAllMocks()
  process.env.MS365_ALLOWED_MAILBOXES = BEL
  vi.mocked(getGraphClient).mockReturnValue(Some(mockClient as never))
  vi.mocked(getAuthMode).mockReturnValue(None())
  vi.mocked(getAccessToken).mockResolvedValue(Right(jwt({ scp: "MailboxSettings.ReadWrite" })))
  mockClient.getAutomaticReplies.mockResolvedValue(Right({ status: "disabled" }))
  mockClient.setAutomaticReplies.mockResolvedValue(Right({}))
  mockClient.listMasterCategories.mockResolvedValue(Right({ value: [] }))
  mockClient.createMasterCategory.mockResolvedValue(Right({ id: "c1", displayName: "A", color: "none" }))
  mockClient.updateMessage.mockResolvedValue(Right({ id: "m1", categories: ["A"] }))
  mockClient.listMessageRules.mockResolvedValue(Right({ value: [] }))
  mockClient.createMessageRule.mockResolvedValue(Right({ id: "r1", displayName: "R" }))
  mockClient.deleteMessageRule.mockResolvedValue(Right({}))
  mockClient.listMailFolders.mockResolvedValue(Right({ value: [] }))
  mockClient.getMailFolder.mockResolvedValue(Right({ id: "ARCHIVE-ID" }))
  // The own-address lookup; no test here names the caller's own address.
  mockClient.request.mockResolvedValue(Left({ type: "api", message: "not under test" }))
})

afterEach(() => {
  delete process.env.MS365_ALLOWED_MAILBOXES
})

const lastPrefix = (mock: { mock: { calls: unknown[][] } }) => mock.mock.calls.at(-1)?.at(-1)

const CASES = [
  ["get_automatic_replies", () => getAutomaticReplies({ mailbox: BEL }), mockClient.getAutomaticReplies],
  [
    "set_automatic_replies",
    () => setAutomaticReplies({ status: "disabled", mailbox: BEL }),
    mockClient.setAutomaticReplies,
  ],
  ["list_mail_categories", () => listMailCategories({ mailbox: BEL }), mockClient.listMasterCategories],
  ["create_mail_category", () => createMailCategory({ name: "A", mailbox: BEL }), mockClient.createMasterCategory],
  [
    "set_message_categories",
    () => setMessageCategories({ message_id: "m1", categories: ["A"], mailbox: BEL }),
    mockClient.updateMessage,
  ],
  ["list_mail_rules", () => listMailRules({ mailbox: BEL }), mockClient.listMessageRules],
  [
    "create_mail_rule",
    () => createMailRule({ name: "R", subject_contains: ["x"], move_to_folder: "archive", mailbox: BEL }),
    mockClient.createMessageRule,
  ],
  ["delete_mail_rule", () => deleteMailRule({ rule_id: "r1", mailbox: BEL }), mockClient.deleteMessageRule],
] as const

describe("mailbox routing", () => {
  it.each(CASES)("%s addresses /users/{mailbox}", async (_name, run, mock) => {
    const result = await run()
    expect(result.isRight()).toBe(true)
    expect(lastPrefix(mock)).toBe(BEL_PREFIX)
  })

  it("looks up a rule's folder in the same mailbox", async () => {
    await createMailRule({ name: "R", subject_contains: ["x"], move_to_folder: "archive", mailbox: BEL })
    expect(lastPrefix(mockClient.getMailFolder)).toBe(BEL_PREFIX)
  })

  it.each(CASES)("%s never reaches Graph for a mailbox outside the allowlist", async (_name, run, mock) => {
    process.env.MS365_ALLOWED_MAILBOXES = "someone-else@example.com"
    const result = await run()
    expect(result.isLeft()).toBe(true)
    expect(mock).not.toHaveBeenCalled()
  })
})
