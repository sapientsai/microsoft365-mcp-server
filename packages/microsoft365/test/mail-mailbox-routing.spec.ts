// The mailbox parameter is only useful if it reaches Graph on every mail path. These
// tests assert the routing rather than the formatting: which prefix each call was made
// with, and that an unpermitted mailbox never reaches Graph at all.

import { Some } from "functype"
import { Right } from "functype/either"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

vi.mock("../src/client/graph-client", () => ({ getGraphClient: vi.fn() }))

import { getGraphClient } from "../src/client/graph-client"
import {
  batchMoveMessages,
  getMessage,
  listAttachments,
  listMessages,
  moveMessage,
  scanMessages,
  searchMessages,
  sendReply,
} from "../src/tools/mail-tools"
import { clearMessageRefs } from "../src/utils/message-refs"

const BEL = "bel@example.com"
const BEL_PREFIX = "/users/bel%40example.com"

const message = { id: "m1", subject: "Hello", from: { emailAddress: { address: "a@b.c" } } }

const mockClient = {
  listMessages: vi.fn(),
  listFolderMessages: vi.fn(),
  getMessage: vi.fn(),
  listMailFolders: vi.fn(),
  moveMessage: vi.fn(),
  listAttachments: vi.fn(),
  searchMessages: vi.fn(),
  sendReply: vi.fn(),
  requestPaginated: vi.fn(),
}

beforeEach(() => {
  vi.clearAllMocks()
  clearMessageRefs()
  process.env.MS365_ALLOWED_MAILBOXES = BEL
  vi.mocked(getGraphClient).mockReturnValue(Some(mockClient as never))

  mockClient.listMessages.mockResolvedValue(Right({ value: [message] }))
  mockClient.listFolderMessages.mockResolvedValue(Right({ value: [message] }))
  mockClient.getMessage.mockResolvedValue(Right(message))
  mockClient.listMailFolders.mockResolvedValue(Right({ value: [{ id: "f-receipts", displayName: "Receipts" }] }))
  mockClient.moveMessage.mockResolvedValue(Right(message))
  mockClient.listAttachments.mockResolvedValue(Right({ value: [] }))
  mockClient.searchMessages.mockResolvedValue(Right({ value: [message] }))
  mockClient.sendReply.mockResolvedValue(Right({}))
  mockClient.requestPaginated.mockResolvedValue(Right([message]))
})

afterEach(() => {
  delete process.env.MS365_ALLOWED_MAILBOXES
})

const prefixOf = (mock: { mock: { calls: unknown[][] } }) => {
  const call = mock.mock.calls.at(-1)
  if (!call) throw new Error("expected the client to have been called")
  return call.at(-1)
}

describe("mailbox routing", () => {
  it("addresses /me when no mailbox is given", async () => {
    await listMessages({})
    expect(prefixOf(mockClient.listMessages)).toBe("/me")
  })

  it("addresses /users/{mailbox} when one is given", async () => {
    await listMessages({ mailbox: BEL })
    expect(prefixOf(mockClient.listMessages)).toBe(BEL_PREFIX)
  })

  it("routes the paginated path too", async () => {
    await listMessages({ mailbox: BEL, fetch_all_pages: true })
    expect(mockClient.requestPaginated).toHaveBeenCalledWith(`${BEL_PREFIX}/messages`, expect.anything())
  })

  it("routes get, search, attachments and reply", async () => {
    await getMessage({ message_id: "m1", mailbox: BEL })
    expect(prefixOf(mockClient.getMessage)).toBe(BEL_PREFIX)

    await searchMessages({ query: "invoice", mailbox: BEL })
    expect(prefixOf(mockClient.searchMessages)).toBe(BEL_PREFIX)

    await listAttachments({ message_id: "m1", mailbox: BEL })
    expect(prefixOf(mockClient.listAttachments)).toBe(BEL_PREFIX)

    await sendReply({ message_id: "m1", comment: "ok", mailbox: BEL })
    expect(prefixOf(mockClient.sendReply)).toBe(BEL_PREFIX)
  })

  // The subtle one: a folder display name resolves to an ID that only exists in one
  // mailbox. Looking it up against /me and then moving in Bel's mailbox would target a
  // folder that is missing there — or, worse, a different folder that happens to exist.
  it("resolves a destination folder name in the mailbox being addressed", async () => {
    await moveMessage({ message_id: "m1", destination: "Receipts", mailbox: BEL })

    expect(prefixOf(mockClient.listMailFolders)).toBe(BEL_PREFIX)
    expect(mockClient.moveMessage).toHaveBeenCalledWith("m1", "f-receipts", BEL_PREFIX)
  })

  it("resolves a scan folder in the mailbox being addressed", async () => {
    await scanMessages({ folder: "Receipts", mailbox: BEL })

    expect(prefixOf(mockClient.listMailFolders)).toBe(BEL_PREFIX)
    expect(mockClient.listFolderMessages).toHaveBeenCalledWith("f-receipts", expect.anything(), BEL_PREFIX)
  })

  it("routes every message in a batch move", async () => {
    await batchMoveMessages({ message_ids: ["m1", "m2"], destination: "archive", mailbox: BEL })

    expect(mockClient.moveMessage).toHaveBeenCalledTimes(2)
    for (const call of mockClient.moveMessage.mock.calls) expect(call.at(-1)).toBe(BEL_PREFIX)
  })
})

describe("mailbox permission", () => {
  it("refuses a mailbox that is not allowed, without calling Graph", async () => {
    const result = await listMessages({ mailbox: "someone@example.com" })

    expect(result.isLeft()).toBe(true)
    expect(mockClient.listMessages).not.toHaveBeenCalled()
  })

  it("refuses a write to a mailbox that is not allowed", async () => {
    const result = await moveMessage({ message_id: "m1", destination: "archive", mailbox: "someone@example.com" })

    expect(result.isLeft()).toBe(true)
    expect(mockClient.moveMessage).not.toHaveBeenCalled()
  })
})

describe("cross-mailbox refs", () => {
  // A ref is a bare number, so nothing about it tells the caller which mailbox it came
  // from. Resolving one against the wrong mailbox would fetch an unrelated message.
  it("refuses a ref minted in another mailbox", async () => {
    await scanMessages({ mailbox: BEL })

    const result = await getMessage({ message_id: "1" })

    expect(result.isLeft()).toBe(true)
    expect((result.value as { message: string }).message).toContain(BEL)
    expect(mockClient.getMessage).not.toHaveBeenCalled()
  })

  it("accepts the ref when the same mailbox is addressed", async () => {
    await scanMessages({ mailbox: BEL })

    const result = await getMessage({ message_id: "1", mailbox: BEL })

    expect(result.isRight()).toBe(true)
    expect(mockClient.getMessage).toHaveBeenCalledWith("m1", undefined, BEL_PREFIX)
  })

  // The row-per-failure reporting is the point: the caller learns which ref was stale. When that
  // stale ref is the ONLY message in the batch, nothing moved, and the call now reports Left —
  // a batch where no message moved is a failure, not a success carrying bad news in its text.
  it("reports a stale ref in a batch move as one failed row, not a failed batch", async () => {
    const result = await batchMoveMessages({ message_ids: ["999", "m1"], destination: "archive" })

    expect(result.isRight()).toBe(true)
    expect(result.value as string).toContain("FAILED")
    expect(mockClient.moveMessage).toHaveBeenCalledTimes(1)
  })

  it("fails the batch when the only message was a stale ref, since nothing moved", async () => {
    const result = await batchMoveMessages({ message_ids: ["999"], destination: "archive" })

    expect(result.isLeft()).toBe(true)
    expect((result.value as Error).message).toContain("FAILED")
    expect(mockClient.moveMessage).not.toHaveBeenCalled()
  })
})
