import { beforeEach, describe, expect, it } from "vitest"

import type { GraphMessage } from "../src/types"
import { formatMessageScan, formatMessageScanRow } from "../src/utils/formatters"
import {
  clearMessageRefs,
  describeRefFailure,
  messageRefCount,
  rememberMessageId,
  resolveMessageIdOrRef,
  resolveMessageRef,
} from "../src/utils/message-refs"

const message = (overrides: Partial<GraphMessage> = {}): GraphMessage => ({
  id: "AAMkAGI0YjA3OTNhLWY2MDEtNGZlYy1hNzU2LTE4NDFiODg5ZjliMgBGAAAAAABtR",
  subject: "Test Subject",
  from: { emailAddress: { name: "John", address: "john@example.com" } },
  receivedDateTime: "2026-01-15T10:30:45Z",
  isRead: true,
  hasAttachments: false,
  ...overrides,
})

describe("message refs", () => {
  beforeEach(() => clearMessageRefs())

  it("assigns stable refs, reusing the same ref for the same id", () => {
    const first = rememberMessageId("id-a")
    const second = rememberMessageId("id-b")

    expect(rememberMessageId("id-a")).toBe(first)
    expect(second).not.toBe(first)
    expect(messageRefCount()).toBe(2)
  })

  it("issues refs as a per-process prefix plus a number that only goes up", () => {
    const [prefix, first] = rememberMessageId("id-a").split("-")
    const [samePrefix, second] = rememberMessageId("id-b").split("-")

    expect(prefix).toMatch(/^[a-z0-9]{4}$/)
    expect(samePrefix).toBe(prefix)
    expect(Number(second)).toBe(Number(first) + 1)

    clearMessageRefs()
    const [, afterClear] = rememberMessageId("id-c").split("-")
    expect(Number(afterClear)).toBeGreaterThan(Number(second))
  })

  // The prefix is what protects a restart: there the counter starts over too, so the
  // same number comes back and only the prefix tells the two refs apart.
  it("does not resolve the old prefix with a number that is live in the new process", () => {
    clearMessageRefs("aaaa")
    rememberMessageId("id-old")

    clearMessageRefs("bbbb")
    const [newPrefix, number] = rememberMessageId("id-new").split("-")

    expect(newPrefix).toBe("bbbb")
    expect(resolveMessageIdOrRef(`aaaa-${number}`)).toMatchObject({ kind: "unknown" })
    expect(resolveMessageIdOrRef(`bbbb-${number}`)).toEqual({ kind: "id", id: "id-new" })
  })

  it("round-trips a ref back to its message id, ignoring case and spaces", () => {
    const ref = rememberMessageId("id-a")
    expect(resolveMessageRef(ref)).toEqual({ kind: "id", id: "id-a" })
    expect(resolveMessageIdOrRef(` ${ref.toUpperCase()} `)).toEqual({ kind: "id", id: "id-a" })
  })

  it("passes a full Graph id straight through", () => {
    const graphId = "AAMkAGI0YjA3OTNhLWY2MDEtNGZlYy1hNzU2LTE4NDFiODg5ZjliMg=="
    expect(resolveMessageIdOrRef(graphId)).toEqual({ kind: "id", id: graphId })
  })

  // The dangerous failure is resolving to the wrong message rather than to nothing,
  // so an unknown ref must be reported as unknown for the caller to turn into an error.
  it("reports a ref that was never issued", () => {
    const ref = rememberMessageId("id-a")
    const prefix = ref.split("-")[0]!
    expect(resolveMessageIdOrRef(`${prefix}-999`)).toEqual({ kind: "unknown", ref: `${prefix}-999` })
  })

  // Numbering used to come from the map's size and restart at 1 with the process, so a
  // ref held across a restart opened whichever message took its number next. This only
  // clears the maps; the test above covers the prefix change a real restart brings.
  it("does not resolve a ref from before a clear to a newer message", () => {
    const before = rememberMessageId("id-old")
    clearMessageRefs()
    rememberMessageId("id-new")

    expect(resolveMessageIdOrRef(before)).toMatchObject({ kind: "unknown" })
  })

  it("treats a bare number as a stale ref, never as a message id", () => {
    rememberMessageId("id-a")
    expect(resolveMessageIdOrRef("1")).toEqual({ kind: "unknown", ref: "1" })
  })

  // A message id means nothing outside its own mailbox: a ref minted while scanning a
  // delegated mailbox must not silently resolve against the signed-in user's.
  it("keeps refs from different mailboxes apart", () => {
    const own = rememberMessageId("id-own")
    const delegated = rememberMessageId("id-bel", "bel@example.com")

    expect(own).not.toBe(delegated)
    expect(resolveMessageRef(own)).toEqual({ kind: "id", id: "id-own" })
    expect(resolveMessageRef(delegated, "bel@example.com")).toEqual({ kind: "id", id: "id-bel" })
  })

  it("refuses a ref minted against another mailbox", () => {
    const delegated = rememberMessageId("id-bel", "bel@example.com")

    expect(resolveMessageIdOrRef(delegated)).toEqual({
      kind: "wrong-mailbox",
      ref: delegated,
      mintedFor: "bel@example.com",
    })
  })

  it("explains a cross-mailbox ref in terms of both mailboxes", () => {
    const delegated = rememberMessageId("id-bel", "bel@example.com")
    const resolution = resolveMessageRef(delegated)
    if (resolution.kind === "id") throw new Error("expected a mismatch")

    const message = describeRefFailure(resolution, undefined)
    expect(message).toContain("bel@example.com")
    expect(message).toContain("your own mailbox")
  })

  // The same message id in two mailboxes is two distinct refs — collapsing them would
  // reintroduce exactly the cross-mailbox confusion the qualification prevents.
  it("issues distinct refs for the same id in different mailboxes", () => {
    expect(rememberMessageId("shared-id")).not.toBe(rememberMessageId("shared-id", "bel@example.com"))
  })
})

describe("formatMessageScanRow", () => {
  it("emits ref, minute-precision date, sender, subject and flags", () => {
    expect(formatMessageScanRow(message(), "k7q2-7")).toBe("k7q2-7|2026-01-15T10:30|John|Test Subject|")
  })

  it("marks unread and attachments", () => {
    const row = formatMessageScanRow(message({ isRead: false, hasAttachments: true }), "k7q2-1")
    expect(row.endsWith("|UA")).toBe(true)
  })

  it("falls back to the address when the sender has no display name", () => {
    const row = formatMessageScanRow(message({ from: { emailAddress: { address: "a@b.com" } } }), "k7q2-1")
    expect(row).toContain("|a@b.com|")
  })

  // A subject containing a pipe or newline would otherwise forge extra columns or rows.
  it("strips delimiters out of the subject", () => {
    const row = formatMessageScanRow(message({ subject: "a|b\nc" }), "k7q2-1")
    expect(row.split("|")).toHaveLength(5)
  })

  it("truncates a runaway subject", () => {
    const row = formatMessageScanRow(message({ subject: "x".repeat(500) }), "k7q2-1")
    expect(row.length).toBeLessThan(200)
  })

  it("handles a message with no subject", () => {
    expect(formatMessageScanRow(message({ subject: undefined }), "k7q2-1")).toContain("(No Subject)")
  })
})

describe("formatMessageScan", () => {
  it("reports an empty result plainly", () => {
    expect(formatMessageScan([], [], {})).toContain("No messages found on this page.")
  })

  it("names the scanned folder and explains the ref column", () => {
    const out = formatMessageScan([message()], ["k7q2-1"], { folder: "archive" })

    expect(out).toContain("1 in archive")
    expect(out).toContain("ref|received|from|subject|flags")
    expect(out).not.toContain("More available")
  })

  it("tells the caller how to fetch the next page", () => {
    const out = formatMessageScan([message()], ["k7q2-1"], { nextCursor: "next-page" })
    expect(out).toContain('cursor: "next-page"')
  })

  // The whole point of the format is that scanning stays affordable in bulk.
  it("stays far cheaper per row than the markdown list format", () => {
    const many = Array.from({ length: 100 }, () => message())
    const refs = many.map((_, i) => `k7q2-${i + 1}`)

    const perRow = formatMessageScan(many, refs, {}).length / 100
    expect(perRow).toBeLessThan(100)
  })
})

describe("scan coverage honesty", () => {
  it("preserves the continuation on an empty page", () => {
    const out = formatMessageScan([], [], { nextCursor: "next-page" })
    expect(out).toContain("INCOMPLETE")
    expect(out).toContain('cursor: "next-page"')
  })

  it("warns about the search cap even on the final page", () => {
    const out = formatMessageScan([message()], ["k7q2-1"], { searched: true })
    expect(out).toContain("1,000")
    expect(out).toContain("does not establish mailbox coverage")
    expect(out).toContain("filter")
    expect(out).not.toContain("cannot be paged")
  })

  it("offers the same cursor continuation for search", () => {
    const out = formatMessageScan([message()], ["k7q2-1"], { searched: true, nextCursor: "next-page" })
    expect(out).toContain("INCOMPLETE")
    expect(out).toContain('cursor: "next-page"')
    expect(out).not.toContain("skip:")
  })

  it("marks the end only when there is no continuation", () => {
    const out = formatMessageScan([message()], ["k7q2-1"], {})
    expect(out).not.toContain("INCOMPLETE")
    expect(out).toContain("End of results for this query")
  })
})
