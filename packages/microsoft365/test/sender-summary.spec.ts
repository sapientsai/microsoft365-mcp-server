import { describe, expect, it } from "vitest"

import { formatSenderSummary, senderDomain, summarizeSenders } from "../src/mail/sender-summary"
import type { GraphMessage } from "../src/types"

const msg = (address: string, received: string, overrides: Partial<GraphMessage> = {}): GraphMessage => ({
  id: `${address}-${received}`,
  subject: `Subject ${received}`,
  from: { emailAddress: { name: address.split("@")[0], address } },
  receivedDateTime: `${received}T09:00:00Z`,
  isRead: false,
  ...overrides,
})

describe("summarizeSenders", () => {
  it("counts per address, most frequent first", () => {
    const rows = summarizeSenders([
      msg("a@x.com", "2026-01-01"),
      msg("b@y.com", "2026-01-02"),
      msg("a@x.com", "2026-01-03"),
    ])
    expect(rows.map((r) => [r.key, r.count])).toEqual([
      ["a@x.com", 2],
      ["b@y.com", 1],
    ])
  })

  it("is case-insensitive on the address", () => {
    const rows = summarizeSenders([msg("A@X.com", "2026-01-01"), msg("a@x.com", "2026-01-02")])
    expect(rows).toHaveLength(1)
    expect(rows[0]!.key).toBe("a@x.com")
  })

  // The dates and latest subject are what tell a caller "still arriving daily" from
  // "stopped two years ago", which decides whether a sweep is worth it.
  it("tracks first and last day, unread count and the latest subject", () => {
    const rows = summarizeSenders([
      msg("a@x.com", "2026-03-01", { isRead: true }),
      msg("a@x.com", "2025-01-15"),
      msg("a@x.com", "2026-02-10"),
    ])
    expect(rows[0]).toMatchObject({
      first: "2025-01-15",
      last: "2026-03-01",
      unread: 2,
      latestSubject: "Subject 2026-03-01",
    })
  })

  it("groups by domain when asked", () => {
    const rows = summarizeSenders(
      [msg("a@x.com", "2026-01-01"), msg("b@x.com", "2026-01-02"), msg("c@y.com", "2026-01-03")],
      "domain",
    )
    expect(rows.map((r) => [r.key, r.count])).toEqual([
      ["x.com", 2],
      ["y.com", 1],
    ])
  })

  it("puts messages without a sender in their own bucket rather than dropping them", () => {
    const rows = summarizeSenders([msg("a@x.com", "2026-01-01", { from: undefined })])
    expect(rows[0]!.key).toBe("(no sender)")
  })

  it("derives the domain from the last @", () => {
    expect(senderDomain("weird@name@x.com")).toBe("x.com")
    expect(senderDomain("nodomain")).toBe("nodomain")
  })
})

describe("formatSenderSummary", () => {
  const rows = summarizeSenders([
    msg("a@x.com", "2026-01-01"),
    msg("a@x.com", "2026-01-02"),
    msg("b@y.com", "2026-01-03"),
  ])

  it("reports totals and one pipe row per sender", () => {
    const out = formatSenderSummary(rows, { folder: "inbox", total: 3, unread: 3, groupBy: "address", top: 100 })
    expect(out).toContain("3 messages in inbox")
    expect(out).toContain("2 distinct senders")
    expect(out).toContain("2|2|2026-01-01|2026-01-02|a@x.com|a|Subject 2026-01-02")
    expect(out).toContain("All 2 senders shown.")
  })

  it("says when rows were cut off by top", () => {
    const out = formatSenderSummary(rows, { folder: "inbox", total: 3, unread: 3, groupBy: "address", top: 1 })
    expect(out).toContain("Showing the top 1 by count")
    expect(out).not.toContain("b@y.com")
  })

  it("strips pipes from cells so rows stay parseable", () => {
    const piped = summarizeSenders([msg("a@x.com", "2026-01-01", { subject: "a | b" })])
    const out = formatSenderSummary(piped, { folder: "inbox", total: 1, unread: 1, groupBy: "address", top: 10 })
    expect(out).toContain("|a   b")
    expect(out).not.toContain("a | b")
  })

  it("handles an empty folder", () => {
    expect(formatSenderSummary([], { folder: "inbox", total: 0, unread: 0, groupBy: "address", top: 10 })).toBe(
      "No messages in inbox.",
    )
  })
})
