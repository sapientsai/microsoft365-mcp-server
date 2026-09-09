// Sender-level roll-up of a folder, for deciding what to sweep.
//
// A large inbox is mostly a few hundred senders repeating. Finding that out by listing
// headers costs the caller a page of context per thousand rows; counting here costs
// only Graph requests. The aggregation is pure so it can be tested without a mailbox,
// and so the tool layer stays a thin fetch-then-format.

import type { GraphMessage } from "../types"

export type SenderGroupBy = "address" | "domain"

export type SenderSummaryRow = {
  /** The sender address, or its domain when grouped by domain. Lowercased. */
  readonly key: string
  /** Display name from the most recently received message in the group. */
  readonly name: string
  readonly count: number
  readonly unread: number
  /** YYYY-MM-DD of the oldest and newest message in the group. */
  readonly first: string
  readonly last: string
  readonly latestSubject: string
}

type Accumulator = {
  name: string
  count: number
  unread: number
  first: string
  last: string
  latestSubject: string
}

const NO_SENDER = "(no sender)"

export const senderAddress = (msg: GraphMessage): string => (msg.from?.emailAddress.address ?? "").trim().toLowerCase()

export const senderDomain = (address: string): string => {
  const at = address.lastIndexOf("@")
  return at >= 0 ? address.slice(at + 1) : address
}

const groupKey = (msg: GraphMessage, groupBy: SenderGroupBy): string => {
  const address = senderAddress(msg)
  if (address.length === 0) return NO_SENDER
  return groupBy === "address" ? address : senderDomain(address)
}

// Day precision is what a sweep decision needs ("still arriving?" / "how far back?").
const day = (msg: GraphMessage): string => msg.receivedDateTime?.slice(0, 10) ?? ""

export const summarizeSenders = (
  messages: ReadonlyArray<GraphMessage>,
  groupBy: SenderGroupBy = "address",
): ReadonlyArray<SenderSummaryRow> => {
  // A mutable Map inside the fold is deliberate: this runs over tens of thousands of
  // rows, and rebuilding an immutable map per message is measurable for no clarity gain.
  const groups = messages.reduce((acc, msg) => {
    const key = groupKey(msg, groupBy)
    const received = day(msg)
    const unread = msg.isRead === false ? 1 : 0
    const existing = acc.get(key)
    if (existing === undefined) {
      acc.set(key, {
        name: msg.from?.emailAddress.name ?? "",
        count: 1,
        unread,
        first: received,
        last: received,
        latestSubject: msg.subject ?? "",
      })
      return acc
    }
    existing.count += 1
    existing.unread += unread
    if (received !== "" && (existing.first === "" || received < existing.first)) existing.first = received
    if (received >= existing.last) {
      existing.last = received
      existing.latestSubject = msg.subject ?? ""
      existing.name = msg.from?.emailAddress.name ?? existing.name
    }
    return acc
  }, new Map<string, Accumulator>())

  return [...groups.entries()]
    .map(([key, g]) => ({ key, ...g }))
    .sort((a, b) => b.count - a.count || a.key.localeCompare(b.key))
}

const cell = (value: string, max: number): string => value.replace(/[\r\n|]+/g, " ").slice(0, max)

export const formatSenderSummary = (
  rows: ReadonlyArray<SenderSummaryRow>,
  meta: {
    readonly folder: string
    readonly total: number
    readonly unread: number
    readonly groupBy: SenderGroupBy
    readonly top: number
    readonly filter?: string
  },
): string => {
  if (meta.total === 0) return `No messages in ${meta.folder}${meta.filter ? ` matching ${meta.filter}` : ""}.`

  const shown = rows.slice(0, meta.top)
  const scope = meta.filter ? ` matching "${meta.filter}"` : ""
  const label = meta.groupBy === "address" ? "sender" : "domain"
  const header = [
    `# Sender summary — ${meta.total} messages in ${meta.folder}${scope}, ${rows.length} distinct ${label}s (${meta.unread} unread)`,
    "",
    shown.length < rows.length
      ? `Showing the top ${shown.length} by count; pass a larger top for the rest.`
      : `All ${rows.length} ${label}s shown.`,
    meta.groupBy === "address"
      ? "To act on one sender, pass its address in move_messages_matching senders: [...]."
      : "Domain rows cannot be swept directly — re-run with group_by: address to get addresses.",
    "",
    `count|unread|first|last|${label}|name|latest subject`,
  ].join("\n")

  const lines = shown.map(
    (r) =>
      `${r.count}|${r.unread}|${r.first}|${r.last}|${cell(r.key, 60)}|${cell(r.name, 40)}|${cell(r.latestSubject, 80)}`,
  )

  return `${header}\n${lines.join("\n")}`
}
