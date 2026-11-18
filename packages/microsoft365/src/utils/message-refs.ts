import { randomInt } from "node:crypto"

// Graph message IDs are ~152 characters. When a caller is scanning thousands of
// messages to decide which few are worth opening, those IDs dominate the output —
// they cost more than the subject, sender and date combined, and none of it is
// information the caller reads. So a scan hands back short refs instead and
// remembers the mapping here.
//
// The cache lives for the process, which matches how scanning is actually used: list
// a page, pick the interesting rows, fetch those. A ref from before a restart must
// resolve to a clear error rather than a wrong message, because silently fetching or
// moving the wrong email is far worse than asking for a re-scan. Two things make that
// hold:
//
// - Each process picks a random prefix, so a ref like "k7q2-12" names the process
//   that issued it. After a restart the prefix no longer matches and the ref fails.
// - Numbers come from a counter that only goes up, never from the map's size, so
//   numbers are never reused, even if entries are later evicted.
//
// Refs are qualified by the mailbox they were minted against, for the same reason. A
// message ID is only meaningful within its own mailbox, so a ref from a scan of a
// delegated mailbox must not resolve while addressing /me — that would fetch an
// unrelated message, or 404 confusingly. The mailbox is not part of the ref the caller
// sees (it stays short, which is the whole point); it is checked on resolve, and a
// mismatch is reported as such.

const REF_PREFIX_LENGTH = 4
const REF_ALPHABET = "abcdefghijklmnopqrstuvwxyz0123456789"

const newPrefix = (): string =>
  Array.from({ length: REF_PREFIX_LENGTH }, () => REF_ALPHABET[randomInt(REF_ALPHABET.length)]).join("")

const REF_PATTERN = /^([a-z0-9]{4})-(\d+)$/

// Undefined key = the signed-in user's own mailbox, matching MailboxScope.mailbox.
type MailboxKey = string | undefined

type RefEntry = {
  readonly id: string
  readonly mailbox: MailboxKey
}

export type RefResolution =
  | { readonly kind: "id"; readonly id: string }
  | { readonly kind: "unknown"; readonly ref: string }
  | { readonly kind: "wrong-mailbox"; readonly ref: string; readonly mintedFor: MailboxKey }

// The maps would otherwise grow for the life of the process. Past this many refs the least
// recently used one is dropped; it then resolves as unknown, and the caller re-runs the scan.
// The counter never goes down, so a dropped ref is never reissued to a different message.
export const MAX_MESSAGE_REFS = 25_000

const state = { prefix: newPrefix(), counter: 0 }
const refToEntry = new Map<string, RefEntry>()
const keyToRef = new Map<string, string>()

// One namespace per mailbox: the same message ID scanned in two mailboxes is two refs.
// "\u0000" as an escape, not a literal NUL: an invisible control byte in source makes
// the file read as binary to git and unreviewable in a diff. It separates the two
// parts safely because it cannot occur in an address or a Graph message id.
const cacheKey = (mailbox: MailboxKey, id: string): string => `${mailbox ?? ""}\u0000${id}`

// A Map iterates in insertion order, so re-inserting an entry marks it most recently used and
// the first key is always the least recently used.
const touch = (ref: string, entry: RefEntry): void => {
  refToEntry.delete(ref)
  refToEntry.set(ref, entry)
}

const evictOldest = (): void => {
  const oldest = refToEntry.entries().next()
  if (oldest.done) return
  const [ref, entry] = oldest.value
  refToEntry.delete(ref)
  keyToRef.delete(cacheKey(entry.mailbox, entry.id))
}

export const rememberMessageId = (id: string, mailbox?: string): string => {
  const key = cacheKey(mailbox, id)
  const existing = keyToRef.get(key)
  if (existing !== undefined) {
    touch(existing, { id, mailbox })
    return existing
  }

  // Starts at 1 so "0" in output is obviously a bug.
  state.counter += 1
  const ref = `${state.prefix}-${state.counter}`
  refToEntry.set(ref, { id, mailbox })
  keyToRef.set(key, ref)
  if (refToEntry.size > MAX_MESSAGE_REFS) evictOldest()
  return ref
}

export const resolveMessageRef = (ref: string, mailbox?: string): RefResolution => {
  const normalized = ref.trim().toLowerCase()
  const entry = refToEntry.get(normalized)
  if (entry === undefined) return { kind: "unknown", ref: normalized }
  touch(normalized, entry)
  if (entry.mailbox !== mailbox) return { kind: "wrong-mailbox", ref: normalized, mintedFor: entry.mailbox }
  return { kind: "id", id: entry.id }
}

// A caller may pass either a ref from a scan or a full Graph ID. Anything shaped like a
// ref is looked up, and a ref from another process (or a bare number, the old format)
// comes back unknown so the caller reports it. Everything else goes to Graph untouched —
// a raw ID carries its own mailbox context, so there is nothing to check.
export const resolveMessageIdOrRef = (idOrRef: string, mailbox?: string): RefResolution => {
  const trimmed = idOrRef.trim()
  if (/^\d+$/.test(trimmed)) return { kind: "unknown", ref: trimmed }
  if (!REF_PATTERN.test(trimmed.toLowerCase())) return { kind: "id", id: trimmed }

  return resolveMessageRef(trimmed, mailbox)
}

const describeMailbox = (mailbox: MailboxKey): string => mailbox ?? "your own mailbox"

/** Human-readable reason a ref did not resolve, for the tool layer to surface. */
export const describeRefFailure = (
  resolution: Exclude<RefResolution, { kind: "id" }>,
  requested: MailboxKey,
): string =>
  resolution.kind === "unknown"
    ? `Unknown message ref "${resolution.ref}". Refs come from scan_messages and expire after a restart, or once many newer ones push them out — re-run the scan to refresh them.`
    : `Message ref "${resolution.ref}" was scanned in ${describeMailbox(resolution.mintedFor)}, but this call addresses ${describeMailbox(requested)}. Re-scan that mailbox, or pass the mailbox the ref came from.`

export const messageRefCount = (): number => refToEntry.size

// Test seam only — refs are process-scoped in normal use. A new prefix stands in for a
// server restart; tests pass one so two random prefixes can never collide.
export const clearMessageRefs = (prefix: string = newPrefix()): void => {
  refToEntry.clear()
  keyToRef.clear()
  state.prefix = prefix
}
