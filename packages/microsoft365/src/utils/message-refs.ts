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

const REF_PREFIX_LENGTH = 4
const REF_ALPHABET = "abcdefghijklmnopqrstuvwxyz0123456789"

const newPrefix = (): string =>
  Array.from({ length: REF_PREFIX_LENGTH }, () => REF_ALPHABET[randomInt(REF_ALPHABET.length)]).join("")

const REF_PATTERN = /^([a-z0-9]{4})-(\d+)$/

const state = { prefix: newPrefix(), counter: 0 }
const refToId = new Map<string, string>()
const idToRef = new Map<string, string>()

export const rememberMessageId = (id: string): string => {
  const existing = idToRef.get(id)
  if (existing !== undefined) return existing

  // Starts at 1 so "0" in output is obviously a bug.
  state.counter += 1
  const ref = `${state.prefix}-${state.counter}`
  refToId.set(ref, id)
  idToRef.set(id, ref)
  return ref
}

export const resolveMessageRef = (ref: string): string | undefined => refToId.get(ref.trim().toLowerCase())

// A caller may pass either a ref from a scan or a full Graph ID. Anything shaped like a
// ref is looked up, and a ref from another process (or a bare number, the old format)
// comes back undefined so the caller reports it. Everything else goes to Graph untouched.
export const resolveMessageIdOrRef = (idOrRef: string): string | undefined => {
  const trimmed = idOrRef.trim()
  if (/^\d+$/.test(trimmed)) return undefined
  if (!REF_PATTERN.test(trimmed.toLowerCase())) return trimmed

  return resolveMessageRef(trimmed)
}

export const messageRefCount = (): number => refToId.size

// Test seam only — refs are process-scoped in normal use. A new prefix stands in for a
// server restart; tests pass one so two random prefixes can never collide.
export const clearMessageRefs = (prefix: string = newPrefix()): void => {
  refToId.clear()
  idToRef.clear()
  state.prefix = prefix
}
