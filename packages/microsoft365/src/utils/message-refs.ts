// Graph message IDs are ~152 characters. When a caller is scanning thousands of
// messages to decide which few are worth opening, those IDs dominate the output —
// they cost more than the subject, sender and date combined, and none of it is
// information the caller reads. So a scan hands back short indices instead and
// remembers the mapping here.
//
// The cache lives for the process, which matches how scanning is actually used: list
// a page, pick the interesting rows, fetch those. A ref that has expired (server
// restarted mid-triage) resolves to a clear error rather than a wrong message,
// because silently fetching the wrong email is far worse than asking for a re-scan.
//
// Refs are qualified by the mailbox they were minted against, for the same reason. A
// message ID is only meaningful within its own mailbox, so a ref from a scan of a
// delegated mailbox must not resolve while addressing /me — that would fetch an
// unrelated message, or 404 confusingly. The mailbox is not part of the ref the caller
// sees (it stays a bare number, which is the whole point); it is checked on resolve,
// and a mismatch is reported as such.

// Undefined key = the signed-in user's own mailbox, matching MailboxScope.mailbox.
type MailboxKey = string | undefined

type RefEntry = {
  readonly id: string
  readonly mailbox: MailboxKey
}

export type RefResolution =
  | { readonly kind: "id"; readonly id: string }
  | { readonly kind: "unknown"; readonly ref: number }
  | { readonly kind: "wrong-mailbox"; readonly ref: number; readonly mintedFor: MailboxKey }

const refToEntry = new Map<number, RefEntry>()
const keyToRef = new Map<string, number>()

// Starts at 1 so a ref is never falsy, and so "0" in output is obviously a bug.
const nextRef = () => refToEntry.size + 1

// One namespace per mailbox: the same message ID scanned in two mailboxes is two refs.
// "\u0000" as an escape, not a literal NUL: an invisible control byte in source makes
// the file read as binary to git and unreviewable in a diff. It separates the two
// parts safely because it cannot occur in an address or a Graph message id.
const cacheKey = (mailbox: MailboxKey, id: string): string => `${mailbox ?? ""}\u0000${id}`

export const rememberMessageId = (id: string, mailbox?: string): number => {
  const key = cacheKey(mailbox, id)
  const existing = keyToRef.get(key)
  if (existing !== undefined) return existing

  const ref = nextRef()
  refToEntry.set(ref, { id, mailbox })
  keyToRef.set(key, ref)
  return ref
}

export const resolveMessageRef = (ref: number, mailbox?: string): RefResolution => {
  const entry = refToEntry.get(ref)
  if (entry === undefined) return { kind: "unknown", ref }
  if (entry.mailbox !== mailbox) return { kind: "wrong-mailbox", ref, mintedFor: entry.mailbox }
  return { kind: "id", id: entry.id }
}

// A caller may pass either a short ref from a scan or a full Graph ID. Numeric
// strings are refs; anything else is passed through to Graph untouched — a raw ID
// carries its own mailbox context, so there is nothing to check.
export const resolveMessageIdOrRef = (idOrRef: string, mailbox?: string): RefResolution => {
  const trimmed = idOrRef.trim()
  if (!/^\d+$/.test(trimmed)) return { kind: "id", id: trimmed }

  return resolveMessageRef(Number(trimmed), mailbox)
}

const describeMailbox = (mailbox: MailboxKey): string => mailbox ?? "your own mailbox"

/** Human-readable reason a ref did not resolve, for the tool layer to surface. */
export const describeRefFailure = (
  resolution: Exclude<RefResolution, { kind: "id" }>,
  requested: MailboxKey,
): string =>
  resolution.kind === "unknown"
    ? `Message ref ${resolution.ref} is not known — it may be from before a restart. Re-run scan_messages to get current refs.`
    : `Message ref ${resolution.ref} was scanned in ${describeMailbox(resolution.mintedFor)}, but this call addresses ${describeMailbox(requested)}. Re-scan that mailbox, or pass the mailbox the ref came from.`

export const messageRefCount = (): number => refToEntry.size

// Test seam only — refs are process-scoped in normal use.
export const clearMessageRefs = (): void => {
  refToEntry.clear()
  keyToRef.clear()
}
