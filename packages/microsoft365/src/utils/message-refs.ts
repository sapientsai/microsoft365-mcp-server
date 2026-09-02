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

const refToId = new Map<number, string>()
const idToRef = new Map<string, number>()

// Starts at 1 so a ref is never falsy, and so "0" in output is obviously a bug.
const nextRef = () => refToId.size + 1

export const rememberMessageId = (id: string): number => {
  const existing = idToRef.get(id)
  if (existing !== undefined) return existing

  const ref = nextRef()
  refToId.set(ref, id)
  idToRef.set(id, ref)
  return ref
}

export const resolveMessageRef = (ref: number): string | undefined => refToId.get(ref)

// A caller may pass either a short ref from a scan or a full Graph ID. Numeric
// strings are refs; anything else is passed through to Graph untouched.
export const resolveMessageIdOrRef = (idOrRef: string): string | undefined => {
  const trimmed = idOrRef.trim()
  if (!/^\d+$/.test(trimmed)) return trimmed

  return resolveMessageRef(Number(trimmed))
}

export const messageRefCount = (): number => refToId.size

// Test seam only — refs are process-scoped in normal use.
export const clearMessageRefs = (): void => {
  refToId.clear()
  idToRef.clear()
}
