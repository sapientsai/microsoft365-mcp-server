import { randomBytes } from "node:crypto"

// Opaque, short-lived upload tickets. The real credential (a delegated Microsoft Graph token, or
// the app-only server's API key) is held server-side and mapped to a random ticket id; only the
// id is handed back to the client, and therefore to the transcript. An out-of-band `curl` uses
// it to authenticate to the /upload relay.
//
// A ticket lands in a model's context and in any transcript or log of it, so it is kept as narrow
// as the job: it is good for the one upload it was issued for (a drive path, or a draft in a
// mailbox), on the relay that does that upload, and only once. Redeeming it removes it, whether
// it matched or not. The relay already splits a large file into chunks inside one request, so a
// legitimate upload never needs a second use; after a failed curl, the caller asks for a new one.

const DEFAULT_TTL_MS = 10 * 60 * 1000 // 10 minutes — long enough for a large chunked upload
const TICKET_PREFIX = "upl_"

/** The one upload a ticket is good for. */
export type UploadTarget =
  | { readonly relay: "drive"; readonly path: string }
  | { readonly relay: "attachment"; readonly messageId: string; readonly mailbox?: string }

type Ticket = { readonly token: string; readonly target: UploadTarget; readonly expiresAt: number }

const store = new Map<string, Ticket>()

const sweep = (now: number): void => {
  for (const [id, ticket] of store) {
    if (ticket.expiresAt <= now) store.delete(id)
  }
}

const sameTarget = (a: UploadTarget, b: UploadTarget): boolean =>
  a.relay === "drive" && b.relay === "drive"
    ? a.path === b.path
    : a.relay === "attachment" && b.relay === "attachment"
      ? a.messageId === b.messageId && (a.mailbox ?? "") === (b.mailbox ?? "")
      : false

export const isUploadTicket = (value: string): boolean => value.startsWith(TICKET_PREFIX)

export const mintUploadTicket = (
  token: string,
  target: UploadTarget,
  ttlMs: number = DEFAULT_TTL_MS,
  now: number = Date.now(),
): string => {
  sweep(now)
  const id = `${TICKET_PREFIX}${randomBytes(24).toString("base64url")}`
  store.set(id, { token, target, expiresAt: now + ttlMs })
  return id
}

/**
 * The credential behind a ticket, if the ticket exists, hasn't expired, and was issued for this
 * upload. The ticket is spent either way: a second use, or a use for another upload, is refused.
 */
export const redeemUploadTicket = (id: string, target: UploadTarget, now: number = Date.now()): string | undefined => {
  const ticket = store.get(id)
  if (!ticket) return undefined
  store.delete(id)
  if (ticket.expiresAt <= now) return undefined
  return sameTarget(ticket.target, target) ? ticket.token : undefined
}
