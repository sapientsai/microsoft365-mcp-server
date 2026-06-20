import { randomBytes } from "node:crypto"

// Opaque, short-lived upload tickets. The delegated Microsoft Graph token is held
// server-side and mapped to a random ticket id; only the ticket id is handed back to
// the client (and therefore the LLM transcript). This keeps the raw Graph JWT off the
// transcript / log surface while still letting an out-of-band `curl` authenticate to
// the /upload relay.

const DEFAULT_TTL_MS = 10 * 60 * 1000 // 10 minutes — long enough for a large chunked upload
const TICKET_PREFIX = "upl_"

type Ticket = { token: string; expiresAt: number }

const store = new Map<string, Ticket>()

const sweep = (now: number): void => {
  for (const [id, ticket] of store) {
    if (ticket.expiresAt <= now) store.delete(id)
  }
}

export const isUploadTicket = (value: string): boolean => value.startsWith(TICKET_PREFIX)

export const mintUploadTicket = (token: string, ttlMs: number = DEFAULT_TTL_MS, now: number = Date.now()): string => {
  sweep(now)
  const id = `${TICKET_PREFIX}${randomBytes(24).toString("base64url")}`
  store.set(id, { token, expiresAt: now + ttlMs })
  return id
}

export const resolveUploadTicket = (id: string, now: number = Date.now()): string | undefined => {
  const ticket = store.get(id)
  if (!ticket) return undefined
  if (ticket.expiresAt <= now) {
    store.delete(id)
    return undefined
  }
  return ticket.token
}
