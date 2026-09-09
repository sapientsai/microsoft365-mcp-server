// Which mailbox a mail request addresses.
//
// Every mail path is either rooted at /me (the signed-in user's own mailbox) or at
// /users/{address} (someone else's, reached by delegation or by an app-only identity).
// Callers pass an address; this module turns it into a path prefix and decides whether
// the deployment is allowed to touch that mailbox at all.
//
// Two distinct boundaries are at work, and they are not redundant:
//
//   Azure          an ApplicationAccessPolicy (app-only) or a delegated grant caps what
//                  the credential could ever reach. This is the security boundary.
//   MS365_ALLOWED_MAILBOXES
//                  narrows one deployment further, within that cap. This is what lets a
//                  single app registration serve several agents with different reach —
//                  a triage bot limited to one mailbox, a household assistant given two —
//                  without minting a registration per agent.
//
// The env var is configuration, not security: it fails closed and gives a clear error
// instead of a Graph 403, but it is Azure that actually enforces access.

import { UserError } from "fastmcp"
import { type Either, Left, Right } from "functype/either"

/** Graph path prefix for a mailbox: "/me" for the signed-in user, else "/users/{address}". */
export type MailboxScope = {
  /** Undefined means the signed-in user's own mailbox. */
  readonly mailbox?: string
  readonly prefix: string
}

export const OWN_MAILBOX: MailboxScope = { prefix: "/me" }

// Addresses are compared case-insensitively — Exchange treats them that way, and an
// allowlist that rejected "Bel@..." against "bel@..." would be a confusing failure.
const normalize = (address: string): string => address.trim().toLowerCase()

/**
 * The mailboxes this deployment may address, from MS365_ALLOWED_MAILBOXES
 * (comma-separated). Unset or empty means no other mailbox is reachable: the server
 * can still work, but only against /me.
 */
export const resolveAllowedMailboxes = (
  raw: string | undefined = process.env.MS365_ALLOWED_MAILBOXES,
): ReadonlySet<string> =>
  new Set(
    (raw ?? "")
      .split(",")
      .map(normalize)
      .filter((entry) => entry.length > 0),
  )

/**
 * Resolve a caller-supplied mailbox into a scope, refusing anything the deployment has
 * not been configured to reach.
 *
 * Omitting the mailbox always yields /me, so every existing caller keeps its behaviour
 * and no configuration is needed to stay on the current path.
 */
export const resolveMailboxScope = (
  mailbox: string | undefined,
  allowed: ReadonlySet<string> = resolveAllowedMailboxes(),
): Either<UserError, MailboxScope> => {
  if (mailbox === undefined) return Right(OWN_MAILBOX)

  const normalized = normalize(mailbox)
  if (normalized.length === 0) return Right(OWN_MAILBOX)

  if (!allowed.has(normalized)) {
    // Name the configured set: the usual cause is a typo or an agent guessing an
    // address, and "which mailboxes may I use" is the question that unblocks it.
    const configured =
      allowed.size === 0
        ? "no mailboxes are configured (set MS365_ALLOWED_MAILBOXES to allow one)"
        : `allowed: ${[...allowed].sort().join(", ")}`
    return Left(new UserError(`Mailbox "${mailbox}" is not permitted by this server's configuration — ${configured}.`))
  }

  // Encoded because it lands in a URL path segment; Graph accepts the address as-is,
  // but an unencoded "+" in an address would otherwise be read as a space.
  return Right({ mailbox: normalized, prefix: `/users/${encodeURIComponent(normalized)}` })
}
