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
// The env var is configuration, not security: it gives a clear error instead of a Graph 403,
// but it is Azure that actually enforces access. What an empty list means depends on who the
// server signs in as:
//
//   delegated      Graph already limits the user to the mailboxes they have been given, so
//                  with no list they reach any of them. Setting the list narrows the agent
//                  further: a user can often open mailboxes in Exchange (a partner's, say)
//                  that they do not want a given agent to touch.
//   app-only       Graph would allow any mailbox in the tenant the ApplicationAccessPolicy
//                  permits, so the list is the server's only limit and an empty one reaches
//                  nothing.
//
// The signed-in user's own address needs no listing: it is the same mailbox as /me, and
// routes there, so a ref or cursor from it matches one issued with mailbox omitted.

import { createHash } from "node:crypto"

import { type Either, Left, Right } from "functype/either"
import { UserError } from "somamcp"

import { getAccessToken, getAuthMode } from "../auth/auth-manager"
import { getContextToken } from "../auth/token-context"
import { parseJwtKind, type TokenKind } from "../auth/token-scopes"
import type { AuthMode, GraphUser, ODataParams } from "../types"

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
 * (comma-separated). Empty means no narrowing when signed in as a user, and no mailbox
 * at all when signed in as an application.
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

// oauth-proxy never runs setupAuth, so getAuthMode() is unset on the hosted server; the env var
// is what chose the mode there, and is read the same way elsewhere (meetings-tools).
const configuredMode = (): string =>
  getAuthMode().fold(
    () => process.env.MS365_AUTH_MODE ?? "interactive",
    (mode): string => mode,
  )

const DELEGATED_MODES: ReadonlySet<string> = new Set<AuthMode>(["interactive", "oauth-proxy"])

const currentToken = async (): Promise<string | undefined> => {
  const contextToken = getContextToken()
  if (contextToken) return contextToken
  const token = await getAccessToken()
  return token.fold(
    (): string | undefined => undefined,
    (value) => value,
  )
}

/**
 * Whether this server acts for a user or as an application. Client-secret and certificate are
 * always app-only. client-token can be handed either kind of token, so it goes by the token's
 * claims, and is "unknown" when they cannot be read: personal-account tokens are opaque, so that
 * case keeps /me, but an empty allowlist then refuses rather than opens. An unrecognised mode
 * is treated as app-only.
 */
export const signInKind = async (): Promise<TokenKind> => {
  const mode = configuredMode()
  if (DELEGATED_MODES.has(mode)) return "delegated"
  if (mode !== "client-token") return "app-only"
  const token = await currentToken()
  return token === undefined ? "unknown" : parseJwtKind(token)
}

const requiresMailbox = () =>
  new UserError(
    "This server signs in as an application, which has no mailbox of its own. " +
      "Pass mailbox with the address of the mailbox to act on.",
  )

const NO_ADDRESSES: ReadonlySet<string> = new Set()

// Every SMTP address that reaches the signed-in user's mailbox: the primary mail, the UPN,
// and each alias. proxyAddresses mixes types (SMTP:, smtp:, SIP:, X500:); only SMTP ones are
// addresses a caller could pass, and the case of the prefix only marks which is primary.
export const ownAddressesOf = (user: Pick<GraphUser, "mail" | "userPrincipalName" | "proxyAddresses">) =>
  new Set(
    [
      user.mail,
      user.userPrincipalName,
      ...(user.proxyAddresses ?? []).filter((entry) => /^smtp:/i.test(entry)).map((entry) => entry.slice(5)),
    ]
      .filter((address): address is string => typeof address === "string")
      .map(normalize)
      .filter((address) => address.length > 0),
  ) as ReadonlySet<string>

// Modes whose credential is one user's for the life of the process. oauth-proxy is handled
// separately below, by the per-request token; app-only modes have no own mailbox.
const OWN_ADDRESS_MODES: ReadonlySet<AuthMode> = new Set(["interactive", "client-token"])

type IdentityClient = {
  readonly request: <T>(
    method: string,
    path: string,
    options?: { odataParams?: ODataParams },
  ) => Promise<Either<unknown, T>>
}

// Keyed by a hash of the token the request actually uses, never by process or session alone.
// In oauth-proxy each request carries its own user's token, so one user's addresses can only
// be found again under that user's token. Outside the proxy, set_access_token can swap a
// client-token deployment to another user, and the old answer must not survive it either.
// Bounded so a long-running proxy does not keep one entry per token it has ever seen; a Map
// iterates in insertion order, so the first key is the oldest.
export const OWN_ADDRESS_CACHE_LIMIT = 256
const ownAddressCache = new Map<string, ReadonlySet<string>>()

export const clearOwnAddressCache = (): void => ownAddressCache.clear()

const remember = (key: string, addresses: ReadonlySet<string>): void => {
  if (ownAddressCache.size >= OWN_ADDRESS_CACHE_LIMIT) {
    const oldest = ownAddressCache.keys().next()
    if (!oldest.done) ownAddressCache.delete(oldest.value)
  }
  ownAddressCache.set(key, addresses)
}

// The token the /me lookup below will be sent with. The Graph client's AuthStrategy is
// getAccessToken, which returns the per-request context token first, so in oauth-proxy both
// this and the lookup use the caller's own token.
const requestToken = async (): Promise<string | undefined> => {
  const contextToken = getContextToken()
  if (contextToken) return contextToken
  const eligible = getAuthMode().fold(
    () => false,
    (mode) => OWN_ADDRESS_MODES.has(mode),
  )
  if (!eligible) return undefined
  const token = await getAccessToken()
  return token.fold(
    (): string | undefined => undefined,
    (value) => value,
  )
}

/**
 * The signed-in user's own addresses, or none when they cannot be known safely. Every failure
 * returns none, which leaves the caller to the allowlist: the shortcut can only widen access
 * when the lookup positively says the address is the caller's own.
 */
export const lookupOwnAddresses = async (client: IdentityClient): Promise<ReadonlySet<string>> => {
  const token = await requestToken()
  if (token === undefined) return NO_ADDRESSES
  const key = createHash("sha256").update(token).digest("hex")
  const cached = ownAddressCache.get(key)
  if (cached) return cached

  const me = await client.request<GraphUser>("GET", "/me", {
    odataParams: { $select: ["mail", "userPrincipalName", "proxyAddresses"] },
  })
  if (me.isLeft()) return NO_ADDRESSES
  const addresses = ownAddressesOf(me.value as GraphUser)
  remember(key, addresses)
  return addresses
}

/**
 * resolveMailboxScope with the deployment's live settings, looking up the signed-in user's
 * own addresses only when a mailbox is actually named in a delegated mode.
 */
export const resolveMailbox = async (
  mailbox: string | undefined,
  client: IdentityClient,
): Promise<Either<UserError, MailboxScope>> => {
  const kind = await signInKind()
  const named = mailbox !== undefined && normalize(mailbox).length > 0
  const own = named && kind !== "app-only" ? await lookupOwnAddresses(client) : NO_ADDRESSES
  return resolveMailboxScope(mailbox, resolveAllowedMailboxes(), kind, own)
}

/**
 * Resolve a caller-supplied mailbox into a scope, refusing anything the deployment has
 * not been configured to reach.
 *
 * Omitting the mailbox yields /me, so every existing delegated caller keeps its behaviour
 * and no configuration is needed to stay on the current path. App-only deployments have
 * no /me, so there omitting it is refused here instead of surfacing as a Graph error.
 * Named mailboxes follow the table at the top of this file.
 */
export const resolveMailboxScope = (
  mailbox: string | undefined,
  allowed: ReadonlySet<string> = resolveAllowedMailboxes(),
  kind: TokenKind = "delegated",
  ownAddresses: ReadonlySet<string> = NO_ADDRESSES,
): Either<UserError, MailboxScope> => {
  const appOnly = kind === "app-only"
  if (mailbox === undefined) return appOnly ? Left(requiresMailbox()) : Right(OWN_MAILBOX)

  // Blank is refused rather than read as "my own mailbox": a caller that meant to name a
  // mailbox and sent an empty value would otherwise act on the wrong one without noticing.
  const normalized = normalize(mailbox)
  if (normalized.length === 0)
    return Left(
      appOnly
        ? requiresMailbox()
        : new UserError("mailbox is empty. Omit mailbox for your own mailbox, or pass an email address."),
    )

  if (!appOnly && ownAddresses.has(normalized)) return Right(OWN_MAILBOX)

  const unrestricted = kind === "delegated" && allowed.size === 0
  if (!unrestricted && !allowed.has(normalized)) {
    // The configured addresses are not listed: on a shared deployment, anyone who
    // mistyped an address would otherwise learn every mailbox the server can reach.
    // Naming the setting is enough for an operator, who can read it.
    return Left(
      new UserError(`Mailbox "${mailbox}" is not permitted by this server's configuration (MS365_ALLOWED_MAILBOXES).`),
    )
  }

  // Encoded because it lands in a URL path segment; Graph accepts the address as-is,
  // but an unencoded "+" in an address would otherwise be read as a space.
  return Right({ mailbox: normalized, prefix: `/users/${encodeURIComponent(normalized)}` })
}
