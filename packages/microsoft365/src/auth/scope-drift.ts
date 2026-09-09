// A cached token can outlive the consent it was granted under.
//
// The credential modes request `https://graph.microsoft.com/.default`, which means
// "whatever this app registration has been consented". That string never changes, so
// MSAL's cache lookup always matches: after a permission is added in Azure, the cache
// keeps serving the token minted under the *old* consent until it expires on its own.
// The symptom is a 403 from Graph on a call the app registration now permits, and the
// usual fix is for someone to work out that the token cache needs deleting by hand.
//
// The granted scopes are in the token itself (the `scp` claim, or `roles` for app-only),
// so the drift is detectable: compare what was granted against what this deployment
// needs, and if something required is missing, discard the token and acquire a new one.
//
// Deliberately one-directional. A token carrying MORE than expected is fine — an app
// registration may legitimately be consented for other clients — so only a missing
// required scope counts as drift. The check must never turn "you have extra permissions"
// into a re-authentication loop.

import { UserError } from "fastmcp"
import jwt from "jsonwebtoken"

/** Scopes a deployment needs, derived from what it has been configured to do. */
export type RequiredScopes = ReadonlyArray<string>

// Graph scope names are case-insensitive in practice, and the casing in a token does not
// always match the casing in an app registration.
const normalize = (scope: string): string => scope.trim().toLowerCase()

/**
 * Which required scopes are absent from a granted set.
 *
 * An empty granted set means the claim could not be read at all — an opaque token, or a
 * shape this does not understand. That is reported as no drift rather than as everything
 * missing: refusing to proceed on an unreadable token would break every deployment whose
 * token this cannot parse, to catch a problem that may not exist.
 */
export const missingScopes = (granted: ReadonlyArray<string>, required: RequiredScopes): ReadonlyArray<string> => {
  if (granted.length === 0) return []

  const have = new Set(granted.map(normalize))
  return required.filter((scope) => !have.has(normalize(scope)))
}

/**
 * The scopes this deployment actually needs, given how it is configured.
 *
 * Only conditional requirements belong here. The broad set in DEFAULT_INTERACTIVE_SCOPES
 * is what the server asks for, not what it cannot run without — treating all of it as
 * required would invalidate the cache for any deployment whose tenant declined one
 * optional permission, which is worse than the problem being solved.
 */
// Delegated modes carry per-user scopes in `scp`; app-only modes carry app roles in
// `roles`, and the two vocabularies do not overlap. A .Shared scope is meaningless
// app-only: the application permission Mail.ReadWrite already reaches every mailbox the
// tenant's ApplicationAccessPolicy allows, and there is no "shared" variant to grant.
const isAppOnly = (mode: string): boolean => mode === "client-secret" || mode === "certificate"

export const resolveRequiredScopes = (env: NodeJS.ProcessEnv = process.env): RequiredScopes => {
  const required: string[] = []

  // Addressing another mailbox needs a .Shared scope; the non-shared one does not grant
  // it. This is the case that motivated the check: MS365_ALLOWED_MAILBOXES is set, the
  // app registration has been updated, and the cached token predates the change.
  //
  // App-only is exempt. Requiring .Shared there would report a shortfall that no consent
  // could ever satisfy, and send the operator off to grant a permission that does not
  // apply to the mode they are running — worse than saying nothing, because it looks
  // like a real finding at exactly the moment they are switching modes.
  const mode = env.MS365_AUTH_MODE ?? "interactive"
  if (!isAppOnly(mode) && (env.MS365_ALLOWED_MAILBOXES ?? "").trim().length > 0) {
    required.push(env.MS365_READ_ONLY === "true" ? "Mail.Read.Shared" : "Mail.ReadWrite.Shared")
  }

  return required
}

/**
 * Explains drift in terms of the two things that actually resolve it.
 *
 * Both are needed and in this order, which is the part that is easy to get wrong: granting
 * the permission in Azure changes nothing on its own, because the cached token was minted
 * under the old consent and `.default` keeps matching it until it expires.
 */
export const describeScopeDrift = (missing: ReadonlyArray<string>, cacheDirectory?: string): string =>
  `The signed-in token is missing ${missing.join(", ")}. ` +
  `Add the permission to the app registration and grant admin consent, then delete the token cache${
    cacheDirectory ? ` at ${cacheDirectory}` : ""
  } and sign in again — the cached token predates the change, so granting the permission alone will not take effect. ` +
  `Until then, calls needing it will be refused by Graph.`

export const scopeDriftError = (missing: ReadonlyArray<string>): UserError => new UserError(describeScopeDrift(missing))

/**
 * The scopes a token was granted: `scp` for delegated tokens, `roles` for app-only.
 *
 * Returns an empty list for anything unreadable, which callers treat as "cannot tell"
 * rather than "nothing granted".
 */
export const parseGrantedScopes = (token: string): ReadonlyArray<string> => {
  try {
    const decoded = jwt.decode(token)
    if (decoded === null || typeof decoded !== "object") return []

    if (typeof decoded.scp === "string") return decoded.scp.split(" ").filter((s: string) => s.length > 0)
    if (Array.isArray(decoded.roles)) return decoded.roles as string[]

    return []
  } catch {
    return []
  }
}
