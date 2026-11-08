// Reads the scopes out of an access token, for status reporting only. Nothing here verifies the
// signature: the token came from Entra ID or the client, and Graph checks it on every call.
//
// A Microsoft access token is a JWT, three base64url segments, and the payload names delegated
// scopes in `scp` (space-separated) or app roles in `roles`. Anything else, including the opaque
// tokens personal accounts receive, reads as no scopes rather than an error.
const JWT_SHAPE = /^[A-Za-z0-9_-]+\.([A-Za-z0-9_-]+)\.[A-Za-z0-9_-]*$/

const decodePayload = (token: string): unknown => {
  const payload = JWT_SHAPE.exec(token.trim())?.[1]
  if (payload === undefined) return undefined
  try {
    return JSON.parse(Buffer.from(payload, "base64url").toString("utf8")) as unknown
  } catch {
    return undefined
  }
}

export const parseJwtScopes = (token: string): ReadonlyArray<string> => {
  const claims = decodePayload(token)
  if (claims === null || typeof claims !== "object") return []
  const { scp, roles } = claims as { scp?: unknown; roles?: unknown }

  if (typeof scp === "string") return scp.split(" ").filter((s) => s.length > 0)
  if (Array.isArray(roles)) return roles.filter((r): r is string => typeof r === "string")
  return []
}

export type TokenKind = "delegated" | "app-only" | "unknown"

/**
 * Whether a token acts for a user (`scp`) or as an application (`roles` with no `scp`). Anything
 * that cannot be read, including opaque tokens, is "unknown" and left to the caller to treat
 * conservatively.
 */
export const parseJwtKind = (token: string): TokenKind => {
  const claims = decodePayload(token)
  if (claims === null || typeof claims !== "object") return "unknown"
  const { scp, roles } = claims as { scp?: unknown; roles?: unknown }

  if (typeof scp === "string" && scp.trim().length > 0) return "delegated"
  if (Array.isArray(roles)) return "app-only"
  return "unknown"
}

/**
 * The tenant the token was issued in (`tid`), for telling a resource in another organization apart
 * from one the caller simply can't reach. Undefined when the token can't be read.
 */
export const parseJwtTenant = (token: string): string | undefined => {
  const claims = decodePayload(token)
  if (claims === null || typeof claims !== "object") return undefined
  const { tid } = claims as { tid?: unknown }
  return typeof tid === "string" && tid.length > 0 ? tid.toLowerCase() : undefined
}
