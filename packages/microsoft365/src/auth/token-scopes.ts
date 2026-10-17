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
