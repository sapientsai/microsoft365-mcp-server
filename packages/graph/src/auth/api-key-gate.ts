import { resolveUploadTicket } from "@sapientsai/ms-graph-core"
import { getRequestHeader } from "somamcp"

// The shared `authenticate` gate for the httpStream transport AND the protected /upload
// route. A presented bearer authorizes if it equals the configured MCP_API_KEY, or is an
// opaque upload ticket (minted by get_upload_config) that resolves to it. Resolving tickets
// here — rather than in a per-route check — is what lets /upload inherit somamcp's built-in
// `protected` gate. A ticket is functionally equivalent to the key (it resolves to it) and
// short-lived, so accepting it on the transport too is not a widening of trust.
export const authorizesWithApiKey = (bearer: string | undefined, apiKey: string): boolean => {
  if (!bearer) return false
  const resolved = resolveUploadTicket(bearer) ?? bearer
  return resolved === apiKey
}

// http.IncomingMessage.url is path-relative ("/mcp?api_key=…"); a Hono Request carries an
// absolute URL. A dummy base parses both. Never used for anything but reading the query.
const QUERY_BASE = "http://request.invalid"

const queryParam = (request: unknown, name: string): string | undefined => {
  const url = (request as { url?: unknown } | null | undefined)?.url
  if (typeof url !== "string") return undefined
  try {
    return new URL(url, QUERY_BASE).searchParams.get(name) ?? undefined
  } catch {
    return undefined
  }
}

// Same shape as config.ts's helper: a header of "Bearer " or "?api_key=" is absent,
// not a key of "".
const blankToUndefined = (value?: string): string | undefined => {
  const trimmed = value?.trim()
  return trimmed === "" ? undefined : trimmed
}

/**
 * The key a caller presented, from either place the archived
 * `sapientsai/microsoft-mcp-server` accepted one:
 *
 * - `Authorization: Bearer <key>` — preferred, and what a client sets when it can.
 * - `?api_key=<key>` — the fallback for MCP clients that can only be handed a URL.
 *   claude.ai custom connectors are the case that matters: they have no header field,
 *   so the key can only travel in the query string.
 *
 * The monorepo port read the header only, which silently 401'd every connector
 * configured the query way. Header wins when both are present.
 *
 * Caveat inherited from the predecessor, and pinned by test: a query string decodes
 * "+" to a space, so a key containing "+" cannot travel this way and must use the
 * header (or be percent-encoded). Prefer base64url-shaped keys for connector URLs.
 */
export const presentedApiKey = (request: unknown): string | undefined =>
  blankToUndefined(getRequestHeader(request, "authorization")?.replace(/^Bearer\s+/i, "")) ??
  blankToUndefined(queryParam(request, "api_key"))
