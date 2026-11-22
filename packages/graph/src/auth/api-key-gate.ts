import { isUploadTicket, redeemUploadTicket, type UploadTarget } from "@sapientsai/ms-graph-core"
import { getRequestHeader } from "somamcp"

// A web Request carries an absolute URL; a dummy base also parses a path-relative one
// ("/mcp?api_key=…"). Never used for anything but reading the query.
const QUERY_BASE = "http://request.invalid"

// The shared `authenticate` gate for the httpStream transport AND the protected /upload
// route. A presented bearer authorizes if it equals the configured MCP_API_KEY, or is an
// opaque upload ticket (minted by get_upload_config) redeemed for the upload it was issued for.
// Resolving tickets here — rather than in a per-route check — is what lets /upload inherit
// somamcp's built-in `protected` gate.
//
// A ticket lands in the model's context and any transcript of it, so it is not the key: it is
// good once, on /upload only, for the one path it names. On /mcp and every other route a ticket
// is refused, so a leaked one can't open the MCP endpoint.
export const authorizesWithApiKey = (
  bearer: string | undefined,
  apiKey: string,
  uploadTarget?: UploadTarget,
): boolean => {
  if (!bearer) return false
  if (isUploadTicket(bearer)) {
    return uploadTarget !== undefined && redeemUploadTicket(bearer, uploadTarget) === apiKey
  }
  return bearer === apiKey
}

/** The upload a request to the /upload relay is for, or undefined for any other route. */
export const uploadTargetOf = (request: unknown): UploadTarget | undefined => {
  const url = (request as { url?: unknown } | null | undefined)?.url
  if (typeof url !== "string") return undefined
  try {
    const parsed = new URL(url, QUERY_BASE)
    const path = parsed.searchParams.get("path")
    return parsed.pathname === "/upload" && path !== null ? { relay: "drive", path } : undefined
  } catch {
    return undefined
  }
}

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
