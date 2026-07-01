import { resolveUploadTicket } from "@sapientsai/ms-graph-core"

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
