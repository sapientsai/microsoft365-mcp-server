import { isUploadTicket, redeemUploadTicket, type UploadTarget } from "@sapientsai/ms-graph-core"

import { getAccessToken } from "../auth"

export type UploadAuthResult = { token?: string; error?: string; status?: number }

// Resolves the credential the /upload relay uses to call Microsoft Graph.
//
// Auth ordering:
//   1. Opaque upload ticket → the Graph token held server-side (raw JWT never travels
//      in the tool transcript). Good once, and only for the upload it was issued for.
//   2. OAuth-proxy mode → the per-request delegated bearer (required).
//   3. Shared-secret mode → caller must present MS365_UPLOAD_TOKEN; the server then
//      uses its own credential to talk to Graph. Only callers that hold the secret out of band
//      use this: the upload config tools hand out tickets and never print the secret.
//
// SECURITY: in non-OAuth mode the endpoint is write-capable using the server's own
// credentials, so it MUST NOT serve a request that presented no valid caller auth.
// If MS365_UPLOAD_TOKEN is unset there is no way to authenticate the caller — refuse
// (503) rather than falling through and uploading with server credentials.
export const resolveUploadAccessToken = async (
  oauthMode: boolean,
  bearer: string | undefined,
  target: UploadTarget,
): Promise<UploadAuthResult> => {
  // A ticket is good once, for the upload it was issued for. Presenting one spends it.
  // A shared secret that happens to start like a ticket is still the secret.
  if (bearer && isUploadTicket(bearer) && bearer !== process.env.MS365_UPLOAD_TOKEN) {
    const ticketed = redeemUploadTicket(bearer, target)
    if (ticketed) return { token: ticketed }
    return {
      error: "Upload ticket expired, already used, or issued for another upload. Get a new upload config.",
      status: 401,
    }
  }

  if (oauthMode) {
    if (!bearer) return { error: "Missing Bearer token", status: 401 }
    return { token: bearer }
  }

  const sharedSecret = process.env.MS365_UPLOAD_TOKEN
  if (!sharedSecret) {
    return {
      error:
        "Upload endpoint is not configured for authentication. Set MS365_UPLOAD_TOKEN (or run in oauth-proxy mode).",
      status: 503,
    }
  }
  if (bearer !== sharedSecret) {
    return { error: "Invalid upload token", status: 401 }
  }

  const result = await getAccessToken()
  if (result.isLeft()) {
    return { error: (result.value as { message: string }).message, status: 401 }
  }
  return { token: result.value as string }
}
