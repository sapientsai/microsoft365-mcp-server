import type { Either } from "functype/either"

import type { AuthError } from "./types"

// The seam between core's Graph plumbing and a server's auth implementation.
//
// Phase 2b will invert graph-client to depend on this interface instead of importing
// a server's auth module directly, so both the delegated (microsoft365) and app-only
// (graph) servers can provide their own adapter:
//   - delegated: wraps interactive/cert/secret/client-token/oauth-proxy + multi-account
//   - app-only:  client_credentials acquisition + refresh
//
// Defined here now so the contract is stable before the graph-client extraction.
export type AuthStrategy = {
  readonly getAccessToken: () => Promise<Either<AuthError, string>>
}
