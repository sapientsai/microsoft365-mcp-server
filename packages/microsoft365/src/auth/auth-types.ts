import type { TokenCredential } from "@azure/identity"

import type { AuthConfig, AuthMode, AuthStatus } from "../types"

export type AuthState = {
  readonly credential: TokenCredential
  readonly config: AuthConfig
  readonly scopes: ReadonlyArray<string>
}

export type TokenInfo = {
  readonly isExpired: boolean
  readonly expiresOn?: Date
  readonly scopes?: ReadonlyArray<string>
}

export type { AuthConfig, AuthMode, AuthStatus, TokenCredential }
