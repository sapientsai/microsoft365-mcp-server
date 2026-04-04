import type { TokenCredential } from "@azure/identity"
import { type Either, Left, Right } from "functype/either"
import { None, type Option, Some } from "functype/option"
import jwt from "jsonwebtoken"

import type { AuthConfig, AuthError, AuthMode, AuthStatus } from "../types"
import { createCredential, isClientProvidedToken, testCredential } from "./auth-modes"
import type { TokenInfo } from "./auth-types"
import { GRAPH_DEFAULT_SCOPE } from "./scopes"
import { getContextToken } from "./token-context"

type MutableAuthState = {
  credential: TokenCredential
  config: AuthConfig
  scopes: string[]
}

let authState: Option<MutableAuthState> = None()

const parseJwtScopes = (token: string): ReadonlyArray<string> => {
  try {
    const decoded = jwt.decode(token)
    if (!decoded || typeof decoded !== "object") return []

    if (typeof decoded.scp === "string") {
      return decoded.scp.split(" ").filter((s: string) => s.length > 0)
    }

    if (Array.isArray(decoded.roles)) {
      return decoded.roles as string[]
    }

    return []
  } catch {
    return []
  }
}

export const initializeAuth = async (config: AuthConfig): Promise<Either<AuthError, true>> => {
  const credentialResult = createCredential(config)

  if (credentialResult.isLeft()) {
    return Left(credentialResult.value as AuthError)
  }

  const credential = credentialResult.value as TokenCredential
  const testResult = await testCredential(credential)

  if (testResult.isLeft()) {
    return Left(testResult.value as AuthError)
  }

  authState = Some({ credential, config, scopes: [] as string[] })
  return Right(true as const)
}

export const getAuthState = (): Option<MutableAuthState> => authState

export const getCredential = (): Option<TokenCredential> => authState.map((s) => s.credential)

export const getAuthMode = (): Option<AuthMode> => authState.map((s) => s.config.mode)

export const setAccessToken = (token: string, expiresOn?: Date): Either<AuthError, true> => {
  if (authState.isNone()) {
    return Left({ type: "config" as const, message: "Auth not initialized" })
  }

  const state = authState.value as MutableAuthState
  if (state.config.mode !== "client-token") {
    return Left({ type: "config" as const, message: "set_access_token only supported in client-token mode" })
  }

  if (isClientProvidedToken(state.credential)) {
    state.credential.updateToken(token, expiresOn)
    return Right(true as const)
  }

  return Left({ type: "credential" as const, message: "Credential is not a client-provided token" })
}

export const getAuthStatus = async (): Promise<Either<AuthError, AuthStatus>> => {
  // In OAuth proxy mode, tokens come per-request via AsyncLocalStorage
  const contextToken = getContextToken()
  if (contextToken) {
    const scopes = parseJwtScopes(contextToken)
    const status: AuthStatus = {
      mode: "oauth-proxy",
      authenticated: true,
      scopes,
    }
    return Right(status)
  }

  if (authState.isNone()) {
    return Left({ type: "config" as const, message: "Auth not initialized" })
  }

  const state = authState.value as MutableAuthState
  const { mode } = state.config
  const tokenInfo = await getTokenInfo(state.credential)

  const status: AuthStatus = {
    mode,
    authenticated: !tokenInfo.isExpired,
    scopes: tokenInfo.scopes ?? [],
    expiresAt: tokenInfo.expiresOn?.toISOString(),
  }
  return Right(status)
}

const getTokenInfo = async (credential: TokenCredential): Promise<TokenInfo> => {
  if (isClientProvidedToken(credential)) {
    const isExpired = credential.isExpired()
    const expiresOn = credential.getExpirationTime()

    if (!isExpired) {
      const accessToken = credential.getAccessTokenValue()
      if (accessToken) {
        const scopes = parseJwtScopes(accessToken)
        return { isExpired, expiresOn, scopes }
      }
    }

    return { isExpired, expiresOn }
  }

  try {
    const token = await credential.getToken(GRAPH_DEFAULT_SCOPE)
    if (token?.token) {
      const scopes = parseJwtScopes(token.token)
      return {
        isExpired: false,
        expiresOn: new Date(token.expiresOnTimestamp),
        scopes,
      }
    }
  } catch {
    // Token acquisition failed
  }

  return { isExpired: true }
}

export const getAccessToken = async (): Promise<Either<AuthError, string>> => {
  // Check AsyncLocalStorage for per-request token (OAuth proxy mode)
  const contextToken = getContextToken()
  if (contextToken) {
    return Right(contextToken)
  }

  if (authState.isNone()) {
    return Left({ type: "config" as const, message: "Auth not initialized" })
  }

  const state = authState.value as MutableAuthState
  try {
    const token = await state.credential.getToken(GRAPH_DEFAULT_SCOPE)
    if (!token?.token) {
      return Left({ type: "token" as const, message: "Failed to acquire access token" })
    }
    return Right(token.token)
  } catch (e) {
    return Left({ type: "token" as const, message: `Token acquisition failed: ${String(e)}` })
  }
}

// For testing: reset auth state
export const resetAuth = (): void => {
  authState = None()
}
