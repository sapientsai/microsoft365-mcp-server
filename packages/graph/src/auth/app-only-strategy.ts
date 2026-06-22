import type { AuthError, AuthStrategy } from "@sapientsai/ms-graph-core"
import { type Either, Left, Right } from "functype/either"
import { Try } from "functype/try"

import type { AppOnlyConfig } from "../config"

const TOKEN_REFRESH_BUFFER_MS = 5 * 60 * 1000 // refresh 5 min before expiry

type TokenResponse = { access_token: string; expires_in: number }
type CachedToken = { accessToken: string; expiresAt: number }

const authError = (type: AuthError["type"], message: string): AuthError => ({ type, message })

// App-only (client_credentials) AuthStrategy adapter — implements core's AuthStrategy so
// packages/graph plugs into the same shared graph plumbing as the delegated server.
// Ported from microsoft-mcp-server/src/auth/token-manager.ts onto the core interface.
export const createAppOnlyAuthStrategy = (
  config: AppOnlyConfig,
  fetchImpl: typeof fetch = fetch,
  now: () => number = Date.now,
): AuthStrategy => {
  const cache: { token: CachedToken | null } = { token: null }

  const isValid = (t: CachedToken): boolean => now() < t.expiresAt - TOKEN_REFRESH_BUFFER_MS

  const fetchToken = async (): Promise<Either<AuthError, CachedToken>> => {
    const tokenUrl = `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`
    const body = new URLSearchParams({
      client_id: config.clientId,
      client_secret: config.clientSecret,
      scope: config.appScopes.join(" "),
      grant_type: "client_credentials",
    })

    const response = await fetchImpl(tokenUrl, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: body.toString(),
    }).catch((error: unknown) => error as Error)

    if (response instanceof Error) {
      return Left(authError("credential", `Token request failed: ${response.message}`))
    }
    if (!response.ok) {
      const text = await response.text()
      const message = Try(() => JSON.parse(text) as { error_description?: string; error?: string }).fold(
        () => `Token request failed: ${response.status} - ${text}`,
        (json) => json.error_description ?? json.error ?? `Token request failed: ${response.status}`,
      )
      return Left(authError("credential", message))
    }

    const data = (await response.json()) as TokenResponse
    return Right({ accessToken: data.access_token, expiresAt: now() + data.expires_in * 1000 })
  }

  const getAccessToken = async (): Promise<Either<AuthError, string>> => {
    if (cache.token && isValid(cache.token)) return Right(cache.token.accessToken)
    const result = await fetchToken()
    return result.map((t) => {
      cache.token = t
      return t.accessToken
    })
  }

  return { getAccessToken }
}
