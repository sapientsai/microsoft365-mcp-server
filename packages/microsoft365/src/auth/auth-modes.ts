import type { AccessToken as AzureAccessToken, TokenCredential } from "@azure/identity"
import {
  ClientCertificateCredential,
  ClientSecretCredential,
  DeviceCodeCredential,
  InteractiveBrowserCredential,
} from "@azure/identity"
import { type Either, Left, Right } from "functype/either"
import { Try } from "functype/try"

import type { AuthConfig, AuthError } from "../types"
import { GRAPH_DEFAULT_SCOPE } from "./scopes"

const ONE_HOUR_MS = 60 * 60 * 1000

const DEFAULT_REDIRECT_URI = "http://localhost:3000"

const tryCredential = (fn: () => TokenCredential, label: string): Either<AuthError, TokenCredential> =>
  Try(fn).fold(
    (e): Either<AuthError, TokenCredential> =>
      Left<AuthError, TokenCredential>({
        type: "credential",
        message: `Failed to create ${label} credential: ${String(e)}`,
      }),
    (cred): Either<AuthError, TokenCredential> => Right(cred),
  )

export const createCredential = (config: AuthConfig): Either<AuthError, TokenCredential> => {
  switch (config.mode) {
    case "interactive":
      return createInteractiveCredential(config)
    case "certificate":
      return createCertificateCredential(config)
    case "client-secret":
      return createClientSecretCredential(config)
    case "client-token":
      return createClientProvidedTokenCredential(config)
    case "oauth-proxy":
      return Left<AuthError, TokenCredential>({
        type: "config",
        message: "OAuth proxy mode uses AzureProvider, not credential-based auth",
      })
  }
}

const createDeviceCodeCredential = (tenantId: string | undefined, clientId: string): TokenCredential =>
  new DeviceCodeCredential({
    tenantId,
    clientId,
    userPromptCallback: (info) => {
      console.error(`\nAuthentication Required:`)
      console.error(`Please visit: ${info.verificationUri}`)
      console.error(`And enter code: ${info.userCode}\n`)
    },
  }) as TokenCredential

// InteractiveBrowserCredential launches the browser lazily, inside getToken() — so a
// launch failure surfaces there, never from the constructor. Wrap it so that failure
// falls back to device code instead of aborting authentication.
//
// The common macOS case: the underlying `open` call fails when the default browser is
// already running and refuses a second instance (Arc reports "Arc is already open.
// Only one instance of Arc can be opened at a time."), leaving the user unable to
// authenticate without quitting their browser first.
export const withDeviceCodeFallback = (
  browser: TokenCredential,
  tenantId: string | undefined,
  clientId: string,
  makeFallback: (t: string | undefined, c: string) => TokenCredential = createDeviceCodeCredential,
): TokenCredential => ({
  getToken: async (scopes, options) => {
    try {
      return await browser.getToken(scopes, options)
    } catch (error) {
      if (!isBrowserLaunchFailure(error)) throw error
      console.error(`\nCould not open a browser (${String(error)}).`)
      console.error(`Falling back to device code authentication.`)
      return makeFallback(tenantId, clientId).getToken(scopes, options)
    }
  },
})

// Only fall back for failures to *launch* the browser. A declined consent, a bad
// client ID, or a network error must surface as itself — retrying those under device
// code would just hide the real problem behind a second prompt.
const BROWSER_LAUNCH_FAILURE_PATTERNS: ReadonlyArray<string> = [
  "already open",
  "only one instance",
  "unable to open",
  "failed to open",
  "could not open",
  "no such file or directory",
  "spawn",
  "enoent",
]

export const isBrowserLaunchFailure = (error: unknown): boolean => {
  const message = (error instanceof Error ? error.message : String(error)).toLowerCase()
  return BROWSER_LAUNCH_FAILURE_PATTERNS.some((pattern) => message.includes(pattern))
}

const createInteractiveCredential = (
  config: Extract<AuthConfig, { mode: "interactive" }>,
): Either<AuthError, TokenCredential> => {
  const { tenantId, clientId } = config

  if (!clientId) {
    return Left<AuthError, TokenCredential>({ type: "config", message: "Interactive mode requires MS365_CLIENT_ID" })
  }

  // Opt out of the browser entirely: headless hosts, SSH sessions, or a machine whose
  // default browser cannot be launched on demand.
  if (config.useDeviceCode) {
    return tryCredential(() => createDeviceCodeCredential(tenantId, clientId), "device code")
  }

  return tryCredential(() => {
    const browser = new InteractiveBrowserCredential({
      tenantId,
      clientId,
      redirectUri: config.redirectUri ?? DEFAULT_REDIRECT_URI,
    }) as TokenCredential
    return withDeviceCodeFallback(browser, tenantId, clientId)
  }, "interactive")
}

const createCertificateCredential = (
  config: Extract<AuthConfig, { mode: "certificate" }>,
): Either<AuthError, TokenCredential> => {
  if (!config.tenantId || !config.clientId || !config.certPath) {
    return Left<AuthError, TokenCredential>({
      type: "config",
      message: "Certificate mode requires MS365_TENANT_ID, MS365_CLIENT_ID, and MS365_CERT_PATH",
    })
  }

  return tryCredential(
    () =>
      new ClientCertificateCredential(config.tenantId, config.clientId, {
        certificatePath: config.certPath,
        certificatePassword: config.certPassword,
      }) as TokenCredential,
    "certificate",
  )
}

const createClientSecretCredential = (
  config: Extract<AuthConfig, { mode: "client-secret" }>,
): Either<AuthError, TokenCredential> => {
  if (!config.tenantId || !config.clientId || !config.clientSecret) {
    return Left<AuthError, TokenCredential>({
      type: "config",
      message: "Client secret mode requires MS365_TENANT_ID, MS365_CLIENT_ID, and MS365_CLIENT_SECRET",
    })
  }

  return tryCredential(
    () => new ClientSecretCredential(config.tenantId, config.clientId, config.clientSecret) as TokenCredential,
    "client-secret",
  )
}

export class ClientProvidedTokenCredential implements TokenCredential {
  private _accessToken: string | undefined
  private _expiresOn: Date

  constructor(accessToken?: string, expiresOn?: Date) {
    this._accessToken = accessToken
    this._expiresOn = expiresOn ?? (accessToken ? new Date(Date.now() + ONE_HOUR_MS) : new Date(0))
  }

  // eslint-disable-next-line @typescript-eslint/require-await -- TokenCredential interface requires async
  async getToken(_scopes: string | string[]): Promise<AzureAccessToken | null> {
    if (!this._accessToken || this._expiresOn <= new Date()) {
      return null
    }
    return { token: this._accessToken, expiresOnTimestamp: this._expiresOn.getTime() }
  }

  updateToken(token: string, expiresOn?: Date): void {
    this._accessToken = token
    this._expiresOn = expiresOn ?? new Date(Date.now() + ONE_HOUR_MS)
  }

  isExpired(): boolean {
    return !this._accessToken || this._expiresOn <= new Date()
  }

  getExpirationTime(): Date {
    return this._expiresOn
  }

  getAccessTokenValue(): string | undefined {
    return this._accessToken
  }
}

const createClientProvidedTokenCredential = (
  config: Extract<AuthConfig, { mode: "client-token" }>,
): Either<AuthError, TokenCredential> =>
  Right(new ClientProvidedTokenCredential(config.accessToken, config.expiresOn) as TokenCredential)

export const isClientProvidedToken = (credential: TokenCredential): credential is ClientProvidedTokenCredential =>
  credential instanceof ClientProvidedTokenCredential

export const testCredential = async (credential: TokenCredential): Promise<Either<AuthError, true>> => {
  // Skip test for client-provided token without initial token
  if (isClientProvidedToken(credential) && credential.isExpired()) {
    return Right(true as const)
  }

  try {
    const token = await credential.getToken(GRAPH_DEFAULT_SCOPE)
    if (!token) {
      return Left<AuthError, true>({ type: "token", message: "Failed to acquire token during credential test" })
    }
    return Right(true as const)
  } catch (e) {
    return Left<AuthError, true>({ type: "token", message: `Authentication test failed: ${String(e)}` })
  }
}
