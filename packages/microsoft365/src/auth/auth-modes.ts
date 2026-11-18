import type {
  AccessToken as AzureAccessToken,
  DeviceCodeCredentialOptions,
  GetTokenOptions,
  TokenCredential,
} from "@azure/identity"
import {
  AuthenticationRequiredError,
  ClientCertificateCredential,
  ClientSecretCredential,
  DeviceCodeCredential,
  InteractiveBrowserCredential,
  useIdentityPlugin,
} from "@azure/identity"
import { Ref } from "functype"
import { type Either, Left, Right } from "functype/either"
import { Try } from "functype/try"

import type { AuthConfig, AuthError } from "../types"
import { configFingerprint } from "./config-fingerprint"
import { GRAPH_DEFAULT_SCOPE } from "./scopes"
import {
  dropAccessTokensOnConfigChange,
  fileCachePersistencePlugin,
  readAuthenticationRecord,
  writeAuthenticationRecord,
} from "./token-cache"

const ONE_HOUR_MS = 60 * 60 * 1000

const DEFAULT_REDIRECT_URI = "http://localhost:3000"

const TOKEN_CACHE_NAME = "microsoft365-mcp-server"

// useIdentityPlugin mutates module-level state in @azure/identity, so it must run once
// and before any credential is constructed. Ref holds the one-shot flag without a
// mutable binding, matching the functional style enforced across this package.
const cacheRegistered = Ref(false)

const enableTokenCachePersistence = (): boolean => {
  // Opt out with MS365_TOKEN_CACHE=false — e.g. a shared machine, or a CI run where a
  // persisted token would outlive the job.
  if (process.env.MS365_TOKEN_CACHE === "false") return false

  if (!cacheRegistered.get()) {
    useIdentityPlugin(fileCachePersistencePlugin)
    cacheRegistered.set(true)
  }
  return true
}

/** Test seam: reset the one-shot plugin registration. */
export const resetTokenCacheRegistrationForTests = (): void => cacheRegistered.set(false)

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

type PersistenceOptions = Pick<
  DeviceCodeCredentialOptions,
  "tokenCachePersistenceOptions" | "authenticationRecord" | "disableAutomaticAuthentication"
>

// Opt out of the browser entirely: headless hosts, SSH sessions, or a machine whose
// default browser cannot be launched on demand.
const createDeviceCodeCredential = (
  tenantId: string | undefined,
  clientId: string,
  persistence: PersistenceOptions = {},
): DeviceCodeCredential =>
  new DeviceCodeCredential({
    tenantId,
    clientId,
    ...persistence,
    userPromptCallback: (info) => {
      console.error(`\nAuthentication Required:`)
      console.error(`Please visit: ${info.verificationUri}`)
      console.error(`And enter code: ${info.userCode}\n`)
    },
  })

// Persisting the cache is not enough on its own: the credential only consults it once it
// knows which account to look for, and that comes from an AuthenticationRecord. So both
// the browser and device-code credentials are driven through their documented two-step
// flow instead of letting them sign in implicitly:
//
// - Constructed with disableAutomaticAuthentication, getToken tries the cache silently
//   and throws AuthenticationRequiredError rather than prompting when it cannot.
// - authenticate() then signs in interactively and returns the AuthenticationRecord,
//   which is persisted so the next start can go straight to the cache.
//
// A refused silent attempt is logged when a record existed, because otherwise a stale
// record, a scope change or an expired refresh token all look the same: "it asked me to
// sign in again". Identity replaces the underlying reason with a generic message when it
// throws, and only writes the original to its own logger, hence the AZURE_LOG_LEVEL hint.
//
// With disableSignIn the refusal is final. A host that launches the server unattended
// (an agent gateway, a scheduled job) has nobody to read the device code or click
// through the browser, so authenticate() would just block until the host's own startup
// timeout. Failing at once lets the host carry on without these tools and report why.
const signInWhenRequired = (
  credential: InteractiveBrowserCredential | DeviceCodeCredential,
  hadRecord: boolean,
  disableSignIn = false,
): TokenCredential => {
  // Parallel Graph calls can all miss the cache at once; they share one sign-in rather
  // than each prompting separately.
  const pending = Ref<Promise<void> | undefined>(undefined)

  const signIn = async (scopes: string | string[], options: GetTokenOptions | undefined): Promise<void> => {
    const record = await credential.authenticate(scopes, options)
    if (record) writeAuthenticationRecord(record)
  }

  return {
    getToken: async (scopes, options) => {
      try {
        return await credential.getToken(scopes, options)
      } catch (error) {
        if (!(error instanceof AuthenticationRequiredError)) throw error
        if (disableSignIn) {
          throw new Error(
            `[Auth] Sign-in required: ${hadRecord ? "the saved sign-in could not be reused" : "there is no saved sign-in"}, ` +
              `and MS365_DISABLE_SIGN_IN is set, so none was started. Run the server once without it to sign in.${
                // A saved sign-in can fail to refresh for a passing reason, such as the network, and
                // then signing in again won't help. The Azure log says which.
                hadRecord ? " Set AZURE_LOG_LEVEL=info to see why it could not be reused." : ""
              }`,
            { cause: error },
          )
        }
        if (hadRecord) {
          console.error(
            "[Auth] Cached credentials could not be reused, so signing in again. " +
              "Set AZURE_LOG_LEVEL=info to see why.",
          )
        }
        const inFlight = pending.get() ?? signIn(scopes, options).finally(() => pending.set(undefined))
        pending.set(inFlight)
        await inFlight
        try {
          return await credential.getToken(scopes, options)
        } catch (retryError) {
          // Left as is, this surfaces identity's "Automatic authentication has been
          // disabled... call authenticate()", which reads as advice to the user and is
          // wrong: authenticate() has just been called.
          throw new Error(
            "[Auth] Sign-in completed, but a token could still not be obtained for the signed-in account. " +
              `Set AZURE_LOG_LEVEL=info to see why. (${retryError instanceof Error ? retryError.message : String(retryError)})`,
            { cause: retryError },
          )
        }
      }
    },
  }
}

const createInteractiveCredential = (
  config: Extract<AuthConfig, { mode: "interactive" }>,
): Either<AuthError, TokenCredential> => {
  const { tenantId, clientId } = config

  if (!clientId) {
    return Left<AuthError, TokenCredential>({ type: "config", message: "Interactive mode requires MS365_CLIENT_ID" })
  }

  const redirectUri = config.redirectUri ?? DEFAULT_REDIRECT_URI

  if (!enableTokenCachePersistence()) {
    // Without the cache every start needs a sign-in, so refusing one would never work.
    if (config.disableSignIn) {
      return Left<AuthError, TokenCredential>({
        type: "config",
        message: "MS365_DISABLE_SIGN_IN needs the token cache; it cannot be combined with MS365_TOKEN_CACHE=false",
      })
    }
    return config.useDeviceCode
      ? tryCredential(() => createDeviceCodeCredential(tenantId, clientId), "device code")
      : tryCredential(() => new InteractiveBrowserCredential({ tenantId, clientId, redirectUri }), "interactive")
  }

  dropAccessTokensOnConfigChange(configFingerprint(config), TOKEN_CACHE_NAME)

  // Both flows share this: whichever one is used, the resulting token survives a
  // restart, which is the entire point of the interactive mode being usable at all.
  const authenticationRecord = readAuthenticationRecord()
  const persistence: PersistenceOptions = {
    tokenCachePersistenceOptions: { enabled: true, name: TOKEN_CACHE_NAME },
    authenticationRecord,
    disableAutomaticAuthentication: true,
  }
  const hadRecord = authenticationRecord !== undefined
  const disableSignIn = config.disableSignIn ?? false

  if (config.useDeviceCode) {
    return tryCredential(
      () => signInWhenRequired(createDeviceCodeCredential(tenantId, clientId, persistence), hadRecord, disableSignIn),
      "device code",
    )
  }

  return tryCredential(
    () =>
      signInWhenRequired(
        new InteractiveBrowserCredential({ tenantId, clientId, redirectUri, ...persistence }),
        hadRecord,
        disableSignIn,
      ),
    "interactive",
  )
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
