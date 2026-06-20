import { chmodSync, mkdirSync } from "node:fs"

import { AzureProvider, DiskStore } from "fastmcp/auth"

import { DEFAULT_INTERACTIVE_SCOPES } from "./scopes"

export type OAuthProxyConfig = {
  readonly baseUrl: string
  readonly clientId: string
  readonly clientSecret: string
  readonly tenantId?: string
  readonly scopes?: ReadonlyArray<string>
}

// fastmcp v4 defaults allowedRedirectUriPatterns to [] (rejects all DCR).
// Explicit allow-list:
//   - Claude.ai / Claude.com (browser clients)
//   - IPv4 loopback per RFC 8252 §7.3 for native clients (Claude Code CLI binds
//     http://localhost or http://127.0.0.1 on a random port per run).
// Override via MS365_ALLOWED_REDIRECT_URI_PATTERNS for other deployments.
const DEFAULT_REDIRECT_URI_PATTERNS: ReadonlyArray<string> = [
  "https://claude.ai/*",
  "https://claude.com/*",
  "http://localhost/*",
  "http://localhost:*/*",
  "http://127.0.0.1/*",
  "http://127.0.0.1:*/*",
]

const resolveAllowedRedirectUriPatterns = (): ReadonlyArray<string> => {
  const fromEnv = process.env.MS365_ALLOWED_REDIRECT_URI_PATTERNS
  if (fromEnv) {
    return fromEnv
      .split(",")
      .map((s) => s.trim())
      .filter((s) => s.length > 0)
  }
  return DEFAULT_REDIRECT_URI_PATTERNS
}

// Persisted tokens are encrypted; create the store directory owner-only (0700) so the
// encrypted material is never world-readable (the previous /tmp default left it open).
// Override the location with TOKEN_STORAGE_PATH (recommend a persistent, private dir).
export const resolveTokenStoragePath = (): string => process.env.TOKEN_STORAGE_PATH ?? "/tmp/ms365-tokens"

const createTokenStorage = (): DiskStore => {
  const directory = resolveTokenStoragePath()
  mkdirSync(directory, { recursive: true, mode: 0o700 })
  chmodSync(directory, 0o700) // enforce on a pre-existing (possibly world-readable) dir too
  return new DiskStore({ directory })
}

// JWT signing and token-encryption keys must be independent secrets, NOT reused from
// the OAuth client secret. Prefer dedicated env vars; warn loudly when falling back
// (key reuse is acceptable only for local/dev — set both in production).
export const resolveSigningKey = (clientSecret: string): string => {
  const dedicated = process.env.MS365_JWT_SIGNING_KEY
  if (dedicated) return dedicated
  console.error(
    "[Auth][WARN] MS365_JWT_SIGNING_KEY is not set; falling back to the OAuth client secret for JWT signing. " +
      "Set a dedicated MS365_JWT_SIGNING_KEY in production to avoid key reuse.",
  )
  return clientSecret
}

export const resolveEncryptionKey = (clientSecret: string): string => {
  const dedicated = process.env.MS365_TOKEN_ENCRYPTION_KEY
  if (dedicated) return dedicated
  console.error(
    "[Auth][WARN] MS365_TOKEN_ENCRYPTION_KEY is not set; falling back to the OAuth client secret for token encryption. " +
      "Set a dedicated MS365_TOKEN_ENCRYPTION_KEY in production to avoid key reuse.",
  )
  return clientSecret
}

export const createAzureAuthProvider = (config: OAuthProxyConfig): AzureProvider =>
  new AzureProvider({
    baseUrl: config.baseUrl,
    clientId: config.clientId,
    clientSecret: config.clientSecret,
    tenantId: config.tenantId ?? "common",
    scopes: [...(config.scopes ?? DEFAULT_INTERACTIVE_SCOPES)],
    jwtSigningKey: resolveSigningKey(config.clientSecret),
    tokenStorage: createTokenStorage(),
    encryptionKey: resolveEncryptionKey(config.clientSecret),
    allowedRedirectUriPatterns: [...resolveAllowedRedirectUriPatterns()],
  })
