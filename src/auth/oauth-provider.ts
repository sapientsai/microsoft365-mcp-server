import { AzureProvider, DiskStore } from "fastmcp/auth"

import { DEFAULT_INTERACTIVE_SCOPES } from "./scopes"

export type OAuthProxyConfig = {
  readonly baseUrl: string
  readonly clientId: string
  readonly clientSecret: string
  readonly tenantId?: string
  readonly scopes?: ReadonlyArray<string>
}

const tokenStorage = new DiskStore({
  directory: process.env.TOKEN_STORAGE_PATH ?? "/tmp/ms365-tokens",
})

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

// EncryptedTokenStorage uses an ephemeral random key when encryptionKey is unset,
// which makes persisted tokens unreadable after restart. Bind it to the client
// secret so the key is stable across restarts; rotating the secret invalidates
// stored tokens (intended).
const resolveEncryptionKey = (clientSecret: string): string =>
  process.env.MS365_TOKEN_ENCRYPTION_KEY ?? clientSecret

export const createAzureAuthProvider = (config: OAuthProxyConfig): AzureProvider =>
  new AzureProvider({
    baseUrl: config.baseUrl,
    clientId: config.clientId,
    clientSecret: config.clientSecret,
    tenantId: config.tenantId ?? "common",
    scopes: [...(config.scopes ?? DEFAULT_INTERACTIVE_SCOPES)],
    jwtSigningKey: config.clientSecret,
    tokenStorage,
    encryptionKey: resolveEncryptionKey(config.clientSecret),
    allowedRedirectUriPatterns: [...resolveAllowedRedirectUriPatterns()],
  })
