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
// Explicit allow-list for Claude.ai clients; override via env for others.
const DEFAULT_REDIRECT_URI_PATTERNS: ReadonlyArray<string> = ["https://claude.ai/*", "https://claude.com/*"]

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

export const createAzureAuthProvider = (config: OAuthProxyConfig): AzureProvider =>
  new AzureProvider({
    baseUrl: config.baseUrl,
    clientId: config.clientId,
    clientSecret: config.clientSecret,
    tenantId: config.tenantId ?? "common",
    scopes: [...(config.scopes ?? DEFAULT_INTERACTIVE_SCOPES)],
    jwtSigningKey: config.clientSecret,
    tokenStorage,
    allowedRedirectUriPatterns: [...resolveAllowedRedirectUriPatterns()],
  })
