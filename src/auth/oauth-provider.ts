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

export const createAzureAuthProvider = (config: OAuthProxyConfig): AzureProvider =>
  new AzureProvider({
    baseUrl: config.baseUrl,
    clientId: config.clientId,
    clientSecret: config.clientSecret,
    tenantId: config.tenantId ?? "common",
    scopes: [...(config.scopes ?? DEFAULT_INTERACTIVE_SCOPES)],
    tokenStorage,
  })
