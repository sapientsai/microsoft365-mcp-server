import { AzureProvider } from "fastmcp/auth"

import { DEFAULT_INTERACTIVE_SCOPES } from "./scopes"

export type OAuthProxyConfig = {
  readonly baseUrl: string
  readonly clientId: string
  readonly clientSecret: string
  readonly tenantId?: string
  readonly scopes?: ReadonlyArray<string>
}

export const createAzureAuthProvider = (config: OAuthProxyConfig): AzureProvider =>
  new AzureProvider({
    baseUrl: config.baseUrl,
    clientId: config.clientId,
    clientSecret: config.clientSecret,
    tenantId: config.tenantId ?? "common",
    scopes: [...(config.scopes ?? DEFAULT_INTERACTIVE_SCOPES)],
  })
