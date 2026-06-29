import { type Either, Left, Right } from "functype/either"

import type { AiSearchConfig } from "./search/ai-search-client"
import type { SharePointSearchConfig } from "./tools/sharepoint-search"

export type AppOnlyConfig = {
  readonly tenantId: string
  readonly clientId: string
  readonly clientSecret: string
  readonly appScopes: ReadonlyArray<string>
}

export type ServerRuntimeConfig = {
  readonly auth: AppOnlyConfig
  readonly apiKey?: string
  readonly transport: "stdio" | "httpStream"
  readonly port: number
  readonly host: string
  readonly publicBaseUrl: string
  readonly aiSearch?: AiSearchConfig
  readonly sharePointSearch: SharePointSearchConfig
}

// App-only SharePoint search uses the Graph Search API with a region (default NAM).
export const resolveSharePointSearchConfig = (env: NodeJS.ProcessEnv = process.env): SharePointSearchConfig => ({
  region: blankToUndefined(env.GRAPH_SEARCH_REGION) ?? "NAM",
  defaultSiteId: blankToUndefined(env.SITE_ID),
  defaultSiteUrl: blankToUndefined(env.SITE_URL)?.replace(/\/$/, ""),
})

const blankToUndefined = (value?: string): string | undefined => {
  const trimmed = value?.trim()
  return trimmed === "" ? undefined : trimmed
}

// Azure AI Search is optional — present only when endpoint + api key + index are all set.
export const resolveAiSearchConfig = (env: NodeJS.ProcessEnv = process.env): AiSearchConfig | undefined => {
  const endpoint = blankToUndefined(env.AZURE_AI_SEARCH_ENDPOINT)
  const apiKey = env.AZURE_AI_SEARCH_API_KEY
  const indexName = blankToUndefined(env.AZURE_AI_SEARCH_INDEX)
  if (!endpoint || !apiKey || !indexName) return undefined
  return {
    endpoint: endpoint.replace(/\/$/, ""),
    apiKey,
    indexName,
    semanticConfiguration: blankToUndefined(env.AZURE_AI_SEARCH_SEMANTIC_CONFIG),
    vectorFields: blankToUndefined(env.AZURE_AI_SEARCH_VECTOR_FIELDS),
    selectFields: blankToUndefined(env.AZURE_AI_SEARCH_SELECT_FIELDS),
  }
}

const DEFAULT_APP_SCOPE = "https://graph.microsoft.com/.default"

// client_credentials requires a single concrete tenant — the multi-tenant aliases
// have no app identity to issue an app-only token against.
const MULTITENANT_ALIASES = new Set(["common", "organizations", "consumers"])

export const resolveAppOnlyConfig = (env: NodeJS.ProcessEnv = process.env): Either<string, AppOnlyConfig> => {
  const tenantId = env.MS_GRAPH_TENANT_ID?.trim() ?? ""
  const clientId = env.MS_GRAPH_CLIENT_ID?.trim() ?? ""
  const clientSecret = env.MS_GRAPH_CLIENT_SECRET ?? ""

  if (!tenantId) return Left("MS_GRAPH_TENANT_ID is required (app-only auth needs a concrete tenant).")
  if (MULTITENANT_ALIASES.has(tenantId.toLowerCase())) {
    return Left(`MS_GRAPH_TENANT_ID must be a concrete tenant, not "${tenantId}" — app-only auth cannot use it.`)
  }
  if (!clientId) return Left("MS_GRAPH_CLIENT_ID is required.")
  if (!clientSecret) return Left("MS_GRAPH_CLIENT_SECRET is required for app-only (client_credentials) auth.")

  const appScopes = (
    env.MS_GRAPH_APP_SCOPES?.split(",")
      .map((s) => s.trim())
      .filter((s) => s.length > 0) ?? []
  ).length
    ? (env.MS_GRAPH_APP_SCOPES as string)
        .split(",")
        .map((s) => s.trim())
        .filter((s) => s.length > 0)
    : [DEFAULT_APP_SCOPE]

  return Right({ tenantId, clientId, clientSecret, appScopes })
}

export const resolveServerRuntimeConfig = (env: NodeJS.ProcessEnv = process.env): Either<string, ServerRuntimeConfig> =>
  resolveAppOnlyConfig(env).map((auth) => {
    const trimmedKey = env.MCP_API_KEY?.trim()
    const port = parseInt(env.PORT ?? "8080", 10)
    const host = env.HOST ?? env.FASTMCP_HOST ?? "127.0.0.1"
    return {
      auth,
      apiKey: trimmedKey === "" ? undefined : trimmedKey,
      transport: env.TRANSPORT_TYPE === "stdio" ? ("stdio" as const) : ("httpStream" as const),
      port,
      host,
      publicBaseUrl: (blankToUndefined(env.MCP_PUBLIC_BASE_URL) ?? `http://${host}:${port}`).replace(/\/$/, ""),
      aiSearch: resolveAiSearchConfig(env),
      sharePointSearch: resolveSharePointSearchConfig(env),
    }
  })
