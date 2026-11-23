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
  /** Where callers reach this server. Undefined when it can't be told (see resolvePublicBaseUrl). */
  readonly publicBaseUrl?: string
  readonly aiSearch?: AiSearchConfig
  readonly enableDownloadFile: boolean
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

// A wildcard bind ("listen on every interface") names no address a caller can reach. An empty HOST
// (a "HOST=" line) and "0" bind every interface too, as do the IPv4 and IPv6 all-zeros addresses.
const isWildcardHost = (host: string): boolean => {
  const bare = host.trim().replace(/^\[(.*)\]$/, "$1")
  return bare === "" || /^0+(\.0+){0,3}$/.test(bare) || /^[0:]+$/.test(bare)
}

// An IPv6 literal goes in brackets in a URL: http://[::1]:8080, not http://::1:8080.
const urlHost = (host: string): string => (host.includes(":") && !host.startsWith("[") ? `[${host}]` : host)

/**
 * The URL callers use to reach this server: MCP_PUBLIC_BASE_URL, or the bind address when that is
 * a real one (127.0.0.1 for a local run). A wildcard bind, as in the Docker image, gives undefined:
 * http://0.0.0.0:8080 works only from the same host or network namespace, not from a caller elsewhere, so get_upload_config refuses
 * rather than hand out a curl that can't connect.
 */
export const resolvePublicBaseUrl = (env: NodeJS.ProcessEnv, host: string, port: number): string | undefined => {
  const configured = blankToUndefined(env.MCP_PUBLIC_BASE_URL)
  if (configured !== undefined) return configured.replace(/\/$/, "")
  return isWildcardHost(host) ? undefined : `http://${urlHost(host.trim())}:${port}`
}

const optionalPublicBaseUrl = (env: NodeJS.ProcessEnv, host: string, port: number) => {
  const url = resolvePublicBaseUrl(env, host, port)
  return url === undefined ? {} : { publicBaseUrl: url }
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
      ...optionalPublicBaseUrl(env, host, port),
      aiSearch: resolveAiSearchConfig(env),
      // Defaults ON: the eight deployed app-only containers already expose download_file, and
      // they lose it the moment they re-pull :main. Opt-out, not opt-in.
      enableDownloadFile: env.GRAPH_ENABLE_DOWNLOAD_FILE?.trim().toLowerCase() !== "false",
      sharePointSearch: resolveSharePointSearchConfig(env),
    }
  })
