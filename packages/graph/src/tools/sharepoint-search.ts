import type { GraphApiError, GraphRequest } from "@sapientsai/ms-graph-core"
import { type Either, Left, Right } from "functype/either"
import { z } from "zod"

// SharePoint document-library search for the app-only server. Uses the Graph Search API
// (/search/query) with a region — which is what app-only (client_credentials) always does
// (the gateway defaults the region to "NAM" for that mode, so its drive-fan-out + site-cache
// path is unreachable here; intentionally not ported). Ported onto core's GraphRequest.

export type SharePointSearchConfig = {
  readonly region: string
  readonly defaultSiteId?: string
  readonly defaultSiteUrl?: string
}

type SearchResult = {
  name: string
  driveItemId: string
  driveId: string
  siteId: string
  webUrl: string
  lastModified: string
  size: number
  mimeType: string
  hitHighlights: string[]
}

type SearchHit = {
  resource?: {
    id?: string
    name?: string
    webUrl?: string
    lastModifiedDateTime?: string
    size?: number
    file?: { mimeType?: string }
    parentReference?: { driveId?: string; siteId?: string }
  }
  _summary?: string
}
type SearchResponse = {
  value?: ReadonlyArray<{ hitsContainers?: ReadonlyArray<{ hits?: readonly SearchHit[] }> }>
}

const searchParameters = z.object({
  query: z.string().describe("Search query keywords"),
  siteId: z.string().optional().describe("Scope search to a specific SharePoint site ID"),
  top: z.number().int().min(1).max(50).default(10).describe("Max results (1-50)"),
  fileTypes: z.array(z.string()).optional().describe('Filter by file extensions, e.g. ["docx", "pdf", "xlsx"]'),
})

const mapHit = (hit: SearchHit): SearchResult => {
  const r = hit.resource ?? {}
  return {
    name: r.name ?? "",
    driveItemId: r.id ?? "",
    driveId: r.parentReference?.driveId ?? "",
    siteId: r.parentReference?.siteId ?? "",
    webUrl: r.webUrl ?? "",
    lastModified: r.lastModifiedDateTime ?? "",
    size: r.size ?? 0,
    mimeType: r.file?.mimeType ?? "",
    hitHighlights: hit._summary ? [hit._summary] : [],
  }
}

// KQL site: clause requires a site URL, not a composite id. Prefer the configured default
// site URL (no API call); otherwise resolve an explicit siteId to its webUrl.
const resolveSiteClause = async (
  graph: GraphRequest,
  args: { siteId?: string },
  config: SharePointSearchConfig,
): Promise<Either<GraphApiError, string>> => {
  if (config.defaultSiteUrl && !args.siteId) return Right(` site:${config.defaultSiteUrl}`)
  if (!args.siteId) return Right("")
  const result = await graph.request<{ webUrl?: string }>("GET", `/sites/${args.siteId}`, {
    odataParams: { $select: ["webUrl"] },
  })
  return result.fold<Either<GraphApiError, string>>(
    (error) => Left(error),
    (site) =>
      site.webUrl
        ? Right(` site:${site.webUrl}`)
        : Left({ type: "not_found", message: `Site ${args.siteId} has no webUrl.`, status: 404 }),
  )
}

export const buildSharePointSearchTool = (graph: GraphRequest, config: SharePointSearchConfig) => ({
  name: "sharepoint_search" as const,
  description:
    "Search SharePoint document libraries by keyword. Returns file metadata (name, driveId, driveItemId) usable " +
    "with read_document via /drives/{driveId}/items/{driveItemId}/content. Supports site and file-type filtering.",
  parameters: searchParameters,
  execute: async (args: z.infer<typeof searchParameters>): Promise<string> => {
    const effectiveSiteId = config.defaultSiteId && !args.siteId ? config.defaultSiteId : args.siteId
    const siteClauseResult = await resolveSiteClause(graph, { siteId: effectiveSiteId }, config)
    const siteClause = siteClauseResult.fold<string>(
      (error) => {
        throw new Error(error.message)
      },
      (clause) => clause,
    )

    const fileTypeClause =
      args.fileTypes && args.fileTypes.length > 0
        ? ` (${args.fileTypes.map((ext) => `filetype:${ext.replace(/^\./, "")}`).join(" OR ")})`
        : ""
    const queryString = `${args.query}${fileTypeClause}${siteClause}`

    const result = await graph.request<SearchResponse>("POST", "/search/query", {
      body: {
        requests: [
          { entityTypes: ["driveItem"], query: { queryString }, from: 0, size: args.top, region: config.region },
        ],
      },
    })

    const data = result.fold<SearchResponse>(
      (error) => {
        throw new Error(error.message)
      },
      (res) => res,
    )

    const hits = data.value?.[0]?.hitsContainers?.[0]?.hits ?? []
    const results = hits.map(mapHit)
    return JSON.stringify({ results, totalCount: results.length }, null, 2)
  },
})
