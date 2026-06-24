import type { GraphApiError } from "@sapientsai/ms-graph-core"
import type { Either } from "functype/either"
import { z } from "zod"

import { AI_SEARCH_API_VERSION, type AiSearchConfig, aiSearchFetch } from "../search/ai-search-client"

type AiSearchHit = {
  "@search.score"?: number
  "@search.rerankerScore"?: number
  "@search.captions"?: ReadonlyArray<{ text?: string }>
  "@search.highlights"?: Record<string, readonly string[]>
  [key: string]: unknown
}

type AiSearchResponse = {
  "@odata.count"?: number
  "@search.answers"?: ReadonlyArray<{ text?: string; key?: string; score?: number }>
  value?: readonly AiSearchHit[]
}

const aiSearchParameters = z.object({
  query: z.string().describe("Search text"),
  queryType: z
    .enum(["simple", "full", "semantic"])
    .default("semantic")
    .describe('Query type: "simple" keyword, "full" Lucene, or "semantic" for AI-ranked results'),
  vectorSearch: z
    .boolean()
    .default(false)
    .describe("Enable hybrid search via server-side vectorization (index must have vector fields)"),
  filter: z.string().optional().describe("OData filter expression"),
  select: z.string().optional().describe("Comma-separated fields to return"),
  top: z.number().int().min(1).max(50).default(10).describe("Max results (1-50)"),
  skip: z.number().int().optional().describe("Results to skip (pagination)"),
  orderby: z.string().optional().describe("OData $orderby expression"),
  highlightFields: z.string().optional().describe("Comma-separated fields to highlight"),
  includeTotalCount: z.boolean().default(false).describe("Include total count of matching documents"),
})

const mapHit = (hit: AiSearchHit) => {
  const {
    "@search.score": score,
    "@search.rerankerScore": rerankerScore,
    "@search.captions": captions,
    "@search.highlights": highlights,
    ...document
  } = hit
  return {
    score: score ?? 0,
    ...(rerankerScore !== undefined ? { rerankerScore } : {}),
    document,
    captions: (captions ?? []).flatMap((c) => (c.text ? [c.text] : [])),
    highlights: highlights ?? {},
  }
}

export const buildAiSearchTool = (config: AiSearchConfig, fetchImpl: typeof fetch = fetch) => {
  const searchUrl = `${config.endpoint}/indexes/${config.indexName}/docs/search?api-version=${AI_SEARCH_API_VERSION}`

  return {
    name: "azure_ai_search" as const,
    description:
      "Search an Azure AI Search index with text, semantic, or hybrid (text + vector) queries. Returns ranked " +
      "results with scores and optional captions/highlights.",
    parameters: aiSearchParameters,
    execute: async (args: z.infer<typeof aiSearchParameters>): Promise<string> => {
      if (args.queryType === "semantic" && !config.semanticConfiguration) {
        throw new Error(
          "Semantic search requires a semantic configuration (AZURE_AI_SEARCH_SEMANTIC_CONFIG). Use queryType 'simple' or 'full'.",
        )
      }

      const body: Record<string, unknown> = {
        search: args.query,
        queryType: args.queryType,
        top: args.top,
        count: args.includeTotalCount,
      }
      if (args.filter) body.filter = args.filter
      if (args.skip) body.skip = args.skip
      if (args.orderby) body.orderby = args.orderby
      if (args.highlightFields) body.highlightFields = args.highlightFields

      const selectFields = args.select ?? config.selectFields
      if (selectFields) body.select = selectFields

      if (args.queryType === "semantic") {
        body.semanticConfiguration = config.semanticConfiguration
        body.captions = "extractive"
        body.answers = "extractive"
      }
      if (args.vectorSearch) {
        body.vectorQueries = [
          { kind: "text", text: args.query, fields: config.vectorFields ?? "contentVector", k: args.top },
        ]
      }

      const result: Either<GraphApiError, Response> = await aiSearchFetch(
        searchUrl,
        config.apiKey,
        { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) },
        fetchImpl,
      )

      const response = result.fold<Response>(
        (error) => {
          throw new Error(error.message)
        },
        (res) => res,
      )

      const data = (await response.json()) as AiSearchResponse
      const output: Record<string, unknown> = { results: (data.value ?? []).map(mapHit) }
      if (args.includeTotalCount && data["@odata.count"] !== undefined) output.totalCount = data["@odata.count"]
      if (data["@search.answers"] && data["@search.answers"].length > 0) output.answers = data["@search.answers"]

      return JSON.stringify(output, null, 2)
    },
  }
}
