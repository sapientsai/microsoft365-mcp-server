import type { GraphApiError } from "@sapientsai/ms-graph-core"
import { type Either, Left, Right } from "functype/either"

// Azure AI Search REST client (separate service from Microsoft Graph — uses an api-key
// header and its own endpoint/index). Ported from microsoft-mcp-server onto core's
// GraphApiError. Optional capability; only wired when AZURE_AI_SEARCH_* env is set.

export const AI_SEARCH_API_VERSION = "2025-09-01"

export type AiSearchConfig = {
  readonly endpoint: string
  readonly apiKey: string
  readonly indexName: string
  readonly semanticConfiguration?: string
  readonly vectorFields?: string
  readonly selectFields?: string
}

export const parseAiSearchError = async (response: Response): Promise<string> => {
  const fallback = `HTTP ${response.status}: ${response.statusText}`
  try {
    const data = (await response.json()) as { error?: { message?: string; code?: string } }
    if (data.error?.message) {
      return data.error.code ? `${data.error.code}: ${data.error.message}` : data.error.message
    }
  } catch {
    // not JSON — fall through
  }
  return fallback
}

export const aiSearchFetch = async (
  url: string,
  apiKey: string,
  init: RequestInit = {},
  fetchImpl: typeof fetch = fetch,
): Promise<Either<GraphApiError, Response>> => {
  const response = await fetchImpl(url, {
    ...init,
    headers: { "api-key": apiKey, ...init.headers },
  }).catch((error: unknown) => error as Error)

  if (response instanceof Error) {
    return Left({ type: "network", message: `Azure AI Search request failed: ${response.message}` })
  }
  if (!response.ok) {
    return Left({ type: "api", message: await parseAiSearchError(response), status: response.status })
  }
  return Right(response)
}
