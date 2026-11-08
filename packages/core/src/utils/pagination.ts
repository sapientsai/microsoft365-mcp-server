import { type Either, Left, Right } from "functype/either"
import { Try } from "functype/try"

import { GRAPH_API_BASE } from "../constants"
import type { GraphApiError, ODataResponse } from "../types"

const MAX_PAGES = 50

const GRAPH_ORIGIN = new URL(GRAPH_API_BASE).origin

/**
 * Why a page's @odata.nextLink must not be followed, or undefined when it may be. The link comes
 * from Graph's response, and following it sends the bearer token, so it has to point back at Graph:
 * Graph's origin, no username or password, no fragment, and a /v1.0/ or /beta/ path. Each caller's
 * links have their own shapes, so the path is checked no further than the version.
 */
export const nextLinkProblem = (link: string): string | undefined => {
  const url = URL.parse(link)
  if (url === null) return "it isn't a URL"
  if (url.origin !== GRAPH_ORIGIN) return `it points at ${url.origin}, not ${GRAPH_ORIGIN}`
  if (url.username !== "" || url.password !== "") return "it carries credentials"
  if (url.hash !== "") return "it carries a fragment"
  if (!/^\/(?:v1\.0|beta)\//.test(url.pathname)) return "its path isn't under /v1.0/ or /beta/"
  return undefined
}

export const fetchAllPages = async <T>(
  fetchPage: (url: string) => Promise<Either<GraphApiError, ODataResponse<T>>>,
  initialUrl: string,
): Promise<Either<GraphApiError, ReadonlyArray<T>>> => {
  const allItems: T[] = []
  /* eslint-disable functype/no-let -- bounded paging loop; cursor and counter advance with each Graph @odata.nextLink */
  let currentUrl: string | null = initialUrl
  let pageCount = 0
  /* eslint-enable functype/no-let */

  while (currentUrl && pageCount < MAX_PAGES) {
    const result = await fetchPage(currentUrl)

    if (result.isLeft()) {
      return Left<GraphApiError, ReadonlyArray<T>>(result.value as GraphApiError)
    }

    const page = result.value as ODataResponse<T>
    allItems.push(...page.value)
    const next = page["@odata.nextLink"] ?? null
    const problem = next === null ? undefined : nextLinkProblem(next)
    if (problem !== undefined) {
      // No partial result: a caller can't tell a list cut short here from a complete one.
      return Left<GraphApiError, ReadonlyArray<T>>({
        type: "api",
        message: `Stopped paging after ${pageCount + 1} page(s): Graph's next-page link wasn't followed, because ${problem}.`,
      })
    }
    currentUrl = next
    pageCount++
  }

  return Right<GraphApiError, ReadonlyArray<T>>(allItems)
}

export const parseJsonResponse = <T>(text: string): Either<GraphApiError, T> =>
  Try(() => JSON.parse(text) as T).fold(
    (): Either<GraphApiError, T> =>
      Left<GraphApiError, T>({
        type: "parse",
        message: "Failed to parse Microsoft Graph API response",
      }),
    (data): Either<GraphApiError, T> => Right<GraphApiError, T>(data),
  )
