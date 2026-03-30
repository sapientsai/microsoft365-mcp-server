import { type Either, Left, Right } from "functype/either"
import { Try } from "functype/try"

import type { GraphApiError, ODataResponse } from "../types"

const MAX_PAGES = 50

export const fetchAllPages = async <T>(
  fetchPage: (url: string) => Promise<Either<GraphApiError, ODataResponse<T>>>,
  initialUrl: string,
): Promise<Either<GraphApiError, ReadonlyArray<T>>> => {
  const allItems: T[] = []
  let currentUrl: string | null = initialUrl
  let pageCount = 0

  while (currentUrl && pageCount < MAX_PAGES) {
    const result = await fetchPage(currentUrl)

    if (result.isLeft()) {
      return Left<GraphApiError, ReadonlyArray<T>>(result.value as GraphApiError)
    }

    const page = result.value as ODataResponse<T>
    allItems.push(...page.value)
    currentUrl = page["@odata.nextLink"] ?? null
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
