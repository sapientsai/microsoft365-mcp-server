import { type Either, Left, Right } from "functype/either"

import type { AuthStrategy } from "./auth-strategy"
import { GRAPH_API_BASE } from "./constants"
import type { GraphApiError, GraphApiVersion, ODataParams, ODataResponse } from "./types"
import { appendODataQuery, buildODataQuery } from "./utils/odata-helpers"
import { fetchAllPages, parseJsonResponse } from "./utils/pagination"

export type GraphRequestOptions = {
  readonly version?: GraphApiVersion
  readonly body?: Record<string, unknown> | readonly unknown[] | string
  readonly contentType?: string
  readonly responseType?: "json" | "text"
  readonly odataParams?: ODataParams
  readonly headers?: Record<string, string>
}

export type GraphRequestConfig = {
  // Resolved per call (servers may read an env var at request time), so it's a function.
  readonly defaultVersion?: () => GraphApiVersion
}

export type GraphRequest = {
  request: <T>(method: string, path: string, options?: GraphRequestOptions) => Promise<Either<GraphApiError, T>>
  requestPaginated: <T>(path: string, options?: GraphRequestOptions) => Promise<Either<GraphApiError, ReadonlyArray<T>>>
}

// Maps a non-OK Graph HTTP response to a typed GraphApiError. Exported for the few
// methods that do their own fetch (binary download, abort-controlled upload).
export const mapHttpError = async <T>(response: Response): Promise<Either<GraphApiError, T>> => {
  const fallbackMessage = `Microsoft Graph API error: ${response.status} ${response.statusText}`

  const { message, graphErrorCode } = await (async (): Promise<{ message: string; graphErrorCode?: string }> => {
    try {
      const errorBody = await response.json()
      return {
        message: (errorBody?.error?.message as string | undefined) ?? fallbackMessage,
        graphErrorCode: errorBody?.error?.code as string | undefined,
      }
    } catch {
      return { message: fallbackMessage }
    }
  })()

  const retryAfter = response.headers.get("Retry-After")

  switch (response.status) {
    case 401:
      return Left<GraphApiError, T>({ type: "auth", message, status: 401, graphErrorCode })
    case 403:
      return Left<GraphApiError, T>({ type: "forbidden", message, status: 403, graphErrorCode })
    case 404:
      return Left<GraphApiError, T>({ type: "not_found", message, status: 404, graphErrorCode })
    case 429:
      return Left<GraphApiError, T>({
        type: "throttle",
        message,
        status: 429,
        graphErrorCode,
        retryAfter: retryAfter ? parseInt(retryAfter, 10) : undefined,
      })
    default:
      return Left<GraphApiError, T>({ type: "api", message, status: response.status, graphErrorCode })
  }
}

// Generic Microsoft Graph request layer: auth via an injected AuthStrategy, OData query
// building, typed error mapping, and auto-pagination. Domain methods (listMessages, etc.)
// are built on top of this in each server.
export const createGraphRequest = (auth: AuthStrategy, config: GraphRequestConfig = {}): GraphRequest => {
  const resolveVersion = config.defaultVersion ?? (() => "v1.0" as const)

  const request = async <T>(
    method: string,
    path: string,
    options?: GraphRequestOptions,
  ): Promise<Either<GraphApiError, T>> => {
    const tokenResult = await auth.getAccessToken()
    if (tokenResult.isLeft()) {
      return Left<GraphApiError, T>({ type: "auth", message: (tokenResult.value as { message: string }).message })
    }

    const token = tokenResult.value as string
    const version = options?.version ?? resolveVersion()
    const queryString = buildODataQuery(options?.odataParams)
    const url = `${GRAPH_API_BASE}/${version}${appendODataQuery(path, queryString)}`

    try {
      const fetchOptions: RequestInit = {
        method,
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": options?.contentType ?? "application/json",
          ...(options?.headers ?? {}),
        },
      }

      if (options?.body !== undefined && (method === "POST" || method === "PUT" || method === "PATCH")) {
        fetchOptions.body = typeof options.body === "string" ? options.body : JSON.stringify(options.body)
      }

      const response = await fetch(url, fetchOptions)

      if (!response.ok) return mapHttpError<T>(response)
      if (response.status === 204) return Right<GraphApiError, T>({} as T)

      const text = await response.text()
      if (!text || text.trim() === "") return Right<GraphApiError, T>({} as T)
      if (options?.responseType === "text") return Right<GraphApiError, T>(text as T)

      return parseJsonResponse<T>(text)
    } catch (error) {
      return Left<GraphApiError, T>({
        type: "network",
        message: `Network error: ${error instanceof Error ? error.message : String(error)}`,
      })
    }
  }

  const requestPaginated = async <T>(
    path: string,
    options?: GraphRequestOptions,
  ): Promise<Either<GraphApiError, ReadonlyArray<T>>> => {
    const version = options?.version ?? resolveVersion()
    const queryString = buildODataQuery(options?.odataParams)
    const initialUrl = `${GRAPH_API_BASE}/${version}${appendODataQuery(path, queryString)}`

    return fetchAllPages<T>(async (url: string) => {
      const tokenResult = await auth.getAccessToken()
      if (tokenResult.isLeft()) {
        return Left<GraphApiError, ODataResponse<T>>({
          type: "auth",
          message: (tokenResult.value as { message: string }).message,
        })
      }

      const token = tokenResult.value as string
      try {
        const response = await fetch(url, {
          headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": "application/json",
            ...(options?.headers ?? {}),
          },
        })
        if (!response.ok) return mapHttpError<ODataResponse<T>>(response)
        const text = await response.text()
        return parseJsonResponse<ODataResponse<T>>(text)
      } catch (error) {
        return Left<GraphApiError, ODataResponse<T>>({
          type: "network",
          message: `Network error: ${error instanceof Error ? error.message : String(error)}`,
        })
      }
    }, initialUrl)
  }

  return { request, requestPaginated }
}
