import { type Either, Left, Right } from "functype/either"

import type { AuthStrategy } from "./auth-strategy"
import { GRAPH_API_BASE } from "./constants"
import { graphErrorMessage } from "./graph-error-message"
import { notSentError } from "./request-check"
import { type FetchLike, isCircuitOpenError, resilientFetch } from "./resilience"
import type { GraphApiError, GraphApiVersion, ODataParams, ODataResponse } from "./types"
import { graphPathProblem } from "./utils/graph-path"
import { appendODataQuery, buildODataQuery } from "./utils/odata-helpers"
import { fetchAllPages, parseJsonResponse } from "./utils/pagination"

export type GraphRequestOptions = {
  readonly version?: GraphApiVersion
  readonly body?: Record<string, unknown> | readonly unknown[] | string
  readonly contentType?: string
  readonly responseType?: "json" | "text"
  readonly odataParams?: ODataParams
  readonly headers?: Record<string, string>
  /**
   * false stops retries on 503, 504, timeouts and network errors (429 is still retried). For a
   * DELETE whose retry would turn a delete that landed into a "not found" failure.
   */
  readonly retryTransient?: boolean
}

export type GraphRequestConfig = {
  // Resolved per call (servers may read an env var at request time), so it's a function.
  readonly defaultVersion?: () => GraphApiVersion
  // Injectable for tests. Defaults to the retry/timeout/circuit-breaker wrapper;
  // pass plain `fetch` to opt a caller out of resilience entirely.
  readonly fetchImpl?: FetchLike
}

export type GraphRequest = {
  request: <T>(method: string, path: string, options?: GraphRequestOptions) => Promise<Either<GraphApiError, T>>
  requestPaginated: <T>(path: string, options?: GraphRequestOptions) => Promise<Either<GraphApiError, ReadonlyArray<T>>>
}

// Maps a non-OK Graph HTTP response to a typed GraphApiError. Exported for the few
// methods that do their own fetch (binary download, abort-controlled upload).
export const mapHttpError = async <T>(response: Response): Promise<Either<GraphApiError, T>> => {
  const { message, graphErrorCode, graphInnerErrorCode } = await (async (): Promise<{
    message: string
    graphErrorCode?: string
    graphInnerErrorCode?: string
  }> => {
    try {
      const errorBody = await response.json()
      return {
        message: graphErrorMessage(response.status, response.statusText, errorBody),
        graphErrorCode: errorBody?.error?.code as string | undefined,
        graphInnerErrorCode: errorBody?.error?.innerError?.code as string | undefined,
      }
    } catch {
      return { message: graphErrorMessage(response.status, response.statusText, undefined) }
    }
  })()

  const retryAfter = response.headers.get("Retry-After")

  switch (response.status) {
    case 401:
      return Left<GraphApiError, T>({ type: "auth", message, status: 401, graphErrorCode, graphInnerErrorCode })
    case 403:
      return Left<GraphApiError, T>({ type: "forbidden", message, status: 403, graphErrorCode, graphInnerErrorCode })
    case 404:
      return Left<GraphApiError, T>({ type: "not_found", message, status: 404, graphErrorCode, graphInnerErrorCode })
    case 429:
      return Left<GraphApiError, T>({
        type: "throttle",
        message,
        status: 429,
        graphErrorCode,
        graphInnerErrorCode,
        retryAfter: retryAfter ? parseInt(retryAfter, 10) : undefined,
      })
    default:
      return Left<GraphApiError, T>({
        type: "api",
        message,
        status: response.status,
        graphErrorCode,
        graphInnerErrorCode,
      })
  }
}

// Generic Microsoft Graph request layer: auth via an injected AuthStrategy, OData query
// building, typed error mapping, and auto-pagination. Domain methods (listMessages, etc.)
// are built on top of this in each server.
// A tripped circuit is a throttle-shaped condition, not a network fault: the caller
// should back off, not retry immediately. Map it onto the existing typed error union
// so no call site needs to learn a new failure mode.
const circuitOpenError = (error: { message: string; cooldownMs: number }): GraphApiError => ({
  type: "throttle",
  message: error.message,
  retryAfter: Math.ceil(error.cooldownMs / 1000),
})

/**
 * The headers a request goes out with. A caller's Content-Type replaces the default, whatever its
 * case, instead of being joined to it. Authorization is always this layer's own: a caller's is
 * dropped, so a caller can't swap the token or have fetch join two into one header.
 */
const requestHeaders = (
  token: string,
  defaultContentType: string,
  caller: Record<string, string> | undefined,
): Record<string, string> => {
  const own = Object.entries(caller ?? {}).filter(([name]) => name.toLowerCase() !== "authorization")
  const setsContentType = own.some(([name]) => name.toLowerCase() === "content-type")
  return {
    ...(setsContentType ? {} : { "Content-Type": defaultContentType }),
    ...Object.fromEntries(own),
    Authorization: `Bearer ${token}`,
  }
}

const BODY_METHODS: ReadonlySet<string> = new Set(["POST", "PUT", "PATCH"])

export const createGraphRequest = (auth: AuthStrategy, config: GraphRequestConfig = {}): GraphRequest => {
  const resolveVersion = config.defaultVersion ?? (() => "v1.0" as const)
  const doFetch = config.fetchImpl ?? resilientFetch

  const request = async <T>(
    method: string,
    path: string,
    options?: GraphRequestOptions,
  ): Promise<Either<GraphApiError, T>> => {
    const pathIssue = graphPathProblem(path)
    if (pathIssue !== undefined) return Left<GraphApiError, T>({ type: "api", status: 400, message: pathIssue })

    const tokenResult = await auth.getAccessToken()
    if (tokenResult.isLeft()) {
      return Left<GraphApiError, T>({ type: "auth", message: (tokenResult.value as { message: string }).message })
    }

    const token = tokenResult.value as string
    const version = options?.version ?? resolveVersion()
    const queryString = buildODataQuery(options?.odataParams)
    const url = `${GRAPH_API_BASE}/${version}${appendODataQuery(path, queryString)}`

    // fetch upper-cases only some methods, and would send "patch" as written. Comparing the
    // upper-cased form also keeps a lower-case "post" from silently dropping its body.
    const verb = method.toUpperCase()

    try {
      const fetchOptions: RequestInit = {
        method: verb,
        headers: requestHeaders(token, options?.contentType ?? "application/json", options?.headers),
      }

      if (options?.body !== undefined && BODY_METHODS.has(verb)) {
        fetchOptions.body = typeof options.body === "string" ? options.body : JSON.stringify(options.body)
      }

      const response = await doFetch(
        url,
        fetchOptions,
        options?.retryTransient === undefined ? undefined : { retryTransient: options.retryTransient },
      )

      if (!response.ok) return mapHttpError<T>(response)
      if (response.status === 204) return Right<GraphApiError, T>({} as T)

      const text = await response.text()
      if (!text || text.trim() === "") return Right<GraphApiError, T>({} as T)
      if (options?.responseType === "text") return Right<GraphApiError, T>(text as T)

      return parseJsonResponse<T>(text)
    } catch (error) {
      if (isCircuitOpenError(error)) return Left<GraphApiError, T>(circuitOpenError(error))
      const notSent = notSentError(error)
      if (notSent) return Left<GraphApiError, T>(notSent)
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
    // Only the first path is the caller's; the pages after it come from Graph's own nextLink.
    const pathIssue = graphPathProblem(path)
    if (pathIssue !== undefined) {
      return Left<GraphApiError, ReadonlyArray<T>>({ type: "api", status: 400, message: pathIssue })
    }
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
        const response = await doFetch(url, { headers: requestHeaders(token, "application/json", options?.headers) })
        if (!response.ok) return mapHttpError<ODataResponse<T>>(response)
        const text = await response.text()
        return parseJsonResponse<ODataResponse<T>>(text)
      } catch (error) {
        if (isCircuitOpenError(error)) return Left<GraphApiError, ODataResponse<T>>(circuitOpenError(error))
        const notSent = notSentError(error)
        if (notSent) return Left<GraphApiError, ODataResponse<T>>(notSent)
        return Left<GraphApiError, ODataResponse<T>>({
          type: "network",
          message: `Network error: ${error instanceof Error ? error.message : String(error)}`,
        })
      }
    }, initialUrl)
  }

  return { request, requestPaginated }
}
