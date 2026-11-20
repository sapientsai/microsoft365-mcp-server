import type { GraphApiError } from "./types"

/**
 * Checks for a request that fetch would refuse to send.
 *
 * fetch throws on a malformed method or header before anything goes on the wire, and the
 * resilience layer would read that throw as a network failure: retried, and counted against the
 * process-wide circuit breaker. A few bad calls would then fail every tool for every user. So a
 * request is checked first, and one that can't be sent fails with RequestNotSentError, which is
 * neither retried nor counted.
 *
 * Messages never echo a header value: it may be a bearer token, and it may hold a line break.
 */

export class RequestNotSentError extends Error {
  readonly code = "request_not_sent"
  constructor(problem: string) {
    super(`Request not sent: ${problem}`)
    this.name = "RequestNotSentError"
  }
}

export const isRequestNotSentError = (err: unknown): err is RequestNotSentError => err instanceof RequestNotSentError

/** The 400 a refused request maps to, or undefined for any other error. */
export const notSentError = (err: unknown): GraphApiError | undefined =>
  isRequestNotSentError(err) ? { type: "api", status: 400, message: err.message } : undefined

/** A caller-chosen string, quoted and cut short so it can't add lines to a message. */
export const quoteForMessage = (value: string, max = 60): string => JSON.stringify(value.slice(0, max))

// RFC 9110 token characters: what a method or header name may hold.
const TOKEN = /^[!#$%&'*+.^_`|~0-9A-Za-z-]+$/

const UNSUPPORTED_METHODS: ReadonlySet<string> = new Set(["CONNECT", "TRACE", "TRACK"])

// undici refuses all but Connection at dispatch, after the Headers check has passed. Connection
// is refused as policy: fetch manages the connection, and no Graph call needs to set it.
const REFUSED_HEADERS: ReadonlySet<string> = new Set([
  "connection",
  "expect",
  "keep-alive",
  "transfer-encoding",
  "upgrade",
])

// What undici sends once fetch has trimmed the value: tab, visible ASCII, space, and Latin-1.
const SENDABLE_VALUE = /^[\t\x20-\x7e\x80-\xff]*$/

// fetch trims these from both ends of a value before checking it.
const trimHeaderValue = (value: string): string => value.replace(/^[\t\n\r ]+|[\t\n\r ]+$/g, "")

const headerEntries = (headers: HeadersInit | undefined): ReadonlyArray<readonly [string, string]> => {
  if (headers === undefined) return []
  if (headers instanceof Headers) return [...headers.entries()]
  if (Array.isArray(headers)) return headers.map((pair) => [String(pair[0]), String(pair[1])] as const)
  // fetch sends a value that isn't a string (undefined, a number) as its string form.
  return Object.entries(headers as Record<string, unknown>).map(([name, value]) => [name, String(value)] as const)
}

const bodyByteLength = (body: RequestInit["body"]): number | undefined => {
  if (body === undefined || body === null) return 0
  if (typeof body === "string") return Buffer.byteLength(body, "utf8")
  if (body instanceof ArrayBuffer) return body.byteLength
  if (ArrayBuffer.isView(body)) return body.byteLength
  return undefined
}

const contentLengthProblem = (value: string, body: RequestInit["body"]): string | undefined => {
  if (!/^\d+$/.test(value)) return "the Content-Length header is not a number."
  const length = bodyByteLength(body)
  // A wrong count makes undici send the headers and then fail, which would count as an outage.
  if (length !== undefined && Number(value) !== length) return "the Content-Length header doesn't match the body."
  return undefined
}

const headerProblem = (name: string, value: string, body: RequestInit["body"]): string | undefined => {
  if (!TOKEN.test(name)) return `the header name ${quoteForMessage(name)} is not a valid HTTP header name.`
  const lower = name.toLowerCase()
  // A valid name holds no space or line break, but may still be long.
  const label = name.slice(0, 60)
  if (REFUSED_HEADERS.has(lower)) return `the ${label} header can't be set on a Graph request.`
  const trimmed = trimHeaderValue(value)
  if (!SENDABLE_VALUE.test(trimmed)) {
    return `the ${label} header holds a character a header can't carry, such as a line break or a control character.`
  }
  return lower === "content-length" ? contentLengthProblem(trimmed, body) : undefined
}

/** Why fetch would refuse to send this request, or undefined when it can be sent. */
export const requestProblem = (url: string, init: RequestInit | undefined): string | undefined => {
  if (!URL.canParse(url)) return "the request URL is not a valid URL."
  const parsed = new URL(url)
  if (parsed.protocol !== "https:" && parsed.protocol !== "http:") return "the request URL is not an http(s) URL."
  // fetch refuses a URL carrying a user name or password.
  if (parsed.username !== "" || parsed.password !== "") return "the request URL holds a user name or password."
  const method = init?.method ?? "GET"
  if (!TOKEN.test(method)) return `the HTTP method ${quoteForMessage(method, 20)} is not valid.`
  const upper = method.toUpperCase()
  if (UNSUPPORTED_METHODS.has(upper)) return `the HTTP method ${method} is not supported.`
  if ((upper === "GET" || upper === "HEAD") && init?.body !== undefined && init.body !== null) {
    return `a ${upper} request can't carry a body.`
  }
  return headerEntries(init?.headers).reduce<string | undefined>(
    (problem, [name, value]) => problem ?? headerProblem(name, value, init?.body),
    undefined,
  )
}

// undici's codes for a request it refused to build. Nothing was sent, so the outage breaker
// must not hear about it. This catches what a future undici refuses that requestProblem misses.
// UND_ERR_REQ_CONTENT_LENGTH_MISMATCH is not here: undici raises it after the headers have gone
// out, so the request did reach the server. requestProblem refuses a mismatch before that.
const REFUSED_BEFORE_SENDING: ReadonlySet<string> = new Set(["UND_ERR_INVALID_ARG", "UND_ERR_NOT_SUPPORTED"])

/** True when fetch threw because it refused to build the request, not because the network failed. */
export const wasRefusedBeforeSending = (err: unknown): boolean => {
  const cause = err instanceof Error ? (err as { cause?: unknown }).cause : undefined
  const code = typeof cause === "object" && cause !== null ? (cause as { code?: unknown }).code : undefined
  return typeof code === "string" && REFUSED_BEFORE_SENDING.has(code)
}
