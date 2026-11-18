// Graph's error body is `{ "error": { "code", "message", "innerError" } }`. On an outage or a
// permission failure Graph can send `"code": "UnknownError"` with an empty message, and the HTTP
// status and the code are then all that tell a 503 from a 403. So both always go into the text,
// and a blank message counts as missing.

const isRecord = (value: unknown): value is Record<string, unknown> => typeof value === "object" && value !== null

const trimmedString = (value: unknown): string => (typeof value === "string" ? value.trim() : "")

/**
 * The text of a failed Graph call: Graph's own message (or a generic one when it sent none),
 * followed by the HTTP status and the Graph error code, e.g.
 * `Microsoft Graph API error (HTTP 503 Service Unavailable, UnknownError)`.
 *
 * `body` is the parsed response body, or `undefined` when it was not JSON.
 */
export const graphErrorMessage = (status: number, statusText: string, body: unknown): string => {
  const error = isRecord(body) && isRecord(body.error) ? body.error : undefined
  const message = trimmedString(error?.message)
  const code = trimmedString(error?.code)
  const httpStatus = [String(status), trimmedString(statusText)].filter((part) => part !== "").join(" ")
  const detail = [`HTTP ${httpStatus}`, code].filter((part) => part !== "").join(", ")
  return `${message === "" ? "Microsoft Graph API error" : message} (${detail})`
}
