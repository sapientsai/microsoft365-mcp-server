import { GRAPH_API_BASE } from "@sapientsai/ms-graph-core"
import { UserError } from "fastmcp"
import type { Either } from "functype/either"
import { Left, Right } from "functype/either"

import { getGraphClient } from "../client/graph-client"
import type { GraphApiVersion } from "../types"
import { requireDraftEnabled } from "./tool-registry"

const requireClient = () => {
  const client = getGraphClient()
  if (client.isNone()) return null
  return client.orThrow()
}

const GRAPH_API_VERSIONS: ReadonlyArray<GraphApiVersion> = ["v1.0", "beta"]

const isGraphApiVersion = (value: string): value is GraphApiVersion =>
  (GRAPH_API_VERSIONS as ReadonlyArray<string>).includes(value)

// Graph actions that send mail at once, with no draft for anyone to review: sendMail, reply,
// replyAll and forward on a message, reply and forward on a group thread or post, and forward on
// an event. MS365_REQUIRE_DRAFT hides the send_* tools, but graph_query reaches the same endpoints,
// so it has to refuse them too or the flag is only advisory. Sending an existing draft
// (/messages/{id}/send) stays allowed: it is exactly what send_draft does. Teams replies use the
// plural "replies" and are not matched.
const SEND_ACTIONS: ReadonlySet<string> = new Set(["sendmail", "reply", "replyall", "forward"])

const safeDecode = (value: string): string => {
  try {
    return decodeURIComponent(value)
  } catch {
    return value
  }
}

// Match on the path fetch will actually send, not the string the caller typed. The WHATWG URL
// parser drops fragments, strips tabs, newlines and trailing spaces, turns "\" into "/", and
// resolves "." and ".." segments, so "/me/sendMail#x" or "/me/sendMail/x/.." would otherwise slip
// past a check on the raw string. Relative batch URLs resolve against the version root.
const wirePath = (url: string): string => {
  try {
    return new URL(url, `${GRAPH_API_BASE}/v1.0/`).pathname
  } catch {
    return url
  }
}

// Decoded before splitting, so an encoded "/" still separates segments. Graph paths are
// case-insensitive and accept fully qualified (microsoft.graph.reply) and call-style
// (sendMail()) action names, so all of those are normalized away.
const pathSegments = (url: string): ReadonlyArray<string> =>
  safeDecode(wirePath(url))
    .split(/[/\\]/)
    .map((segment) =>
      segment
        .trim()
        .toLowerCase()
        .replace(/^microsoft\.graph\./, "")
        .replace(/\(\)$/, ""),
    )
    .filter((segment) => segment.length > 0)

// Graph's batch parser may treat property names case-insensitively, so read them that way too.
const property = (value: unknown, name: string): unknown => {
  if (typeof value !== "object" || value === null) return undefined
  const key = Object.keys(value).find((k) => k.toLowerCase() === name)
  return key === undefined ? undefined : (value as Record<string, unknown>)[key]
}

/**
 * True when a request to this URL (a full URL, or a path relative to the Graph version root)
 * would send mail without going through a draft — directly, or as one of the requests inside a
 * JSON $batch. The method is deliberately ignored: none of these actions has a legitimate
 * non-POST use, and ignoring it closes off method-override headers.
 */
export const isDirectMailSend = (url: string, body?: unknown): boolean => {
  const action = pathSegments(url).at(-1)
  if (action === undefined) return false
  if (SEND_ACTIONS.has(action)) return true
  if (action !== "$batch") return false

  const requests = property(body, "requests")
  if (!Array.isArray(requests)) return false
  return requests.some((request) => {
    const inner = property(request, "url")
    return typeof inner === "string" && isDirectMailSend(inner, property(request, "body"))
  })
}

const DRAFT_REQUIRED_MESSAGE =
  "This server requires mail to go through a draft (MS365_REQUIRE_DRAFT), so graph_query can't send it " +
  "directly. Create a draft with create_draft, create_reply_draft, create_reply_all_draft or " +
  "create_forward_draft, then send it with send_draft."

const parseBody = (raw: string | undefined): Either<UserError, Record<string, unknown> | undefined> => {
  if (!raw) return Right(undefined)
  try {
    return Right(JSON.parse(raw) as Record<string, unknown>)
  } catch (err) {
    return Left(new UserError(`body is not valid JSON: ${err instanceof Error ? err.message : String(err)}`))
  }
}

export const graphQuery = async (params: {
  method: string
  path: string
  body?: string
  version?: string
  headers?: Record<string, string>
}): Promise<Either<UserError, string>> => {
  // The version is spliced into the URL, so anything but a known version could rewrite the path
  // ("v1.0/me/sendMail#") after the checks below have passed.
  if (params.version !== undefined && !isGraphApiVersion(params.version)) {
    return Left(new UserError(`version must be one of ${GRAPH_API_VERSIONS.join(", ")}`))
  }
  const { version } = params

  const parsed = parseBody(params.body)
  if (parsed.isLeft()) return Left(parsed.value as UserError)
  const body = parsed.value as Record<string, unknown> | undefined

  // Checked on every call, before any client work, against the same URL core will build.
  const url = `${GRAPH_API_BASE}/${version ?? "v1.0"}${params.path}`
  if (requireDraftEnabled() && isDirectMailSend(url, body)) {
    return Left(new UserError(DRAFT_REQUIRED_MESSAGE))
  }

  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.graphQuery(params.method, params.path, body, version, params.headers)
  return result
    .mapLeft((error) => new UserError(`Graph query failed: ${error.message}`))
    .map((data) => JSON.stringify(data, null, 2))
}
