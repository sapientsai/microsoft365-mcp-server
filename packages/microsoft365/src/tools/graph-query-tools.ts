import { UserError } from "fastmcp"
import type { Either } from "functype/either"
import { Left } from "functype/either"

import { getGraphClient } from "../client/graph-client"
import type { GraphApiVersion } from "../types"

const requireClient = () => {
  const client = getGraphClient()
  if (client.isNone()) return null
  return client.orThrow()
}

// Graph actions that send mail at once, with no draft for anyone to review. MS365_REQUIRE_DRAFT
// hides the send_* tools, but graph_query reaches the same endpoints, so it has to refuse them
// too or the flag is only advisory. Sending an existing draft (/messages/{id}/send) stays allowed:
// it is exactly what send_draft does.
const SEND_MAIL_ACTION = "sendmail"
const MESSAGE_SEND_ACTIONS: ReadonlySet<string> = new Set(["reply", "replyall", "forward"])

const decodeSegment = (segment: string): string => {
  try {
    return decodeURIComponent(segment)
  } catch {
    return segment
  }
}

// Graph paths are case-insensitive and accept fully qualified action names
// (/messages/{id}/microsoft.graph.reply), so both are normalized away before matching.
const pathSegments = (path: string): ReadonlyArray<string> =>
  (path.split("?")[0] ?? "")
    .split("/")
    .filter((segment) => segment.length > 0)
    .map((segment) =>
      decodeSegment(segment)
        .toLowerCase()
        .replace(/^microsoft\.graph\./, ""),
    )

const isMessagesSegment = (segment: string): boolean => segment === "messages" || segment.startsWith("messages(")

const isDirectSendPath = (path: string): boolean => {
  const segments = pathSegments(path)
  const action = segments.at(-1)
  if (action === undefined) return false
  if (action === SEND_MAIL_ACTION) return true
  return MESSAGE_SEND_ACTIONS.has(action) && segments.some(isMessagesSegment)
}

type BatchRequest = { readonly method?: unknown; readonly url?: unknown }

const batchRequests = (body: Record<string, unknown> | undefined): ReadonlyArray<BatchRequest> =>
  Array.isArray(body?.requests) ? (body.requests as ReadonlyArray<BatchRequest>) : []

/**
 * True when the request would send mail without going through a draft — directly, or as one
 * of the requests inside a JSON $batch.
 */
export const isDirectMailSend = (method: string, path: string, body?: Record<string, unknown>): boolean => {
  if (method.toUpperCase() !== "POST") return false
  if (isDirectSendPath(path)) return true

  const segments = pathSegments(path)
  if (segments.at(-1) !== "$batch") return false
  return batchRequests(body).some(
    (request) =>
      typeof request.method === "string" &&
      typeof request.url === "string" &&
      isDirectMailSend(request.method, request.url),
  )
}

const DRAFT_REQUIRED_MESSAGE =
  "This server requires mail to go through a draft (MS365_REQUIRE_DRAFT), so graph_query can't send it " +
  "directly. Create a draft with create_draft, create_reply_draft, create_reply_all_draft or " +
  "create_forward_draft, then send it with send_draft."

export const graphQuery = async (
  params: {
    method: string
    path: string
    body?: string
    version?: string
    headers?: Record<string, string>
  },
  options: { readonly requireDraft?: boolean } = {},
): Promise<Either<UserError, string>> => {
  const body = params.body ? (JSON.parse(params.body) as Record<string, unknown>) : undefined
  if (options.requireDraft && isDirectMailSend(params.method, params.path, body)) {
    return Left(new UserError(DRAFT_REQUIRED_MESSAGE))
  }

  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const version = params.version as GraphApiVersion | undefined

  const result = await client.graphQuery(params.method, params.path, body, version, params.headers)
  return result
    .mapLeft((error) => new UserError(`Graph query failed: ${error.message}`))
    .map((data) => JSON.stringify(data, null, 2))
}
