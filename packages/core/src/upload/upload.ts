import type { Either } from "functype/either"
import { Left, Right } from "functype/either"

import { graphErrorMessage } from "../graph-error-message"
import { notSentError } from "../request-check"
import { MAY_HAVE_LANDED, mayHaveLandedStatus, resilientFetch } from "../resilience"
import type { GraphApiError, GraphDriveItem } from "../types"
import { graphPathProblem } from "../utils/graph-path"
import { conflictBehaviorProblem, contentTypeProblem, ifMatchProblem } from "../utils/header-values"

export const SIMPLE_UPLOAD_LIMIT = 4 * 1024 * 1024 // 4 MB
export const MAX_UPLOAD_SIZE = 250 * 1024 * 1024 // 250 MB
const CHUNK_SIZE = 10 * 1024 * 1024 // 10 MB (must be multiple of 320 KiB)

// undici's fetch wraps the real failure as `error.cause` and reports a generic
// "fetch failed" message. Surface the cause (and its `code`, e.g. ENOTFOUND /
// ECONNREFUSED / UND_ERR_*) so a network/egress failure names itself instead of
// hiding behind "fetch failed".
export const describeFetchError = (error: unknown): { message: string; code?: string } => {
  if (!(error instanceof Error)) return { message: String(error) }
  const { cause } = error as { cause?: unknown }
  if (cause instanceof Error) {
    const { code } = cause as { code?: unknown }
    const codeStr = typeof code === "string" ? code : undefined
    return {
      message: `${error.message} (cause: ${cause.message}${codeStr ? `, code: ${codeStr}` : ""})`,
      ...(codeStr ? { code: codeStr } : {}),
    }
  }
  if (cause !== undefined) return { message: `${error.message} (cause: ${String(cause)})` }
  return { message: error.message }
}

const parseGraphError = async (response: Response): Promise<string> =>
  graphErrorMessage(response.status, response.statusText, await response.json().catch(() => undefined))

// A conditional drive write (If-Match) answers 412 when the file changed since the caller read
// it. That is a lost race, not a failure to retry: only the caller knows how to merge, so it
// gets a distinct error that says what to do.
export const PRECONDITION_FAILED_MESSAGE =
  "precondition-failed: the file changed since it was read, so its eTag no longer matches. Read it again, re-apply your change, and upload with the new eTag."

const preconditionFailed = (): GraphApiError => ({ type: "api", status: 412, message: PRECONDITION_FAILED_MESSAGE })

/**
 * A conditional write replaces the file its eTag names, so it can't also rename or fail on a
 * conflict. Returns the refusal text for any other explicit conflict behavior, or undefined.
 */
export const ifMatchConflictProblem = (conflictBehavior: string | undefined, ifMatch: string | undefined) =>
  ifMatch !== undefined && conflictBehavior !== undefined && conflictBehavior !== "replace"
    ? `if_match replaces the file whose eTag it names, so conflict behavior must be "replace" or left unset, not "${conflictBehavior}".`
    : undefined

// Every upload call is a write. One that failed on a 503, 504 or network error wasn't retried,
// because it may have run (a renamed file created, a chunk accepted), so its message says so.
const toGraphApiError = async (response: Response, prefix: string): Promise<GraphApiError> => {
  const message = await parseGraphError(response)
  const note = mayHaveLandedStatus(response.status) ? ` ${MAY_HAVE_LANDED}` : ""
  return {
    type: "api",
    message: `${prefix}: ${message}${note}`,
    status: response.status,
  }
}

export const decodeBase64Upload = (rawBuffer: Buffer): Buffer =>
  Buffer.from(rawBuffer.toString("utf-8").replace(/\s/g, ""), "base64")

export const simpleUpload = async (
  apiBase: string,
  path: string,
  accessToken: string,
  buffer: Buffer,
  contentType: string,
  conflictBehavior: string,
): Promise<Either<GraphApiError, GraphDriveItem>> => {
  const separator = path.includes("?") ? "&" : "?"
  const url = `${apiBase}${path}${separator}@microsoft.graph.conflictBehavior=${conflictBehavior}`

  try {
    const response = await resilientFetch(url, {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": contentType,
        "Content-Length": String(buffer.length),
      },
      body: new Uint8Array(buffer),
    })

    if (!response.ok) {
      return Left(await toGraphApiError(response, "Upload failed"))
    }

    const item = (await response.json()) as GraphDriveItem
    return Right(item)
  } catch (error) {
    const notSent = notSentError(error)
    if (notSent) return Left(notSent)
    const { message } = describeFetchError(error)
    console.error(`[Upload] simple PUT to Graph threw: ${message}`)
    return Left<GraphApiError, GraphDriveItem>({
      type: "network",
      message: `Network error during upload: ${message} ${MAY_HAVE_LANDED}`,
    })
  }
}

// `ifMatch` goes on the createUploadSession POST, which is where Graph documents it: a mismatch
// answers 412 before any bytes are sent.
export const sessionUpload = async (
  apiBase: string,
  path: string,
  accessToken: string,
  buffer: Buffer,
  conflictBehavior: string,
  ifMatch?: string,
): Promise<Either<GraphApiError, GraphDriveItem>> => {
  // A drive path names the file as ".../root:/a/b.txt:/content", an item ID as ".../items/{id}/content".
  // Each has its own createUploadSession form; anything else has none.
  const sessionPath = /:\/?content$/i.test(path)
    ? path.replace(/:\/?content$/i, ":/createUploadSession")
    : path.replace(/\/content$/i, "/createUploadSession")
  if (sessionPath === path) {
    return Left({
      type: "api",
      status: 400,
      message: 'An upload path must end with ":/content" (a drive path) or "/content" (an item ID).',
    })
  }
  const sessionUrl = `${apiBase}${sessionPath}`

  try {
    const createResponse = await resilientFetch(sessionUrl, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
        ...(ifMatch !== undefined ? { "If-Match": ifMatch } : {}),
      },
      body: JSON.stringify({
        item: { "@microsoft.graph.conflictBehavior": conflictBehavior },
      }),
    })

    if (createResponse.status === 412) return Left(preconditionFailed())
    if (!createResponse.ok) {
      return Left(await toGraphApiError(createResponse, "Failed to create upload session"))
    }

    const session = (await createResponse.json()) as { uploadUrl: string }
    return uploadChunks(session.uploadUrl, buffer, 0)
  } catch (error) {
    const notSent = notSentError(error)
    if (notSent) return Left(notSent)
    const { message } = describeFetchError(error)
    console.error(`[Upload] createUploadSession to Graph threw: ${message}`)
    return Left<GraphApiError, GraphDriveItem>({
      type: "network",
      message: `Network error creating upload session: ${message} ${MAY_HAVE_LANDED}`,
    })
  }
}

const uploadChunks = async (
  uploadUrl: string,
  buffer: Buffer,
  offset: number,
): Promise<Either<GraphApiError, GraphDriveItem>> => {
  const totalSize = buffer.length
  if (offset >= totalSize) {
    return Left<GraphApiError, GraphDriveItem>({
      type: "api",
      message: "Upload completed but no DriveItem response received",
    })
  }

  const end = Math.min(offset + CHUNK_SIZE, totalSize)
  const chunk = buffer.subarray(offset, end)

  try {
    const chunkResponse = await resilientFetch(uploadUrl, {
      method: "PUT",
      headers: {
        "Content-Length": String(chunk.length),
        "Content-Range": `bytes ${offset}-${end - 1}/${totalSize}`,
      },
      body: new Uint8Array(chunk),
    })

    if (!chunkResponse.ok) {
      await resilientFetch(uploadUrl, { method: "DELETE" }).catch(() => {})
      // Graph's docs don't say whether the condition is checked again when the last chunk
      // commits. If it is, a 412 here is the same lost race and gets the same answer.
      if (chunkResponse.status === 412) return Left(preconditionFailed())
      return Left(await toGraphApiError(chunkResponse, `Upload chunk failed at byte ${offset}`))
    }

    if (chunkResponse.status === 200 || chunkResponse.status === 201) {
      const item = (await chunkResponse.json()) as GraphDriveItem
      return Right(item)
    }

    return uploadChunks(uploadUrl, buffer, offset + CHUNK_SIZE)
  } catch (error) {
    const notSent = notSentError(error)
    if (notSent) return Left(notSent)
    const { message } = describeFetchError(error)
    console.error(`[Upload] chunk PUT at byte ${offset} threw: ${message}`)
    return Left<GraphApiError, GraphDriveItem>({
      type: "network",
      message: `Network error uploading chunk at byte ${offset}: ${message} ${MAY_HAVE_LANDED}`,
    })
  }
}

/**
 * Why an upload path can't be written to, or undefined. Content goes only to a drive item's
 * content: a drive path ending ":/content" or an item path ending "/content". Anything else would
 * PUT the bytes to whatever Graph resource the path names. A query string is refused because the
 * upload adds its own.
 */
export const uploadPathProblem = (path: string): string | undefined => {
  const problem = graphPathProblem(path)
  if (problem !== undefined) return problem
  if (path.includes("?")) return "An upload path can't hold a query string."
  if (!/\/content$/i.test(path)) {
    return 'An upload path must end with ":/content" (a drive path) or "/content" (an item ID).'
  }
  return undefined
}

export type DriveUploadOptions = {
  /** The eTag (or cTag) the caller read. The write lands only if the file still has it. */
  readonly ifMatch?: string
}

/**
 * Writes file content to a drive path ending in ":/content". Up to 4 MB goes as one PUT, larger
 * files through an upload session. A conditional write always takes the session route, because
 * Graph documents If-Match on createUploadSession and not on the simple PUT; it also implies
 * "replace", since it names the file it overwrites.
 */
export const uploadDriveContent = async (
  apiBase: string,
  path: string,
  accessToken: string,
  buffer: Buffer,
  contentType: string,
  conflictBehavior: string,
  options: DriveUploadOptions = {},
): Promise<Either<GraphApiError, GraphDriveItem>> => {
  const { ifMatch } = options
  // Checked here because every upload entry point (three tools and both relays) comes through.
  // A blank eTag would go out as an empty If-Match header, which Graph may simply ignore, so the
  // write would run unconditionally while the caller believes it is guarded.
  const problem =
    uploadPathProblem(path) ??
    ifMatchProblem(ifMatch) ??
    contentTypeProblem(contentType) ??
    conflictBehaviorProblem(conflictBehavior)
  if (problem !== undefined) return Left({ type: "api", status: 400, message: problem })
  if (ifMatch === undefined) {
    return buffer.length <= SIMPLE_UPLOAD_LIMIT
      ? simpleUpload(apiBase, path, accessToken, buffer, contentType, conflictBehavior)
      : sessionUpload(apiBase, path, accessToken, buffer, conflictBehavior)
  }
  // An upload session has no chunk to send for empty content, so it could never finish.
  if (buffer.length === 0) {
    return Left({
      type: "api",
      status: 400,
      message: "if_match needs content: an empty file can't be written conditionally.",
    })
  }
  return sessionUpload(apiBase, path, accessToken, buffer, "replace", ifMatch)
}
