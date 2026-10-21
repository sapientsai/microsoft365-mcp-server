import type { Either } from "functype/either"
import { Left, Right } from "functype/either"

import { resilientFetch } from "../resilience"
import type { GraphApiError } from "../types"
import { describeFetchError } from "./upload"

// Outlook attachment limits (https://learn.microsoft.com/graph/outlook-large-attachments):
// under 3 MB goes inline as one POST, and createUploadSession refuses anything smaller;
// 3 MB to 150 MB goes through an upload session in ranges kept under 4 MB.
export const ATTACHMENT_INLINE_LIMIT = 3 * 1024 * 1024
export const MAX_ATTACHMENT_SIZE = 150 * 1024 * 1024
export const ATTACHMENT_CHUNK_SIZE = 12 * 320 * 1024 // 3.75 MiB

// The session uploadUrl carries its own authtoken, so the range PUTs send no Authorization header.
// Only follow it to Outlook's hosts: a forged session response must not receive the file.
const OUTLOOK_UPLOAD_HOSTS: ReadonlySet<string> = new Set(["outlook.office.com", "outlook.office365.com"])

export type MessageAttachment = { readonly id?: string; readonly name: string; readonly size: number }

const graphError = async (response: Response, prefix: string): Promise<GraphApiError> => {
  const error = await response
    .json()
    .then((body: { error?: { code?: string; message?: string } }) => body.error)
    .catch(() => undefined)
  return {
    type: "api",
    message: `${prefix}: ${error?.message ?? `${response.status} ${response.statusText}`}`,
    status: response.status,
    ...(error?.code ? { graphErrorCode: error.code } : {}),
  }
}

// Microsoft never says whether its "3 MB" is 3,000,000 bytes or 3 MiB. A file near the line may be
// too big to POST inline (413) or too small for a session (this code); either refusal attaches
// nothing, so the other route is safe to try once.
const SESSION_TOO_SMALL = "ErrorAttachmentSizeShouldNotBeLessThanMinimumSize"
const inlineTooLarge = (error: GraphApiError): boolean => error.status === 413
const sessionTooSmall = (error: GraphApiError): boolean => error.graphErrorCode === SESSION_TOO_SMALL

const networkError = (what: string, error: unknown): GraphApiError => {
  const { message } = describeFetchError(error)
  console.error(`[AttachmentUpload] ${what} threw: ${message}`)
  return { type: "network", message: `Network error ${what}: ${message}` }
}

const messagePath = (apiBase: string, messageId: string): string =>
  `${apiBase}/me/messages/${encodeURIComponent(messageId)}`

// Graph lets any message take attachments, so the draft rule is ours. The relay URL carries the
// message id in its query string, and a caller could swap it for a received or sent message.
export const requireDraftMessage = async (
  apiBase: string,
  messageId: string,
  accessToken: string,
): Promise<Either<GraphApiError, void>> => {
  try {
    const response = await resilientFetch(`${messagePath(apiBase, messageId)}?$select=id,isDraft`, {
      headers: { Authorization: `Bearer ${accessToken}` },
    })
    if (!response.ok) return Left(await graphError(response, "Failed to read message"))
    const message = (await response.json()) as { isDraft?: boolean }
    if (message.isDraft !== true) {
      return Left({
        type: "api",
        message: "Only drafts can take attachments; this message is not a draft.",
        status: 409,
      })
    }
    return Right(undefined)
  } catch (error) {
    return Left(networkError("reading the message", error))
  }
}

const attachInline = async (
  apiBase: string,
  messageId: string,
  accessToken: string,
  buffer: Buffer,
  name: string,
  contentType: string,
): Promise<Either<GraphApiError, MessageAttachment>> => {
  try {
    // POST is not retried after a lost response, so a timeout cannot attach the file twice.
    const response = await resilientFetch(`${messagePath(apiBase, messageId)}/attachments`, {
      method: "POST",
      headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
      body: JSON.stringify({
        "@odata.type": "#microsoft.graph.fileAttachment",
        name,
        contentType,
        contentBytes: buffer.toString("base64"),
      }),
    })
    if (!response.ok) return Left(await graphError(response, "Failed to attach file"))
    const created = (await response.json()) as { id?: string; size?: number }
    return Right({ id: created.id, name, size: created.size ?? buffer.length })
  } catch (error) {
    return Left(networkError("attaching the file", error))
  }
}

const outlookUploadUrl = (uploadUrl: string | undefined): Either<GraphApiError, string> => {
  const url = uploadUrl && URL.canParse(uploadUrl) ? new URL(uploadUrl) : undefined
  if (url?.protocol !== "https:" || !OUTLOOK_UPLOAD_HOSTS.has(url.hostname.toLowerCase())) {
    return Left({
      type: "api",
      message: `Attachment upload session returned an unexpected host: ${url?.host ?? "(no valid URL)"}`,
    })
  }
  return Right(url.toString())
}

// The final range answers 201 with a Location header and an empty body; earlier ranges answer 200.
const attachmentIdFrom = (location: string | null): string | undefined =>
  location ? /Attachments\('([^']+)'\)/i.exec(location)?.[1] : undefined

const cancelSession = (uploadUrl: string): Promise<void> =>
  resilientFetch(uploadUrl, { method: "DELETE" }).then(
    () => undefined,
    () => undefined,
  )

const uploadRanges = async (
  uploadUrl: string,
  buffer: Buffer,
  name: string,
  offset: number,
): Promise<Either<GraphApiError, MessageAttachment>> => {
  const total = buffer.length
  const end = Math.min(offset + ATTACHMENT_CHUNK_SIZE, total)

  try {
    // An accepted range advances the session; replaying it after a lost response is unsafe.
    const response = await resilientFetch(
      uploadUrl,
      {
        method: "PUT",
        headers: {
          "Content-Type": "application/octet-stream",
          "Content-Length": String(end - offset),
          "Content-Range": `bytes ${offset}-${end - 1}/${total}`,
        },
        body: new Uint8Array(buffer.subarray(offset, end)),
      },
      { retryTransient: false },
    )

    if (!response.ok) {
      await cancelSession(uploadUrl)
      return Left(await graphError(response, `Attachment upload failed at byte ${offset}`))
    }
    // Outlook documents 201 for the last range. Any 2xx there still means the file is in, so a 200
    // must not cancel the session and report a failure that would prompt a duplicate retry.
    if (response.status === 201 || end >= total) {
      return Right({ id: attachmentIdFrom(response.headers.get("location")), name, size: total })
    }
    return uploadRanges(uploadUrl, buffer, name, end)
  } catch (error) {
    await cancelSession(uploadUrl)
    return Left(networkError(`uploading to ${new URL(uploadUrl).host} at byte ${offset}`, error))
  }
}

const attachViaSession = async (
  apiBase: string,
  messageId: string,
  accessToken: string,
  buffer: Buffer,
  name: string,
  contentType: string,
): Promise<Either<GraphApiError, MessageAttachment>> => {
  try {
    const response = await resilientFetch(`${messagePath(apiBase, messageId)}/attachments/createUploadSession`, {
      method: "POST",
      headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
      body: JSON.stringify({ AttachmentItem: { attachmentType: "file", name, size: buffer.length, contentType } }),
    })
    if (!response.ok) return Left(await graphError(response, "Failed to create attachment upload session"))
    const session = (await response.json()) as { uploadUrl?: string }
    const uploadUrl = outlookUploadUrl(session.uploadUrl)
    if (uploadUrl.isLeft()) return Left(uploadUrl.value as GraphApiError)
    return await uploadRanges(uploadUrl.value as string, buffer, name, 0)
  } catch (error) {
    return Left(networkError("creating the attachment upload session", error))
  }
}

/**
 * Attaches a file to a message the caller has already confirmed is a draft (see requireDraftMessage):
 * inline under 3 MB, through an Outlook upload session up to 150 MB.
 */
export const attachFile = async (
  apiBase: string,
  messageId: string,
  accessToken: string,
  buffer: Buffer,
  name: string,
  contentType: string,
): Promise<Either<GraphApiError, MessageAttachment>> => {
  if (buffer.length > MAX_ATTACHMENT_SIZE) {
    return Left({
      type: "api",
      message: `File too large for a mail attachment (max ${MAX_ATTACHMENT_SIZE} bytes)`,
      status: 413,
    })
  }
  type Attached = Promise<Either<GraphApiError, MessageAttachment>>
  const args = [apiBase, messageId, accessToken, buffer, name, contentType] as const
  if (buffer.length < ATTACHMENT_INLINE_LIMIT) {
    const inline = await attachInline(...args)
    return inline.fold<Attached>(
      (error) => (inlineTooLarge(error) ? attachViaSession(...args) : Promise.resolve(inline)),
      () => Promise.resolve(inline),
    )
  }
  const session = await attachViaSession(...args)
  return session.fold<Attached>(
    (error) => (sessionTooSmall(error) ? attachInline(...args) : Promise.resolve(session)),
    () => Promise.resolve(session),
  )
}

/** Confirms the message is a draft, then attaches the file. */
export const attachFileToDraft = async (
  apiBase: string,
  messageId: string,
  accessToken: string,
  buffer: Buffer,
  name: string,
  contentType: string,
): Promise<Either<GraphApiError, MessageAttachment>> => {
  const draft = await requireDraftMessage(apiBase, messageId, accessToken)
  if (draft.isLeft()) return Left(draft.value as GraphApiError)
  return attachFile(apiBase, messageId, accessToken, buffer, name, contentType)
}
