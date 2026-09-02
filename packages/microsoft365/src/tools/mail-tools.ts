import { UserError } from "fastmcp"
import type { Either } from "functype/either"
import { Left, Right } from "functype/either"

import { getGraphClient } from "../client/graph-client"
import type { GraphAttachment, GraphMailFolder, GraphMessage, ODataResponse } from "../types"
import {
  formatAttachmentList,
  formatMailFolderList,
  formatMessageDetail,
  formatMessageList,
  formatMessageScan,
} from "../utils/formatters"
import { rememberMessageId, resolveMessageIdOrRef } from "../utils/message-refs"

// scan_messages hands back short refs instead of 152-character Graph IDs. Every tool
// that takes a message_id should accept either, otherwise the scan-then-act loop
// breaks at whichever tool was overlooked — which is what happened with
// list_attachments, the tool an attachment sweep depends on most.
const resolveMessageId = (idOrRef: string): Either<UserError, string> => {
  const resolved = resolveMessageIdOrRef(idOrRef)
  return resolved
    ? Right(resolved)
    : Left(
        new UserError(
          `Unknown message ref "${idOrRef}". Refs come from scan_messages and last for the session — re-run the scan to refresh them.`,
        ),
      )
}

const requireClient = () => {
  const client = getGraphClient()
  if (client.isNone()) return null
  return client.orThrow()
}

export const listMessages = async (params: {
  top?: number
  filter?: string
  fetch_all_pages?: boolean
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  if (params.fetch_all_pages) {
    const result = await client.requestPaginated<GraphMessage>("/me/messages", {
      odataParams: { $filter: params.filter, $orderby: "receivedDateTime desc" },
    })
    return result
      .mapLeft((error) => new UserError(`Failed to list messages: ${error.message}`))
      .map((items) => formatMessageList(items))
  }

  const result = await client.listMessages({
    $top: params.top ?? 25,
    $filter: params.filter,
    $orderby: "receivedDateTime desc",
  })
  return result
    .mapLeft((error) => new UserError(`Failed to list messages: ${error.message}`))
    .map((response) => formatMessageList((response as ODataResponse<never>).value))
}

export const getMessage = async (params: {
  message_id: string
  body_format?: "text" | "html"
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const messageId = resolveMessageId(params.message_id)
  if (messageId.isLeft()) return messageId as Either<UserError, string>

  const result = await client.getMessage(messageId.orThrow(), params.body_format)
  return result.mapLeft((error) => new UserError(`Failed to get message: ${error.message}`)).map(formatMessageDetail)
}

export const listMailFolders = async (params?: { fetch_all_pages?: boolean }): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  if (params?.fetch_all_pages) {
    const result = await client.requestPaginated<GraphMailFolder>("/me/mailFolders")
    return result
      .mapLeft((error) => new UserError(`Failed to list mail folders: ${error.message}`))
      .map((items) => formatMailFolderList(items))
  }

  const result = await client.listMailFolders({ $top: 100 })
  return result
    .mapLeft((error) => new UserError(`Failed to list mail folders: ${error.message}`))
    .map((response) => formatMailFolderList((response as ODataResponse<never>).value))
}

// Graph accepts these well-known names directly as a destinationId, so a caller can
// say "archive" without first resolving an opaque folder ID.
const WELL_KNOWN_FOLDERS: ReadonlyMap<string, string> = new Map([
  ["archive", "archive"],
  ["deleteditems", "deleteditems"],
  ["deleted items", "deleteditems"],
  ["trash", "deleteditems"],
  ["bin", "deleteditems"],
  ["inbox", "inbox"],
  ["junkemail", "junkemail"],
  ["junk", "junkemail"],
  ["drafts", "drafts"],
  ["sentitems", "sentitems"],
  ["sent items", "sentitems"],
])

const resolveDestination = async (
  client: NonNullable<ReturnType<typeof requireClient>>,
  destination: string,
): Promise<Either<UserError, string>> => {
  const normalized = destination.trim().toLowerCase()
  const wellKnown = WELL_KNOWN_FOLDERS.get(normalized)
  if (wellKnown) return Right(wellKnown)

  // Otherwise treat it as a folder display name and look it up.
  const result = await client.listMailFolders({ $top: 100 })
  return result
    .mapLeft((error) => new UserError(`Failed to resolve destination folder: ${error.message}`))
    .flatMap((response) => {
      const folders = (response as ODataResponse<GraphMailFolder>).value
      const matches = folders.filter((f) => f.displayName?.toLowerCase() === normalized)
      if (matches.length === 1) return Right(matches[0]!.id)
      if (matches.length > 1)
        return Left(
          new UserError(
            `Multiple folders named "${destination}". Pass the folder ID instead: ${matches.map((f) => f.id).join(", ")}`,
          ),
        )
      // No name matched — assume the caller passed a real folder ID and let Graph judge.
      return Right(destination)
    })
}

export const moveMessage = async (params: {
  message_id: string
  destination: string
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const messageId = resolveMessageId(params.message_id)
  if (messageId.isLeft()) return messageId as Either<UserError, string>

  const destination = await resolveDestination(client, params.destination)
  if (destination.isLeft()) return destination

  const result = await client.moveMessage(messageId.orThrow(), destination.orThrow())
  // Deliberately terse: triage moves messages in batches, and echoing each message body
  // back (formatMessageDetail) floods an LLM caller's context with mail the caller has
  // already decided to file. Subject and destination are enough to confirm the move.
  return result
    .mapLeft((error) => new UserError(`Failed to move message: ${error.message}`))
    .map((msg) => `Moved "${msg.subject ?? "(No Subject)"}" to ${params.destination}. New ID: ${msg.id}`)
}

// Graph has no bulk-move endpoint, so this is still one request per message — but it
// resolves the destination once instead of per message, and returns a single summary
// rather than N tool results. Filing an inbox means dozens of moves; at one call each
// the round-trips and the echoed confirmations dominate.
type MoveOutcome = { readonly id: string; readonly subject?: string; readonly error?: string }

const BATCH_MOVE_LIMIT = 50

export const batchMoveMessages = async (params: {
  message_ids: ReadonlyArray<string>
  destination: string
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  if (params.message_ids.length === 0) return Left(new UserError("At least one message ID is required."))
  if (params.message_ids.length > BATCH_MOVE_LIMIT) {
    return Left(
      new UserError(
        `Too many messages: ${params.message_ids.length}. Move at most ${BATCH_MOVE_LIMIT} at a time so a partial failure stays legible.`,
      ),
    )
  }

  const destination = await resolveDestination(client, params.destination)
  if (destination.isLeft()) return destination
  const destinationId = destination.orThrow()

  // Sequential on purpose: Graph throttles per-mailbox, and a 429 midway through a
  // parallel batch leaves the caller unsure which moves actually landed. Reducing over
  // a promise chain keeps that ordering without an imperative loop.
  const moveOne = async (idOrRef: string): Promise<MoveOutcome> => {
    const resolved = resolveMessageIdOrRef(idOrRef)
    // An unresolvable ref fails as its own outcome rather than aborting the batch:
    // filing dozens of messages should not be lost to one stale ref.
    if (!resolved) return { id: idOrRef, error: "Unknown message ref — re-run scan_messages to refresh" }

    const result = await client.moveMessage(resolved, destinationId)
    return result.fold<MoveOutcome>(
      (error) => ({ id: idOrRef, error: (error as { message: string }).message }),
      (msg) => ({ id: idOrRef, subject: (msg as GraphMessage).subject }),
    )
  }

  const outcomes = await params.message_ids.reduce<Promise<ReadonlyArray<MoveOutcome>>>(
    async (acc, id) => [...(await acc), await moveOne(id)],
    Promise.resolve([]),
  )

  const moved = outcomes.filter((o) => !o.error)
  const failed = outcomes.filter((o) => o.error)

  // Report failures individually — a silent partial success is the worst outcome here,
  // since the caller believes the inbox is filed when some of it is not.
  const failureLines = failed.map((f) => `- FAILED ${f.id}: ${f.error}`).join("\n")
  const summary = `Moved ${moved.length}/${outcomes.length} message(s) to ${params.destination}.`

  return failed.length === 0 ? Right(summary) : Right(`${summary}\n\n${failed.length} failed:\n${failureLines}`)
}

export const listAttachments = async (params: { message_id: string }): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const messageId = resolveMessageId(params.message_id)
  if (messageId.isLeft()) return messageId as Either<UserError, string>
  const id = messageId.orThrow()

  const result = await client.listAttachments(id)
  return result
    .mapLeft((error) => new UserError(`Failed to list attachments: ${error.message}`))
    .map((response) => formatAttachmentList(id, (response as ODataResponse<GraphAttachment>).value))
}

export const sendMessage = async (params: {
  to: string
  subject: string
  body: string
  content_type?: string
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const toRecipients = parseRecipients(params.to)
  if (!toRecipients) return Left(new UserError("At least one recipient is required in the 'to' field."))

  const result = await client.sendMessage({
    message: {
      subject: params.subject,
      body: { contentType: params.content_type ?? "Text", content: params.body },
      toRecipients,
    },
  })
  return result
    .mapLeft((error) => new UserError(`Failed to send message: ${error.message}`))
    .map(() => `Message sent to ${params.to}.`)
}

export const sendReply = async (params: {
  message_id: string
  comment: string
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.sendReply(params.message_id, params.comment)
  return result
    .mapLeft((error) => new UserError(`Failed to reply: ${error.message}`))
    .map(() => "Reply sent successfully.")
}

export const sendReplyAll = async (params: {
  message_id: string
  comment: string
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.sendReplyAll(params.message_id, params.comment)
  return result
    .mapLeft((error) => new UserError(`Failed to reply-all: ${error.message}`))
    .map(() => "Reply-all sent successfully.")
}

export const sendForward = async (params: {
  message_id: string
  to: string
  comment?: string
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const toRecipients = parseRecipients(params.to)
  if (!toRecipients) return Left(new UserError("At least one recipient is required in the 'to' field."))

  const result = await client.sendForward(params.message_id, params.comment ?? "", toRecipients)
  return result
    .mapLeft((error) => new UserError(`Failed to forward: ${error.message}`))
    .map(() => `Message forwarded to ${params.to}.`)
}

export const createReplyDraft = async (params: {
  message_id: string
  comment: string
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.createReplyDraft(params.message_id, params.comment)
  return result
    .mapLeft((error) => new UserError(`Failed to create reply draft: ${error.message}`))
    .map(
      (msg) =>
        `Reply draft created (original quoted, threaded). ID: ${(msg as { id: string }).id}. Review in Drafts, then send with send_draft.`,
    )
}

export const createReplyAllDraft = async (params: {
  message_id: string
  comment: string
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.createReplyAllDraft(params.message_id, params.comment)
  return result
    .mapLeft((error) => new UserError(`Failed to create reply-all draft: ${error.message}`))
    .map(
      (msg) =>
        `Reply-all draft created (original quoted, threaded). ID: ${(msg as { id: string }).id}. Review in Drafts, then send with send_draft.`,
    )
}

export const createForwardDraft = async (params: {
  message_id: string
  to: string
  comment?: string
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const toRecipients = parseRecipients(params.to)
  if (!toRecipients) return Left(new UserError("At least one recipient is required in the 'to' field."))

  const result = await client.createForwardDraft(params.message_id, params.comment ?? "", toRecipients)
  return result
    .mapLeft((error) => new UserError(`Failed to create forward draft: ${error.message}`))
    .map(
      (msg) =>
        `Forward draft created (original quoted). ID: ${(msg as { id: string }).id}. Review in Drafts, then send with send_draft.`,
    )
}

const parseRecipients = (
  value: string | undefined,
): ReadonlyArray<{ emailAddress: { address: string } }> | undefined => {
  if (!value) return undefined
  const addresses = value
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean)
  if (addresses.length === 0) return undefined
  return addresses.map((address) => ({ emailAddress: { address } }))
}

export const createDraft = async (params: {
  to: string
  subject: string
  body: string
  content_type?: string
  cc?: string
  bcc?: string
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const toRecipients = parseRecipients(params.to)
  if (!toRecipients) return Left(new UserError("At least one recipient is required in the 'to' field."))

  const message: Record<string, unknown> = {
    subject: params.subject,
    body: { contentType: params.content_type ?? "Text", content: params.body },
    toRecipients,
  }

  const cc = parseRecipients(params.cc)
  if (cc) message.ccRecipients = cc

  const bcc = parseRecipients(params.bcc)
  if (bcc) message.bccRecipients = bcc

  const result = await client.createDraft(message)
  return result
    .mapLeft((error) => new UserError(`Failed to create draft: ${error.message}`))
    .map((msg) => `Draft created. ID: ${(msg as { id: string }).id}`)
}

export const sendDraft = async (params: { message_id: string }): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.sendDraft(params.message_id)
  return result
    .mapLeft((error) => new UserError(`Failed to send draft: ${error.message}`))
    .map(() => "Draft sent successfully.")
}

export const searchMessages = async (params: { query: string; top?: number }): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.searchMessages(params.query, { $top: params.top ?? 25 })
  return result
    .mapLeft((error) => new UserError(`Failed to search messages: ${error.message}`))
    .map((response) => formatMessageList((response as ODataResponse<never>).value))
}

// Only the fields the scan actually prints. Graph returns the full message
// otherwise — including bodyPreview, which alone can be several hundred characters
// per message and is the single biggest waste when scanning thousands of headers.
const SCAN_FIELDS = ["id", "subject", "from", "receivedDateTime", "isRead", "hasAttachments"] as const

// Graph's own ceiling for $top on messages.
const MAX_PAGE = 999

export const scanMessages = async (params: {
  folder?: string
  filter?: string
  search?: string
  top?: number
  skip?: number
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const top = Math.min(params.top ?? 100, MAX_PAGE)

  // Graph ignores $skip when $search is set — it does not error, it silently returns
  // the first page again. A caller paging a search would therefore re-read the same
  // rows while believing it was advancing, and conclude it had seen everything.
  // Refusing the combination is the only way to make that visible.
  if (params.search && params.skip !== undefined) {
    return Left(
      new UserError(
        "skip cannot be combined with search: Graph ignores $skip on a $search query and would silently return the first page again. " +
          "Page a search by narrowing it instead — add a received range to the search string " +
          '(e.g. "invoice AND received:2024-01-01..2024-06-30") and walk the windows.',
      ),
    )
  }

  // Ask for one extra row: if it comes back, there is a further page, and the caller
  // learns that without paying for a separate $count request.
  const odataParams = {
    $select: [...SCAN_FIELDS],
    $filter: params.filter,
    $search: params.search,
    $top: top + 1,
    $skip: params.skip,
    // $search and $orderby are mutually exclusive in Graph — asking for both is a
    // 400, so relevance ordering wins whenever a search term is present.
    $orderby: params.search ? undefined : "receivedDateTime desc",
  }

  const resolved = params.folder ? await resolveDestination(client, params.folder) : undefined
  if (resolved?.isLeft()) return resolved as Either<UserError, string>
  const folderId = resolved?.orThrow()

  const result = folderId
    ? await client.listFolderMessages(folderId, odataParams)
    : await client.listMessages(odataParams)

  return result
    .mapLeft((error) => new UserError(`Failed to scan messages: ${error.message}`))
    .map((response) => {
      const all = (response as ODataResponse<GraphMessage>).value
      const hasMore = all.length > top
      const page = hasMore ? all.slice(0, top) : all
      const refs = page.map((msg) => rememberMessageId(msg.id))

      return formatMessageScan(page, refs, {
        folder: params.folder,
        hasMore,
        // A search cannot be paged with skip (see above), so the caller is told to
        // narrow instead. Only a filter/list scan gets a usable next offset.
        nextSkip: params.search ? undefined : (params.skip ?? 0) + top,
        searched: params.search !== undefined,
      })
    })
}
