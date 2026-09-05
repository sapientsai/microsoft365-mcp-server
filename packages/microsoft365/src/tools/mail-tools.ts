import { UserError } from "fastmcp"
import type { Either } from "functype/either"
import { Left, Right } from "functype/either"

import { getGraphClient } from "../client/graph-client"
import type { GraphAttachment, GraphMailFolder, GraphMessage, ODataResponse } from "../types"
import { formatAttachmentList, formatMailFolderList, formatMessageDetail, formatMessageList } from "../utils/formatters"

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

  const result = await client.getMessage(params.message_id, params.body_format)
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

// What the caller typed is not what the message ends up in. "junk" is a well-known alias AND a
// legal display name for a custom folder, and the alias wins — so a mailbox with a folder named
// "Junk" files the message into Junk Email instead, which is a different folder. Carrying a label
// alongside the id lets the confirmation say which branch actually fired, rather than echoing the
// input back and leaving the caller to assume.
//
// assumedId records that no name matched and we handed the caller's string to Graph as an ID. It
// only changes the message on failure, so it costs nothing and guesses nothing: a typo'd folder
// name and a genuine folder ID are indistinguishable up front, but once Graph has rejected it we
// know which explanation to give.
type ResolvedFolder = { readonly id: string; readonly label: string; readonly assumedId: boolean }

const resolveDestination = async (
  client: NonNullable<ReturnType<typeof requireClient>>,
  destination: string,
): Promise<Either<UserError, ResolvedFolder>> => {
  const normalized = destination.trim().toLowerCase()
  const wellKnown = WELL_KNOWN_FOLDERS.get(normalized)
  if (wellKnown) return Right({ id: wellKnown, label: `the ${wellKnown} folder`, assumedId: false })

  // Otherwise treat it as a folder display name and look it up.
  const result = await client.listMailFolders({ $top: 100 })
  return result
    .mapLeft((error) => new UserError(`Failed to resolve destination folder: ${error.message}`))
    .flatMap((response): Either<UserError, ResolvedFolder> => {
      const folders = (response as ODataResponse<GraphMailFolder>).value
      const matches = folders.filter((f) => f.displayName?.toLowerCase() === normalized)
      if (matches.length > 1)
        return Left(
          new UserError(
            `Multiple folders named "${destination}". Pass the folder ID instead: ${matches.map((f) => f.id).join(", ")}`,
          ),
        )
      if (matches.length === 1) {
        const [match] = matches
        return Right({ id: match.id, label: `"${match.displayName}"`, assumedId: false })
      }
      // No name matched — assume the caller passed a real folder ID and let Graph judge. Note that
      // listMailFolders only sees top-level folders, so a subfolder never matches by name and
      // always lands here; passing its ID is the supported route.
      return Right({ id: destination, label: `folder ID ${destination}`, assumedId: true })
    })
}

export const moveMessage = async (params: {
  message_id: string
  destination: string
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const destination = await resolveDestination(client, params.destination)
  if (destination.isLeft()) return Left(destination.value as UserError)
  const target = destination.orThrow()

  const result = await client.moveMessage(params.message_id, target.id)
  // Deliberately terse: triage moves messages in batches, and echoing each message body
  // back (formatMessageDetail) floods an LLM caller's context with mail the caller has
  // already decided to file. Subject and destination are enough to confirm the move.
  //
  // The label, not params.destination: the caller needs to see where the message actually
  // went when the two differ.
  return result
    .mapLeft((error) =>
      target.assumedId
        ? new UserError(
            `No top-level folder is named "${params.destination}", and Graph rejected it as a folder ID: ` +
              `${error.message}. Check list_mail_folders for the name, or pass a subfolder's ID.`,
          )
        : new UserError(`Failed to move message: ${error.message}`),
    )
    .map((msg) => `Moved "${msg.subject ?? "(No Subject)"}" to ${target.label}. New ID: ${msg.id}`)
}

export const listAttachments = async (params: { message_id: string }): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.listAttachments(params.message_id)
  return result
    .mapLeft((error) => new UserError(`Failed to list attachments: ${error.message}`))
    .map((response) => formatAttachmentList(params.message_id, (response as ODataResponse<GraphAttachment>).value))
}

// Graph has no bulk-move endpoint, so this is still one request per message — but it
// resolves the destination once instead of per message, and returns a single summary
// rather than N tool results. Filing an inbox means dozens of moves; at one call each
// the round-trips and the echoed confirmations dominate.
type MoveOutcome = {
  readonly id: string
  readonly subject?: string
  readonly error?: string
  // Graph throttles per mailbox, so a 429 on one message predicts a 429 on the next.
  readonly throttled?: boolean
}

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
  if (destination.isLeft()) return Left(destination.value as UserError)
  const target = destination.orThrow()

  // Sequential on purpose: Graph throttles per-mailbox, and a 429 midway through a
  // parallel batch leaves the caller unsure which moves actually landed. Reducing over
  // a promise chain keeps that ordering without an imperative loop.
  const moveOne = async (id: string): Promise<MoveOutcome> => {
    const result = await client.moveMessage(id, target.id)
    return result.fold<MoveOutcome>(
      (error) => ({ id, error: error.message, throttled: error.type === "throttle" }),
      (msg) => ({ id, subject: msg.subject }),
    )
  }

  const outcomes = await params.message_ids.reduce<Promise<ReadonlyArray<MoveOutcome>>>(async (acc, id) => {
    const done = await acc
    // Stop at the first throttle. Graph throttles per mailbox, so message N+1 is throttled too:
    // carrying on spends the rest of the batch on calls that cannot succeed and buries the one
    // real cause under 40-odd identical failures. Say what was skipped rather than pretending it
    // was tried.
    if (done.some((o) => o.throttled))
      return [...done, { id, error: "not attempted — the batch stopped after Graph throttled it" }]
    return [...done, await moveOne(id)]
  }, Promise.resolve([]))

  const moved = outcomes.filter((o) => !o.error)
  const failed = outcomes.filter((o) => o.error)

  // Report failures individually — a silent partial success is the worst outcome here,
  // since the caller believes the inbox is filed when some of it is not.
  const failureLines = failed.map((f) => `- FAILED ${f.id}: ${f.error}`).join("\n")
  const summary = `Moved ${moved.length}/${outcomes.length} message(s) to ${target.label}.`
  const detail = `${summary}\n\n${failed.length} failed:\n${failureLines}`

  if (failed.length === 0) return Right(summary)
  // A batch where nothing moved is a failure, not a success carrying bad news. Returning Right
  // leaves MCP's isError unset, so the caller sees a success-shaped result with the failures buried
  // in the text — and an LLM triaging a mailbox reports it as filed when none of it was.
  // A partial success stays Right: some messages really did move, and the caller needs that list.
  return moved.length === 0 ? Left(new UserError(detail)) : Right(detail)
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
