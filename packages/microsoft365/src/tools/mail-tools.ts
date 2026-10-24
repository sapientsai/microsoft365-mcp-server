import { randomUUID } from "node:crypto"

import { formatBytes, MAX_ATTACHMENT_SIZE, resolveUploadContentType } from "@sapientsai/ms-graph-core"
import { UserError } from "fastmcp"
import type { Either } from "functype/either"
import { Left, Right } from "functype/either"

import { getGraphClient } from "../client/graph-client"
import { type MailboxScope, OWN_MAILBOX, resolveMailboxScope } from "../mail/mailbox"
import type { GraphAttachment, GraphMailFolder, GraphMessage, ODataResponse } from "../types"
import {
  formatAttachmentList,
  formatMailFolderList,
  formatMessageDetail,
  formatMessageList,
  formatMessageScan,
  MESSAGE_SUMMARY_FIELDS,
} from "../utils/formatters"
import { describeRefFailure, rememberMessageId, resolveMessageIdOrRef } from "../utils/message-refs"
import { relayAuthHeader, relayCurl, resolveUploadBaseUrl } from "./files-tools"

// scan_messages hands back short refs instead of 152-character Graph IDs. Every tool
// used in the scan-then-act loop resolves refs before calling Graph — against the mailbox
// being addressed, since a ref minted in one mailbox means nothing in another.
const resolveMessageId = (idOrRef: string, scope: MailboxScope): Either<UserError, string> => {
  const resolved = resolveMessageIdOrRef(idOrRef, scope.mailbox)
  return resolved.kind === "id" ? Right(resolved.id) : Left(new UserError(describeRefFailure(resolved, scope.mailbox)))
}

const requireClient = () => {
  const client = getGraphClient()
  if (client.isNone()) return null
  return client.orThrow()
}

// Every mail tool starts the same way: resolve the mailbox, then get the client. Both
// can fail with a UserError, and neither is worth repeating twenty times.
type ScopedClient = { scope: MailboxScope; client: NonNullable<ReturnType<typeof requireClient>> }

const withScope = (mailbox: string | undefined): Either<UserError, ScopedClient> => {
  const scope = resolveMailboxScope(mailbox)
  if (scope.isLeft()) return Left(scope.value as UserError)

  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  return Right({ scope: scope.orThrow(), client })
}

// A failed scope carries a UserError and no Right value, so it cannot simply be cast
// to the handler's return type — re-wrap the error instead.
const scopeFailure = (resolved: Either<UserError, ScopedClient>): Either<UserError, string> =>
  Left(resolved.value as UserError)

// list_messages always sorts by receivedDateTime, and Graph rejects a $filter that does not lead
// with the sort property: "InefficientFilter — The restriction or sort order is too complex for
// this operation" (List messages docs, "Using filter and orderby in the same query"; confirmed live
// 2026-10-04 with "importance eq 'high'"). Putting a condition every message satisfies first lets any
// caller filter through while results stay newest first, which dropping the sort would not.
const ALWAYS_TRUE_DATE_CONDITION = "receivedDateTime ge 1900-01-01T00:00:00Z"
const LEADS_WITH_SORT_FIELD = /^\s*\(*\s*receivedDateTime\b/i

export const orderableFilter = (filter?: string): string | undefined => {
  if (filter === undefined || filter.trim() === "") return undefined
  return LEADS_WITH_SORT_FIELD.test(filter) ? filter : `${ALWAYS_TRUE_DATE_CONDITION} and (${filter})`
}

export const listMessages = async (params: {
  top?: number
  filter?: string
  folder?: string
  include_preview?: boolean
  fetch_all_pages?: boolean
  mailbox?: string
}): Promise<Either<UserError, string>> => {
  const resolvedScope = withScope(params.mailbox)
  if (resolvedScope.isLeft()) return scopeFailure(resolvedScope)
  const { scope, client } = resolvedScope.orThrow()

  // $select keeps full message bodies out of the response: without it Graph returns every body in
  // HTML, which the listing never prints.
  const odataParams = {
    $select: MESSAGE_SUMMARY_FIELDS,
    $filter: orderableFilter(params.filter),
    $orderby: "receivedDateTime desc",
  }
  const options = { preview: params.include_preview ?? false }

  // Same resolution as move_message, so both tools accept the same folder names.
  const resolved = params.folder === undefined ? undefined : await resolveDestination(client, params.folder, scope)
  if (resolved?.isLeft()) return Left(resolved.value as UserError)
  const target = resolved?.orThrow()

  const toUserError = (error: { message: string }) =>
    target?.assumedId
      ? new UserError(
          `No top-level folder is named "${params.folder}", and Graph rejected it as a folder ID: ` +
            `${error.message}. Check list_mail_folders for the name, or pass a subfolder's ID.`,
        )
      : new UserError(`Failed to list messages: ${error.message}`)

  if (params.fetch_all_pages) {
    const result = await client.listAllMessages(odataParams, target?.id, scope.prefix)
    return result.mapLeft(toUserError).map((items) => formatMessageList(items as ReadonlyArray<GraphMessage>, options))
  }

  const result = await client.listMessages({ ...odataParams, $top: params.top ?? 25 }, target?.id, scope.prefix)
  return result
    .mapLeft(toUserError)
    .map((response) => formatMessageList((response as ODataResponse<GraphMessage>).value, options))
}

export const getMessage = async (params: {
  message_id: string
  body_format?: "text" | "html"
  mailbox?: string
}): Promise<Either<UserError, string>> => {
  const resolvedScope = withScope(params.mailbox)
  if (resolvedScope.isLeft()) return scopeFailure(resolvedScope)
  const { scope, client } = resolvedScope.orThrow()

  const messageId = resolveMessageId(params.message_id, scope)
  if (messageId.isLeft()) return messageId as Either<UserError, string>

  const result = await client.getMessage(messageId.orThrow(), params.body_format, scope.prefix)
  return result.mapLeft((error) => new UserError(`Failed to get message: ${error.message}`)).map(formatMessageDetail)
}

export const listMailFolders = async (params?: {
  fetch_all_pages?: boolean
  mailbox?: string
}): Promise<Either<UserError, string>> => {
  const resolvedScope = withScope(params?.mailbox)
  if (resolvedScope.isLeft()) return scopeFailure(resolvedScope)
  const { scope, client } = resolvedScope.orThrow()

  if (params?.fetch_all_pages) {
    const result = await client.requestPaginated<GraphMailFolder>(`${scope.prefix}/mailFolders`)
    return result
      .mapLeft((error) => new UserError(`Failed to list mail folders: ${error.message}`))
      .map((items) => formatMailFolderList(items))
  }

  const result = await client.listMailFolders({ $top: 100 }, scope.prefix)
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
  scope: MailboxScope,
): Promise<Either<UserError, ResolvedFolder>> => {
  const normalized = destination.trim().toLowerCase()
  const wellKnown = WELL_KNOWN_FOLDERS.get(normalized)
  if (wellKnown) return Right({ id: wellKnown, label: `the ${wellKnown} folder`, assumedId: false })

  // Otherwise treat it as a folder display name and look it up — in the mailbox being
  // addressed, not the signed-in user's. Folder IDs are per-mailbox, so resolving a
  // name against the wrong one yields an ID that is missing (or, worse, valid) there.
  const result = await client.listMailFolders({ $top: 100 }, scope.prefix)
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
  mailbox?: string
}): Promise<Either<UserError, string>> => {
  const resolvedScope = withScope(params.mailbox)
  if (resolvedScope.isLeft()) return scopeFailure(resolvedScope)
  const { scope, client } = resolvedScope.orThrow()

  const messageId = resolveMessageId(params.message_id, scope)
  if (messageId.isLeft()) return messageId as Either<UserError, string>

  const destination = await resolveDestination(client, params.destination, scope)
  if (destination.isLeft()) return Left(destination.value as UserError)
  const target = destination.orThrow()

  const result = await client.moveMessage(messageId.orThrow(), target.id, scope.prefix)
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

export const listAttachments = async (params: {
  message_id: string
  mailbox?: string
}): Promise<Either<UserError, string>> => {
  const resolvedScope = withScope(params.mailbox)
  if (resolvedScope.isLeft()) return scopeFailure(resolvedScope)
  const { scope, client } = resolvedScope.orThrow()

  const messageId = resolveMessageId(params.message_id, scope)
  if (messageId.isLeft()) return messageId as Either<UserError, string>
  const id = messageId.orThrow()
  const result = await client.listAttachments(id, scope.prefix)
  return result
    .mapLeft((error) => new UserError(`Failed to list attachments: ${error.message}`))
    .map((response) => formatAttachmentList(id, (response as ODataResponse<GraphAttachment>).value))
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
  // Never sent, because an earlier message was throttled. Distinct from a failure: nothing was
  // attempted, so retrying it is the obvious next move and counting it as "failed" would overstate
  // the damage.
  readonly skipped?: boolean
}

const BATCH_MOVE_LIMIT = 50

export const batchMoveMessages = async (params: {
  message_ids: ReadonlyArray<string>
  destination: string
  mailbox?: string
}): Promise<Either<UserError, string>> => {
  const resolvedScope = withScope(params.mailbox)
  if (resolvedScope.isLeft()) return scopeFailure(resolvedScope)
  const { scope, client } = resolvedScope.orThrow()

  if (params.message_ids.length === 0) return Left(new UserError("At least one message ID is required."))
  if (params.message_ids.length > BATCH_MOVE_LIMIT) {
    return Left(
      new UserError(
        `Too many messages: ${params.message_ids.length}. Move at most ${BATCH_MOVE_LIMIT} at a time so a partial failure stays legible.`,
      ),
    )
  }

  const destination = await resolveDestination(client, params.destination, scope)
  if (destination.isLeft()) return Left(destination.value as UserError)
  const target = destination.orThrow()

  // Sequential on purpose: Graph throttles per-mailbox, and a 429 midway through a
  // parallel batch leaves the caller unsure which moves actually landed. Reducing over
  // a promise chain keeps that ordering without an imperative loop.
  const moveOne = async (id: string): Promise<MoveOutcome> => {
    const resolved = resolveMessageIdOrRef(id, scope.mailbox)
    // An unresolvable ref fails as its own outcome rather than aborting the batch:
    // filing dozens of messages should not be lost to one stale ref.
    if (resolved.kind !== "id") return { id, error: describeRefFailure(resolved, scope.mailbox) }
    const result = await client.moveMessage(resolved.id, target.id, scope.prefix)
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
      return [...done, { id, error: "the batch stopped after Graph throttled it", skipped: true }]
    return [...done, await moveOne(id)]
  }, Promise.resolve([]))

  const moved = outcomes.filter((o) => o.error === undefined)
  const failed = outcomes.filter((o) => o.error !== undefined && o.skipped !== true)
  const skipped = outcomes.filter((o) => o.skipped === true)

  // Report failures individually — a silent partial success is the worst outcome here,
  // since the caller believes the inbox is filed when some of it is not. Skipped messages are
  // listed apart from failures: they were never sent, so they are still safe to retry, and folding
  // them into the failure count would report more damage than actually happened.
  const failureLines = failed.map((f) => `- FAILED ${f.id}: ${f.error}`).join("\n")
  const skippedLines = skipped.map((sk) => `- NOT ATTEMPTED ${sk.id}`).join("\n")
  const summary = `Moved ${moved.length}/${outcomes.length} message(s) to ${target.label}.`
  // A destination that only Graph could judge is worth naming again here: the whole batch went to
  // the same place, so if it was wrong, it was wrong for every message.
  const hint = target.assumedId
    ? `\n\nNo top-level folder is named "${params.destination}"; it was used as a folder ID.`
    : ""
  const detail = `${summary}${hint}${
    failed.length > 0 ? `\n\n${failed.length} failed:\n${failureLines}` : ""
  }${skipped.length > 0 ? `\n\n${skipped.length} not attempted:\n${skippedLines}` : ""}`

  if (failed.length === 0 && skipped.length === 0) return Right(summary)
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
  mailbox?: string
}): Promise<Either<UserError, string>> => {
  const resolvedScope = withScope(params.mailbox)
  if (resolvedScope.isLeft()) return scopeFailure(resolvedScope)
  const { scope, client } = resolvedScope.orThrow()

  const toRecipients = parseRecipients(params.to)
  if (!toRecipients) return Left(new UserError("At least one recipient is required in the 'to' field."))

  const result = await client.sendMessage(
    {
      message: {
        subject: params.subject,
        body: { contentType: params.content_type ?? "Text", content: params.body },
        toRecipients,
      },
    },
    scope.prefix,
  )
  return result
    .mapLeft((error) => new UserError(`Failed to send message: ${error.message}`))
    .map(() => `Message sent to ${params.to}.`)
}

export const sendReply = async (params: {
  message_id: string
  comment: string
  mailbox?: string
}): Promise<Either<UserError, string>> => {
  const resolvedScope = withScope(params.mailbox)
  if (resolvedScope.isLeft()) return scopeFailure(resolvedScope)
  const { scope, client } = resolvedScope.orThrow()

  const result = await client.sendReply(params.message_id, params.comment, scope.prefix)
  return result
    .mapLeft((error) => new UserError(`Failed to reply: ${error.message}`))
    .map(() => "Reply sent successfully.")
}

export const sendReplyAll = async (params: {
  message_id: string
  comment: string
  mailbox?: string
}): Promise<Either<UserError, string>> => {
  const resolvedScope = withScope(params.mailbox)
  if (resolvedScope.isLeft()) return scopeFailure(resolvedScope)
  const { scope, client } = resolvedScope.orThrow()

  const result = await client.sendReplyAll(params.message_id, params.comment, scope.prefix)
  return result
    .mapLeft((error) => new UserError(`Failed to reply-all: ${error.message}`))
    .map(() => "Reply-all sent successfully.")
}

export const sendForward = async (params: {
  message_id: string
  to: string
  comment?: string
  mailbox?: string
}): Promise<Either<UserError, string>> => {
  const resolvedScope = withScope(params.mailbox)
  if (resolvedScope.isLeft()) return scopeFailure(resolvedScope)
  const { scope, client } = resolvedScope.orThrow()

  const toRecipients = parseRecipients(params.to)
  if (!toRecipients) return Left(new UserError("At least one recipient is required in the 'to' field."))

  const result = await client.sendForward(params.message_id, params.comment ?? "", toRecipients, scope.prefix)
  return result
    .mapLeft((error) => new UserError(`Failed to forward: ${error.message}`))
    .map(() => `Message forwarded to ${params.to}.`)
}

export const createReplyDraft = async (params: {
  message_id: string
  comment: string
  mailbox?: string
}): Promise<Either<UserError, string>> => {
  const resolvedScope = withScope(params.mailbox)
  if (resolvedScope.isLeft()) return scopeFailure(resolvedScope)
  const { scope, client } = resolvedScope.orThrow()

  const result = await client.createReplyDraft(params.message_id, params.comment, scope.prefix)
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
  mailbox?: string
}): Promise<Either<UserError, string>> => {
  const resolvedScope = withScope(params.mailbox)
  if (resolvedScope.isLeft()) return scopeFailure(resolvedScope)
  const { scope, client } = resolvedScope.orThrow()

  const result = await client.createReplyAllDraft(params.message_id, params.comment, scope.prefix)
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
  mailbox?: string
}): Promise<Either<UserError, string>> => {
  const resolvedScope = withScope(params.mailbox)
  if (resolvedScope.isLeft()) return scopeFailure(resolvedScope)
  const { scope, client } = resolvedScope.orThrow()

  const toRecipients = parseRecipients(params.to)
  if (!toRecipients) return Left(new UserError("At least one recipient is required in the 'to' field."))

  const result = await client.createForwardDraft(params.message_id, params.comment ?? "", toRecipients, scope.prefix)
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
  mailbox?: string
}): Promise<Either<UserError, string>> => {
  const resolvedScope = withScope(params.mailbox)
  if (resolvedScope.isLeft()) return scopeFailure(resolvedScope)
  const { scope, client } = resolvedScope.orThrow()

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

  const result = await client.createDraft(message, scope.prefix)
  return result
    .mapLeft((error) => new UserError(`Failed to create draft: ${error.message}`))
    .map((msg) => `Draft created. ID: ${(msg as { id: string }).id}`)
}

export const sendDraft = async (params: {
  message_id: string
  mailbox?: string
}): Promise<Either<UserError, string>> => {
  const resolvedScope = withScope(params.mailbox)
  if (resolvedScope.isLeft()) return scopeFailure(resolvedScope)
  const { scope, client } = resolvedScope.orThrow()

  const result = await client.sendDraft(params.message_id, scope.prefix)
  return result
    .mapLeft((error) => new UserError(`Failed to send draft: ${error.message}`))
    .map(() => "Draft sent successfully.")
}

// Graph lets subject, body and recipients change only while a message is a draft, but DELETE works
// on any message. Checking isDraft first keeps the draft tools from touching received or sent mail.
// `rule` names what the calling tool needs a draft for, so its refusal explains the right thing.
const fetchDraft = async (
  client: NonNullable<ReturnType<typeof requireClient>>,
  messageId: string,
  scope: MailboxScope,
  rule = "Only drafts can be edited or deleted",
): Promise<Either<UserError, GraphMessage>> => {
  const result = await client.getDraftState(messageId, scope.prefix)
  if (result.isLeft())
    return Left(new UserError(`Failed to read message: ${(result.value as { message: string }).message}`))
  const message = result.value as GraphMessage
  if (message.isDraft !== true) {
    return Left(
      new UserError(
        `Message ${messageId} is not a draft. ${rule}; find drafts with ` + "list_messages in the drafts folder.",
      ),
    )
  }
  return Right(message)
}

// Undefined leaves the list alone; an empty string clears it.
const replacementRecipients = (
  value: string | undefined,
): ReadonlyArray<{ emailAddress: { address: string } }> | undefined =>
  value === undefined ? undefined : (parseRecipients(value) ?? [])

const draftChanges = (params: {
  subject?: string
  body?: string
  content_type?: string
  to?: string
  cc?: string
  bcc?: string
}): Record<string, unknown> =>
  Object.fromEntries(
    [
      ["subject", params.subject],
      [
        "body",
        params.body === undefined ? undefined : { contentType: params.content_type ?? "Text", content: params.body },
      ],
      ["toRecipients", replacementRecipients(params.to)],
      ["ccRecipients", replacementRecipients(params.cc)],
      ["bccRecipients", replacementRecipients(params.bcc)],
    ].filter(([, value]) => value !== undefined),
  )

export const updateDraft = async (params: {
  message_id: string
  subject?: string
  body?: string
  content_type?: string
  to?: string
  cc?: string
  bcc?: string
  mailbox?: string
}): Promise<Either<UserError, string>> => {
  const resolvedScope = withScope(params.mailbox)
  if (resolvedScope.isLeft()) return scopeFailure(resolvedScope)
  const { scope, client } = resolvedScope.orThrow()

  const changes = draftChanges(params)
  if (Object.keys(changes).length === 0) {
    return Left(new UserError("Nothing to change. Pass at least one of subject, body, to, cc or bcc."))
  }

  const draft = await fetchDraft(client, params.message_id, scope)
  if (draft.isLeft()) return Left(draft.value as UserError)

  const result = await client.updateMessage(params.message_id, changes, scope.prefix)
  return result
    .mapLeft((error) => new UserError(`Failed to update draft: ${error.message}`))
    .map((message) => `Draft updated. Send it with send_draft when ready.\n\n${formatMessageDetail(message)}`)
}

export const deleteDraft = async (params: {
  message_id: string
  mailbox?: string
}): Promise<Either<UserError, string>> => {
  const resolvedScope = withScope(params.mailbox)
  if (resolvedScope.isLeft()) return scopeFailure(resolvedScope)
  const { scope, client } = resolvedScope.orThrow()

  const draft = await fetchDraft(client, params.message_id, scope)
  if (draft.isLeft()) return Left(draft.value as UserError)
  const subject = (draft.value as GraphMessage).subject ?? "(no subject)"

  const result = await client.deleteMessage(params.message_id, scope.prefix)
  return result
    .mapLeft((error) => new UserError(`Failed to delete draft: ${error.message}`))
    .map(() => `Draft "${subject}" deleted.`)
}

// A sandbox such as claude.ai's can reach this server but not Outlook, where a large attachment's
// upload session lives. The relay takes the bytes instead, and the server attaches them, so the
// file never passes through the model or needs a sandbox network rule.
export const getAttachmentUploadConfig = async (params: {
  message_id: string
  name: string
  localFile?: string
  contentType?: string
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const name = params.name.trim()
  if (!name) return Left(new UserError("name is required: the file name the attachment will show, e.g. report.pdf"))

  const messageId = resolveMessageId(params.message_id, OWN_MAILBOX)
  if (messageId.isLeft()) return messageId as Either<UserError, string>
  const id = messageId.orThrow()

  const draft = await fetchDraft(client, id, OWN_MAILBOX, "Attachments can only be added to drafts")
  if (draft.isLeft()) return Left(draft.value as UserError)

  const baseUrl = resolveUploadBaseUrl()
  if (!baseUrl) {
    return Left(
      new UserError(
        "Upload endpoint base URL not configured. Set MS365_PUBLIC_BASE_URL (or MS365_OAUTH_BASE_URL in oauth-proxy mode), or run with TRANSPORT_TYPE=httpStream.",
      ),
    )
  }

  const contentType = resolveUploadContentType(params.contentType, name)
  const query = new URLSearchParams({ messageId: id, name, contentType, encoding: "base64" })
  const uploadUrl = `${baseUrl}/upload/attachment?${query.toString()}`
  const authHeader = await relayAuthHeader()

  return Right(
    JSON.stringify(
      {
        uploadUrl,
        method: "POST",
        contentType,
        maxSize: formatBytes(MAX_ATTACHMENT_SIZE),
        encoding: "base64",
        ...(authHeader ? { authHeader } : {}),
        curl: relayCurl(params.localFile ?? "{local_file_path}", uploadUrl, authHeader),
        notes: [
          "Pipe base64-encoded file bytes to uploadUrl via POST with --data-binary @-.",
          "The server attaches the file to the draft itself: inline under 3 MB, through an Outlook upload session up to 150 MB. Only this server's host needs to be reachable.",
          "The Authorization value is an opaque, short-lived upload ticket (~10 min TTL) — not your Graph token.",
          "Each POST adds one attachment. If the curl fails or times out, run list_attachments on the draft before retrying: a large upload can finish after the client gives up, and a retry would attach the file twice.",
          "Send the draft with send_draft once the attachment is on it.",
        ],
      },
      null,
      2,
    ),
  )
}

export const searchMessages = async (params: {
  query: string
  top?: number
  mailbox?: string
}): Promise<Either<UserError, string>> => {
  const resolvedScope = withScope(params.mailbox)
  if (resolvedScope.isLeft()) return scopeFailure(resolvedScope)
  const { scope, client } = resolvedScope.orThrow()

  const result = await client.searchMessages(params.query, { $top: params.top ?? 25 }, scope.prefix)
  return result
    .mapLeft((error) => new UserError(`Failed to search messages: ${error.message}`))
    .map((response) => formatMessageList((response as ODataResponse<never>).value))
}

// Only the fields the scan actually prints. Graph returns the full message
// otherwise — including bodyPreview, which alone can be several hundred characters
// per message and is the single biggest waste when scanning thousands of headers.
const SCAN_FIELDS = ["id", "subject", "from", "receivedDateTime", "isRead", "hasAttachments"] as const

// Keep Graph continuation URLs server-side, including the original scan context.
// Cursors belong to the client that issued them and cannot become arbitrary URLs.
// They also remember the mailbox they page through: refs minted from a continued page
// are qualified by that mailbox, so continuing under a different one would mislabel them.
const scanCursors = new Map<
  string,
  {
    client: NonNullable<ReturnType<typeof requireClient>>
    mailbox?: string
    nextLink: string
    folder?: string
    searched: boolean
  }
>()

export const scanMessages = async (params: {
  folder?: string
  filter?: string
  search?: string
  top?: number
  cursor?: string
  mailbox?: string
}): Promise<Either<UserError, string>> => {
  const resolvedScope = withScope(params.mailbox)
  if (resolvedScope.isLeft()) return scopeFailure(resolvedScope)
  const { scope, client } = resolvedScope.orThrow()

  if (
    params.cursor !== undefined &&
    [params.folder, params.filter, params.search, params.top].some((value) => value !== undefined)
  ) {
    return Left(new UserError("Pass cursor alone to continue a scan; start a new scan to change its parameters."))
  }
  const continuation = params.cursor === undefined ? undefined : scanCursors.get(params.cursor)
  if (params.cursor !== undefined && continuation?.client !== client) {
    return Left(new UserError("Unknown or expired scan cursor. Re-run the original scan to start again."))
  }
  if (continuation && continuation.mailbox !== scope.mailbox) {
    return Left(
      new UserError(
        `This cursor pages through ${continuation.mailbox ?? "your own mailbox"}; pass the same mailbox to continue it.`,
      ),
    )
  }
  if (params.top !== undefined && (!Number.isInteger(params.top) || params.top < 1 || params.top > 1000)) {
    return Left(new UserError("top must be an integer between 1 and 1000."))
  }

  const resolved = !continuation && params.folder ? await resolveDestination(client, params.folder, scope) : undefined
  if (resolved?.isLeft()) return Left(resolved.value as UserError)
  const folder = continuation?.folder ?? params.folder
  const searched = continuation?.searched ?? params.search !== undefined
  const result = continuation
    ? await client.nextMessageScanPage(continuation.nextLink, scope.prefix)
    : await client.listMessages(
        {
          $select: [...SCAN_FIELDS],
          $filter: params.filter,
          $search: params.search,
          $top: params.top ?? 100,
          // Avoid InefficientFilter: arbitrary filters need not match a date ordering.
          $orderby: params.search || params.filter ? undefined : "receivedDateTime desc",
        },
        resolved?.orThrow().id,
        scope.prefix,
      )

  return result
    .mapLeft((error) => new UserError(`Failed to scan messages: ${error.message}`))
    .map((response) => {
      const page = response as ODataResponse<GraphMessage>
      const nextLink = page["@odata.nextLink"]
      const nextCursor = nextLink ? randomUUID() : undefined
      if (nextCursor && nextLink)
        scanCursors.set(nextCursor, { client, mailbox: scope.mailbox, nextLink, folder, searched })
      return formatMessageScan(
        page.value,
        page.value.map((msg) => rememberMessageId(msg.id, scope.mailbox)),
        {
          folder,
          nextCursor,
          searched,
          mailbox: scope.mailbox,
        },
      )
    })
}
