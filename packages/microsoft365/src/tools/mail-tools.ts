import { UserError } from "fastmcp"
import type { Either } from "functype/either"
import { Left, Right } from "functype/either"

import { getGraphClient } from "../client/graph-client"
import { runMoveBatches } from "../mail/batch-move"
import { type MailboxScope, resolveMailboxScope } from "../mail/mailbox"
import {
  formatSenderSummary,
  senderAddress,
  type SenderGroupBy,
  summarizeSenders as aggregateSenders,
} from "../mail/sender-summary"
import type { GraphAttachment, GraphBatchResponse, GraphMailFolder, GraphMessage, ODataResponse } from "../types"
import {
  formatAttachmentList,
  formatMailFolderList,
  formatMessageDetail,
  formatMessageList,
  formatMessageScan,
} from "../utils/formatters"
import { describeRefFailure, rememberMessageId, resolveMessageIdOrRef } from "../utils/message-refs"

const requireClient = () => {
  const client = getGraphClient()
  if (client.isNone()) return null
  return client.orThrow()
}

// scan_messages hands back short refs instead of 152-character Graph IDs. Every tool
// that takes a message_id should accept either, otherwise the scan-then-act loop
// breaks at whichever tool was overlooked — which is what happened with
// list_attachments, the tool an attachment sweep depends on most.
const resolveMessageId = (idOrRef: string, scope: MailboxScope): Either<UserError, string> => {
  const resolved = resolveMessageIdOrRef(idOrRef, scope.mailbox)
  return resolved.kind === "id" ? Right(resolved.id) : Left(new UserError(describeRefFailure(resolved, scope.mailbox)))
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

export const listMessages = async (params: {
  top?: number
  filter?: string
  fetch_all_pages?: boolean
  mailbox?: string
}): Promise<Either<UserError, string>> => {
  const resolved = withScope(params.mailbox)
  if (resolved.isLeft()) return scopeFailure(resolved)
  const { scope, client } = resolved.orThrow()

  if (params.fetch_all_pages) {
    const result = await client.requestPaginated<GraphMessage>(`${scope.prefix}/messages`, {
      odataParams: { $filter: params.filter, $orderby: "receivedDateTime desc" },
    })
    return result
      .mapLeft((error) => new UserError(`Failed to list messages: ${error.message}`))
      .map((items) => formatMessageList(items))
  }

  const result = await client.listMessages(
    {
      $top: params.top ?? 25,
      $filter: params.filter,
      $orderby: "receivedDateTime desc",
    },
    scope.prefix,
  )
  return result
    .mapLeft((error) => new UserError(`Failed to list messages: ${error.message}`))
    .map((response) => formatMessageList((response as ODataResponse<never>).value))
}

export const getMessage = async (params: {
  message_id: string
  body_format?: "text" | "html"
  mailbox?: string
}): Promise<Either<UserError, string>> => {
  const resolved = withScope(params.mailbox)
  if (resolved.isLeft()) return scopeFailure(resolved)
  const { scope, client } = resolved.orThrow()

  const messageId = resolveMessageId(params.message_id, scope)
  if (messageId.isLeft()) return messageId as Either<UserError, string>

  const result = await client.getMessage(messageId.orThrow(), params.body_format, scope.prefix)
  return result.mapLeft((error) => new UserError(`Failed to get message: ${error.message}`)).map(formatMessageDetail)
}

export const listMailFolders = async (params?: {
  fetch_all_pages?: boolean
  mailbox?: string
}): Promise<Either<UserError, string>> => {
  const resolved = withScope(params?.mailbox)
  if (resolved.isLeft()) return scopeFailure(resolved)
  const { scope, client } = resolved.orThrow()

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
        const match = matches[0]!
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
  const resolved = withScope(params.mailbox)
  if (resolved.isLeft()) return scopeFailure(resolved)
  const { scope, client } = resolved.orThrow()

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
  const moveOne = async (idOrRef: string): Promise<MoveOutcome> => {
    const resolved = resolveMessageIdOrRef(idOrRef, scope.mailbox)
    // An unresolvable ref fails as its own outcome rather than aborting the batch:
    // filing dozens of messages should not be lost to one stale ref.
    if (resolved.kind !== "id") return { id: idOrRef, error: describeRefFailure(resolved, scope.mailbox) }

    const result = await client.moveMessage(resolved.id, target.id, scope.prefix)
    return result.fold<MoveOutcome>(
      (error) => ({
        id: idOrRef,
        error: (error as { message: string }).message,
        throttled: (error as { type?: string }).type === "throttle",
      }),
      (msg) => ({ id: idOrRef, subject: (msg as GraphMessage).subject }),
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

export const listAttachments = async (params: {
  message_id: string
  mailbox?: string
}): Promise<Either<UserError, string>> => {
  const resolved = withScope(params.mailbox)
  if (resolved.isLeft()) return scopeFailure(resolved)
  const { scope, client } = resolved.orThrow()

  const messageId = resolveMessageId(params.message_id, scope)
  if (messageId.isLeft()) return messageId as Either<UserError, string>
  const id = messageId.orThrow()

  const result = await client.listAttachments(id, scope.prefix)
  return result
    .mapLeft((error) => new UserError(`Failed to list attachments: ${error.message}`))
    .map((response) => formatAttachmentList(id, (response as ODataResponse<GraphAttachment>).value))
}

export const sendMessage = async (params: {
  to: string
  subject: string
  body: string
  content_type?: string
  mailbox?: string
}): Promise<Either<UserError, string>> => {
  const resolved = withScope(params.mailbox)
  if (resolved.isLeft()) return scopeFailure(resolved)
  const { scope, client } = resolved.orThrow()

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
  const resolved = withScope(params.mailbox)
  if (resolved.isLeft()) return scopeFailure(resolved)
  const { scope, client } = resolved.orThrow()

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
  const resolved = withScope(params.mailbox)
  if (resolved.isLeft()) return scopeFailure(resolved)
  const { scope, client } = resolved.orThrow()

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
  const resolved = withScope(params.mailbox)
  if (resolved.isLeft()) return scopeFailure(resolved)
  const { scope, client } = resolved.orThrow()

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
  const resolved = withScope(params.mailbox)
  if (resolved.isLeft()) return scopeFailure(resolved)
  const { scope, client } = resolved.orThrow()

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
  const resolved = withScope(params.mailbox)
  if (resolved.isLeft()) return scopeFailure(resolved)
  const { scope, client } = resolved.orThrow()

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
  const resolved = withScope(params.mailbox)
  if (resolved.isLeft()) return scopeFailure(resolved)
  const { scope, client } = resolved.orThrow()

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

// Graph's own ceiling for $top on messages.
const MAX_PAGE = 999

// Graph's rule for combining $filter with $orderby on messages: every property in the
// $orderby must also appear in the $filter, and before any other property. A filter on
// from/emailAddress alone therefore fails with "The restriction or sort order is too
// complex for this operation". Prefixing an always-true receivedDateTime clause
// satisfies the rule without changing what matches.
const ORDERED_FILTER_PREFIX = "receivedDateTime ge 1900-01-01T00:00:00Z"
export const orderedFilter = (filter: string | undefined): string | undefined =>
  filter === undefined || filter.trim().length === 0
    ? undefined
    : /^\s*receivedDateTime\b/.test(filter)
      ? filter
      : `${ORDERED_FILTER_PREFIX} and (${filter})`

export const scanMessages = async (params: {
  folder?: string
  filter?: string
  search?: string
  top?: number
  skip?: number
  mailbox?: string
}): Promise<Either<UserError, string>> => {
  const resolvedScope = withScope(params.mailbox)
  if (resolvedScope.isLeft()) return scopeFailure(resolvedScope)
  const { scope, client } = resolvedScope.orThrow()

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
    // Only a sorted scan needs the prefix; a search has no $orderby to conflict with.
    $filter: params.search ? params.filter : orderedFilter(params.filter),
    $search: params.search,
    $top: top + 1,
    $skip: params.skip,
    // $search and $orderby are mutually exclusive in Graph — asking for both is a
    // 400, so relevance ordering wins whenever a search term is present.
    $orderby: params.search ? undefined : "receivedDateTime desc",
  }

  const resolvedFolder = params.folder ? await resolveDestination(client, params.folder, scope) : undefined
  if (resolvedFolder?.isLeft()) return Left(resolvedFolder.value as UserError)
  const folderId = resolvedFolder?.orThrow().id

  const result = folderId
    ? await client.listFolderMessages(folderId, odataParams, scope.prefix)
    : await client.listMessages(odataParams, scope.prefix)

  return result
    .mapLeft((error) => new UserError(`Failed to scan messages: ${error.message}`))
    .map((response) => {
      const all = (response as ODataResponse<GraphMessage>).value
      const hasMore = all.length > top
      const page = hasMore ? all.slice(0, top) : all
      const refs = page.map((msg) => rememberMessageId(msg.id, scope.mailbox))

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

// --- Folder sweeps -------------------------------------------------------------
//
// Cleaning a large inbox is two questions: "who sends the bulk of this?" and "move
// everything from them". Both need every matching header, not a page of them, and the
// second needs to act on thousands of ids without dragging them through the caller.
// So both fetch server-side, without $orderby (Graph refuses a from/emailAddress
// filter combined with a receivedDateTime sort), and only summaries cross the wire.

const SWEEP_FIELDS = ["id", "subject", "from", "receivedDateTime", "isRead"] as const
const SWEEP_PAGE = 999

// A single quote is the only character OData needs escaped inside a string literal.
const odataString = (value: string): string => `'${value.replace(/'/g, "''")}'`

const senderFilter = (senders: ReadonlyArray<string>): string =>
  senders.map((address) => `from/emailAddress/address eq ${odataString(address.trim().toLowerCase())}`).join(" or ")

const fetchFolderMessages = async (
  client: NonNullable<ReturnType<typeof requireClient>>,
  scope: MailboxScope,
  folder: string,
  filter: string | undefined,
): Promise<Either<UserError, { readonly folderId: string; readonly messages: ReadonlyArray<GraphMessage> }>> => {
  const resolvedFolder = await resolveDestination(client, folder, scope)
  if (resolvedFolder.isLeft()) return Left(resolvedFolder.value as UserError)
  const folderId = resolvedFolder.orThrow().id

  const result = await client.listFolderMessagesAll(
    folderId,
    { $select: [...SWEEP_FIELDS], $filter: filter, $top: SWEEP_PAGE },
    scope.prefix,
  )
  return result
    .mapLeft((error) => new UserError(`Failed to read ${folder}: ${error.message}`))
    .map((messages) => ({ folderId, messages }))
}

export const summarizeSenders = async (params: {
  folder?: string
  filter?: string
  group_by?: SenderGroupBy
  top?: number
  mailbox?: string
}): Promise<Either<UserError, string>> => {
  const resolvedScope = withScope(params.mailbox)
  if (resolvedScope.isLeft()) return scopeFailure(resolvedScope)
  const { scope, client } = resolvedScope.orThrow()

  const folder = params.folder ?? "inbox"
  const groupBy = params.group_by ?? "address"
  const top = Math.min(Math.max(params.top ?? 100, 1), 1000)

  const fetched = await fetchFolderMessages(client, scope, folder, params.filter)
  return fetched.map(({ messages }) =>
    formatSenderSummary(aggregateSenders(messages, groupBy), {
      folder,
      total: messages.length,
      unread: messages.filter((m) => m.isRead === false).length,
      groupBy,
      top,
      filter: params.filter,
    }),
  )
}

const DEFAULT_SWEEP_LIMIT = 1000
const MAX_SWEEP_LIMIT = 20000
const MAX_SENDERS_PER_SWEEP = 20

const formatSweepPreview = (
  messages: ReadonlyArray<GraphMessage>,
  meta: { readonly folder: string; readonly filter: string; readonly destination: string },
): string => {
  if (messages.length === 0) return `No messages in ${meta.folder} match: ${meta.filter}`

  const days = messages.map((m) => m.receivedDateTime?.slice(0, 10) ?? "").filter((d) => d !== "")
  const oldest = days.reduce((a, b) => (a < b ? a : b))
  const newest = days.reduce((a, b) => (a > b ? a : b))
  const unread = messages.filter((m) => m.isRead === false).length

  const senders = aggregateSenders(messages, "address").slice(0, 10)
  const senderLines = senders.map((s) => `- ${s.count} from ${s.key}${s.name ? ` (${s.name})` : ""}`)

  const newestTen = [...messages]
    .sort((a, b) => (b.receivedDateTime ?? "").localeCompare(a.receivedDateTime ?? ""))
    .slice(0, 10)
    .map(
      (m) =>
        `- ${m.receivedDateTime?.slice(0, 10) ?? ""} | ${senderAddress(m)} | ${(m.subject ?? "(No Subject)").replace(/[\r\n]+/g, " ").slice(0, 100)}`,
    )

  return [
    `# Sweep preview — ${messages.length} messages in ${meta.folder} match`,
    `filter: ${meta.filter}`,
    `Oldest ${oldest}, newest ${newest}, ${unread} unread.`,
    "",
    "By sender:",
    ...senderLines,
    "",
    "Newest 10:",
    ...newestTen,
    "",
    `**Nothing was moved.** Re-run with dry_run: false to move all ${messages.length} to ${meta.destination}.`,
  ].join("\n")
}

export const moveMessagesMatching = async (params: {
  folder: string
  destination: string
  filter?: string
  senders?: ReadonlyArray<string>
  dry_run?: boolean
  limit?: number
  mailbox?: string
}): Promise<Either<UserError, string>> => {
  const resolvedScope = withScope(params.mailbox)
  if (resolvedScope.isLeft()) return scopeFailure(resolvedScope)
  const { scope, client } = resolvedScope.orThrow()

  const senders = (params.senders ?? []).map((s) => s.trim()).filter((s) => s.length > 0)
  const filter = (params.filter ?? "").trim()
  if (senders.length === 0 && filter.length === 0) {
    return Left(
      new UserError(
        "Pass senders (a list of addresses) and/or filter (an OData expression). Sweeping a whole folder unconditionally is refused.",
      ),
    )
  }
  if (senders.length > MAX_SENDERS_PER_SWEEP) {
    return Left(new UserError(`At most ${MAX_SENDERS_PER_SWEEP} senders per call; split the list.`))
  }

  // Both given: the sender clause narrows the filter (and), never widens it.
  const combined =
    senders.length > 0 && filter.length > 0
      ? `(${senderFilter(senders)}) and (${filter})`
      : senders.length > 0
        ? senderFilter(senders)
        : filter

  const destination = await resolveDestination(client, params.destination, scope)
  if (destination.isLeft()) return Left(destination.value as UserError)
  const target = destination.orThrow()

  const fetched = await fetchFolderMessages(client, scope, params.folder, combined)
  if (fetched.isLeft()) return Left(fetched.value as UserError)
  const { folderId, messages } = fetched.orThrow()

  if (folderId === target.id) {
    return Left(new UserError(`Destination "${params.destination}" is the folder being swept.`))
  }

  const dryRun = params.dry_run ?? true
  const meta = { folder: params.folder, filter: combined, destination: params.destination }
  if (dryRun) return Right(formatSweepPreview(messages, meta))

  if (messages.length === 0) return Right(`No messages in ${params.folder} match: ${combined}`)

  const limit = Math.min(params.limit ?? DEFAULT_SWEEP_LIMIT, MAX_SWEEP_LIMIT)
  // Refusing, rather than moving the first N, keeps a mis-scoped filter from doing
  // partial damage: the caller sees the real count and decides.
  if (messages.length > limit) {
    return Left(
      new UserError(
        `${messages.length} messages match, above the limit of ${limit}. Narrow the filter, or pass limit: ${messages.length} (max ${MAX_SWEEP_LIMIT}) if that count is intended.`,
      ),
    )
  }

  const outcome = await runMoveBatches(
    messages.map((m) => m.id),
    scope.prefix,
    target.id,
    (requests) =>
      client.batchRequest(requests as ReadonlyArray<Record<string, unknown>>) as Promise<
        Either<never, GraphBatchResponse>
      >,
  )

  const summary = `Moved ${outcome.moved.length}/${messages.length} message(s) from ${params.folder} to ${params.destination}.`
  if (outcome.failed.length === 0) return Right(summary)

  // Failures are listed by subject rather than id: an id is useless to a caller who
  // wants to know what did not move, and the listing is capped to stay legible.
  const subjectById = new Map(messages.map((m) => [m.id, m.subject ?? "(No Subject)"]))
  const shown = outcome.failed.slice(0, 20)
  const failureLines = shown.map((f) => `- FAILED "${subjectById.get(f.id) ?? f.id}": ${f.error}`)
  const overflow = outcome.failed.length > shown.length ? `\n... and ${outcome.failed.length - shown.length} more` : ""
  return Right(`${summary}\n\n${outcome.failed.length} failed:\n${failureLines.join("\n")}${overflow}`)
}
