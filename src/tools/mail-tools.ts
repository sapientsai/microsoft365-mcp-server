import { UserError } from "fastmcp"
import type { Either } from "functype/either"
import { Left } from "functype/either"

import { getGraphClient } from "../client/graph-client"
import type { GraphMessage, ODataResponse } from "../types"
import { formatMessageDetail, formatMessageList } from "../utils/formatters"

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

export const getMessage = async (params: { message_id: string }): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.getMessage(params.message_id)
  return result.mapLeft((error) => new UserError(`Failed to get message: ${error.message}`)).map(formatMessageDetail)
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
