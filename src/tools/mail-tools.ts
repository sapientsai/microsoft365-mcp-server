import { UserError } from "fastmcp"
import type { Either } from "functype/either"
import { Left } from "functype/either"

import { getGraphClient } from "../client/graph-client"
import type { ODataResponse } from "../types"
import { formatMessageDetail, formatMessageList } from "../utils/formatters"

const requireClient = () => {
  const client = getGraphClient()
  if (client.isNone()) return null
  return client.orThrow()
}

export const listMessages = async (params: { top?: number; filter?: string }): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

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

  const result = await client.sendMessage({
    message: {
      subject: params.subject,
      body: { contentType: params.content_type ?? "Text", content: params.body },
      toRecipients: [{ emailAddress: { address: params.to } }],
    },
  })
  return result
    .mapLeft((error) => new UserError(`Failed to send message: ${error.message}`))
    .map(() => `Message sent to ${params.to}.`)
}

export const replyToMessage = async (params: {
  message_id: string
  comment: string
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.replyToMessage(params.message_id, params.comment)
  return result
    .mapLeft((error) => new UserError(`Failed to reply: ${error.message}`))
    .map(() => "Reply sent successfully.")
}

export const searchMessages = async (params: { query: string; top?: number }): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.searchMessages(params.query, { $top: params.top ?? 25 })
  return result
    .mapLeft((error) => new UserError(`Failed to search messages: ${error.message}`))
    .map((response) => formatMessageList((response as ODataResponse<never>).value))
}
