import { UserError } from "fastmcp"
import type { Either } from "functype/either"
import { Left } from "functype/either"

import { getGraphClient } from "../client/graph-client"
import type { GraphChat, GraphChatMessage, ODataResponse } from "../types"
import { formatChatList, formatChatMessageList } from "../utils/formatters"

const requireClient = () => {
  const client = getGraphClient()
  if (client.isNone()) return null
  return client.orThrow()
}

export const listChats = async (params?: {
  top?: number
  fetch_all_pages?: boolean
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  if (params?.fetch_all_pages) {
    const result = await client.requestPaginated<GraphChat>("/me/chats")
    return result
      .mapLeft((error) => new UserError(`Failed to list chats: ${error.message}`))
      .map((items) => formatChatList(items))
  }

  const result = await client.listChats({ $top: params?.top ?? 25 })
  return result
    .mapLeft((error) => new UserError(`Failed to list chats: ${error.message}`))
    .map((response) => formatChatList((response as ODataResponse<never>).value))
}

export const listChatMessages = async (params: {
  chat_id: string
  top?: number
  fetch_all_pages?: boolean
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  if (params.fetch_all_pages) {
    const result = await client.requestPaginated<GraphChatMessage>(`/chats/${params.chat_id}/messages`)
    return result
      .mapLeft((error) => new UserError(`Failed to list chat messages: ${error.message}`))
      .map((items) => formatChatMessageList(items))
  }

  const result = await client.listChatMessages(params.chat_id, { $top: params.top ?? 25 })
  return result
    .mapLeft((error) => new UserError(`Failed to list chat messages: ${error.message}`))
    .map((response) => formatChatMessageList((response as ODataResponse<never>).value))
}

export const sendChatMessage = async (params: {
  chat_id: string
  content: string
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.sendChatMessage(params.chat_id, params.content)
  return result
    .mapLeft((error) => new UserError(`Failed to send chat message: ${error.message}`))
    .map(() => "Chat message sent.")
}
