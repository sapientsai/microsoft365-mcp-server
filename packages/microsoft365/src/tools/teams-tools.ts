import { UserError } from "fastmcp"
import type { Either } from "functype/either"
import { Left } from "functype/either"

import { getGraphClient } from "../client/graph-client"
import type { GraphChannel, GraphChannelMessage, ODataResponse } from "../types"
import { formatChannelList, formatChannelMessageList, formatTeamList } from "../utils/formatters"

const requireClient = () => {
  const client = getGraphClient()
  if (client.isNone()) return null
  return client.orThrow()
}

export const listTeams = async (params?: { fetch_all_pages?: boolean }): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  if (params?.fetch_all_pages) {
    const result = await client.requestPaginated<{ id: string; displayName?: string; description?: string }>(
      "/me/joinedTeams",
    )
    return result
      .mapLeft((error) => new UserError(`Failed to list teams: ${error.message}`))
      .map((items) => formatTeamList(items))
  }

  const result = await client.listTeams()
  return result
    .mapLeft((error) => new UserError(`Failed to list teams: ${error.message}`))
    .map((response) => formatTeamList((response as ODataResponse<never>).value))
}

export const listChannels = async (params: {
  team_id: string
  fetch_all_pages?: boolean
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  if (params.fetch_all_pages) {
    const result = await client.requestPaginated<GraphChannel>(`/teams/${params.team_id}/channels`)
    return result
      .mapLeft((error) => new UserError(`Failed to list channels: ${error.message}`))
      .map((items) => formatChannelList(items))
  }

  const result = await client.listChannels(params.team_id)
  return result
    .mapLeft((error) => new UserError(`Failed to list channels: ${error.message}`))
    .map((response) => formatChannelList((response as ODataResponse<never>).value))
}

export const listChannelMessages = async (params: {
  team_id: string
  channel_id: string
  top?: number
  fetch_all_pages?: boolean
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  if (params.fetch_all_pages) {
    const result = await client.requestPaginated<GraphChannelMessage>(
      `/teams/${params.team_id}/channels/${params.channel_id}/messages`,
    )
    return result
      .mapLeft((error) => new UserError(`Failed to list channel messages: ${error.message}`))
      .map((items) => formatChannelMessageList(items))
  }

  const result = await client.listChannelMessages(params.team_id, params.channel_id, { $top: params.top ?? 25 })
  return result
    .mapLeft((error) => new UserError(`Failed to list channel messages: ${error.message}`))
    .map((response) => formatChannelMessageList((response as ODataResponse<never>).value))
}

export const sendChannelMessage = async (params: {
  team_id: string
  channel_id: string
  content: string
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.sendChannelMessage(params.team_id, params.channel_id, params.content)
  return result
    .mapLeft((error) => new UserError(`Failed to send message: ${error.message}`))
    .map(() => "Message sent to channel.")
}
