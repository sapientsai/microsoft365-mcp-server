import { UserError } from "fastmcp"
import type { Either } from "functype/either"
import { Left } from "functype/either"

import { getGraphClient } from "../client/graph-client"
import type { ODataResponse } from "../types"
import { formatUserDetail, formatUserList } from "../utils/formatters"

const requireClient = () => {
  const client = getGraphClient()
  if (client.isNone()) return null
  return client.orThrow()
}

export const getMe = async (): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.getMe()
  return result.mapLeft((error) => new UserError(`Failed to get profile: ${error.message}`)).map(formatUserDetail)
}

export const listUsers = async (params: { top?: number; filter?: string }): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.listUsers({ $top: params.top ?? 25, $filter: params.filter })
  return result
    .mapLeft((error) => new UserError(`Failed to list users: ${error.message}`))
    .map((response) => formatUserList((response as ODataResponse<never>).value))
}

export const getUser = async (params: { user_id: string }): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.getUser(params.user_id)
  return result.mapLeft((error) => new UserError(`Failed to get user: ${error.message}`)).map(formatUserDetail)
}
