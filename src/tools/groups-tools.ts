import { UserError } from "fastmcp"
import type { Either } from "functype/either"
import { Left } from "functype/either"

import { getGraphClient } from "../client/graph-client"
import type { ODataResponse } from "../types"
import { formatGroupDetail, formatGroupList, formatUserList } from "../utils/formatters"

const requireClient = () => {
  const client = getGraphClient()
  if (client.isNone()) return null
  return client.orThrow()
}

export const listGroups = async (params: { top?: number; filter?: string }): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.listGroups({ $top: params.top ?? 25, $filter: params.filter })
  return result
    .mapLeft((error) => new UserError(`Failed to list groups: ${error.message}`))
    .map((response) => formatGroupList((response as ODataResponse<never>).value))
}

export const getGroup = async (params: { group_id: string }): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.getGroup(params.group_id)
  return result.mapLeft((error) => new UserError(`Failed to get group: ${error.message}`)).map(formatGroupDetail)
}

export const listGroupMembers = async (params: { group_id: string }): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.listGroupMembers(params.group_id)
  return result
    .mapLeft((error) => new UserError(`Failed to list group members: ${error.message}`))
    .map((response) => formatUserList((response as ODataResponse<never>).value))
}
