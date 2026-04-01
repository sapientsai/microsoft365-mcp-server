import { UserError } from "fastmcp"
import type { Either } from "functype/either"
import { Left } from "functype/either"

import { getGraphClient } from "../client/graph-client"
import type { GraphDriveItem, ODataResponse } from "../types"
import { formatDriveItemDetail, formatDriveItemList } from "../utils/formatters"

const requireClient = () => {
  const client = getGraphClient()
  if (client.isNone()) return null
  return client.orThrow()
}

export const listDriveItems = async (params: {
  folder_id?: string
  fetch_all_pages?: boolean
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  if (params.fetch_all_pages) {
    const path = params.folder_id ? `/me/drive/items/${params.folder_id}/children` : "/me/drive/root/children"
    const result = await client.requestPaginated<GraphDriveItem>(path)
    return result
      .mapLeft((error) => new UserError(`Failed to list drive items: ${error.message}`))
      .map((items) => formatDriveItemList(items))
  }

  const result = await client.listDriveItems(params.folder_id)
  return result
    .mapLeft((error) => new UserError(`Failed to list drive items: ${error.message}`))
    .map((response) => formatDriveItemList((response as ODataResponse<never>).value))
}

export const getDriveItem = async (params: { item_id: string }): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.getDriveItem(params.item_id)
  return result
    .mapLeft((error) => new UserError(`Failed to get drive item: ${error.message}`))
    .map(formatDriveItemDetail)
}

export const searchFiles = async (params: { query: string }): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.searchFiles(params.query)
  return result
    .mapLeft((error) => new UserError(`Failed to search files: ${error.message}`))
    .map((response) => formatDriveItemList((response as ODataResponse<never>).value))
}

export const downloadFile = async (params: { item_id: string }): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.downloadFile(params.item_id)
  return result
    .mapLeft((error) => new UserError(`Failed to get file info: ${error.message}`))
    .map(formatDriveItemDetail)
}

export const createFolder = async (params: { parent_id: string; name: string }): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.createFolder(params.parent_id, params.name)
  return result
    .mapLeft((error) => new UserError(`Failed to create folder: ${error.message}`))
    .map((item) => `Folder created.\n\n${formatDriveItemDetail(item)}`)
}
