import { UserError } from "fastmcp"
import type { Either } from "functype/either"
import { Left, Right } from "functype/either"

import { getGraphClient } from "../client/graph-client"
import type { GraphDriveItem, ODataResponse } from "../types"
import { formatDriveItemDetail, formatDriveItemList } from "../utils/formatters"

const requireClient = () => {
  const client = getGraphClient()
  if (client.isNone()) return null
  return client.orThrow()
}

const TEXT_MIME_PREFIXES = ["text/", "application/json", "application/xml", "application/javascript"]
const MAX_INLINE_SIZE = 100 * 1024 // 100KB

const isTextFile = (mimeType?: string): boolean =>
  mimeType !== undefined && TEXT_MIME_PREFIXES.some((prefix) => mimeType.startsWith(prefix))

export const listDriveItems = async (params: {
  folder_id?: string
  folder_path?: string
  fetch_all_pages?: boolean
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  if (params.folder_path) {
    const path = params.folder_path.replace(/^\/+/, "")
    if (params.fetch_all_pages) {
      const result = await client.requestPaginated<GraphDriveItem>(`/me/drive/root:/${path}:/children`)
      return result
        .mapLeft((error) => new UserError(`Failed to list drive items: ${error.message}`))
        .map((items) => formatDriveItemList(items))
    }
    const result = await client.listDriveItemsByPath(path)
    return result
      .mapLeft((error) => new UserError(`Failed to list drive items: ${error.message}`))
      .map((response) => formatDriveItemList((response as ODataResponse<never>).value))
  }

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

  const metaResult = await client.downloadFile(params.item_id)
  if (metaResult.isLeft()) {
    return Left(new UserError(`Failed to get file info: ${(metaResult.value as { message: string }).message}`))
  }

  const item = metaResult.value as GraphDriveItem
  const detail = formatDriveItemDetail(item)

  if (isTextFile(item.file?.mimeType) && (item.size ?? 0) <= MAX_INLINE_SIZE) {
    const contentResult = await client.downloadFileContent(params.item_id)
    if (contentResult.isRight()) {
      const content = contentResult.value as string
      return Right(`${detail}\n\n## Content\n\n\`\`\`\n${content}\n\`\`\``)
    }
  }

  return Right(detail)
}

export const createFolder = async (params: { parent_id: string; name: string }): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.createFolder(params.parent_id, params.name)
  return result
    .mapLeft((error) => new UserError(`Failed to create folder: ${error.message}`))
    .map((item) => `Folder created.\n\n${formatDriveItemDetail(item)}`)
}

export const uploadFile = async (params: {
  path: string
  content: string
  content_type?: string
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.uploadFile(params.path, params.content, params.content_type ?? "text/plain")
  return result
    .mapLeft((error) => new UserError(`Failed to upload file: ${error.message}`))
    .map((item) => `File uploaded.\n\n${formatDriveItemDetail(item)}`)
}
