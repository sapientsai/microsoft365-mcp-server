import { UserError } from "fastmcp"
import type { Either } from "functype/either"
import { Left } from "functype/either"

import { getGraphClient } from "../client/graph-client"
import type { ODataResponse } from "../types"
import { formatNotebookList, formatPageList, formatSectionList } from "../utils/formatters"

const requireClient = () => {
  const client = getGraphClient()
  if (client.isNone()) return null
  return client.orThrow()
}

export const listNotebooks = async (): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.listNotebooks()
  return result
    .mapLeft((error) => new UserError(`Failed to list notebooks: ${error.message}`))
    .map((response) => formatNotebookList((response as ODataResponse<never>).value))
}

export const listSections = async (params: { notebook_id: string }): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.listSections(params.notebook_id)
  return result
    .mapLeft((error) => new UserError(`Failed to list sections: ${error.message}`))
    .map((response) => formatSectionList((response as ODataResponse<never>).value))
}

export const listPages = async (params: { section_id: string }): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.listPages(params.section_id)
  return result
    .mapLeft((error) => new UserError(`Failed to list pages: ${error.message}`))
    .map((response) => formatPageList((response as ODataResponse<never>).value))
}

export const getPageContent = async (params: { page_id: string }): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.getPageContent(params.page_id)
  return result
    .mapLeft((error) => new UserError(`Failed to get page content: ${error.message}`))
    .map((content) => `# Page Content\n\n${content}`)
}
