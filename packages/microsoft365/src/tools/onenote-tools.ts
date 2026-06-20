import { UserError } from "fastmcp"
import type { Either } from "functype/either"
import { Left } from "functype/either"

import { getGraphClient } from "../client/graph-client"
import type { GraphNotebook, GraphPage, GraphSection, ODataResponse } from "../types"
import { formatNotebookList, formatPageList, formatSectionList } from "../utils/formatters"

const requireClient = () => {
  const client = getGraphClient()
  if (client.isNone()) return null
  return client.orThrow()
}

const escapeHtml = (value: string): string => value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")

// OneNote create-page expects a full HTML document; the <title> becomes the page title.
const buildPageHtml = (title: string, content: string): string =>
  `<!DOCTYPE html><html><head><title>${escapeHtml(title)}</title></head><body>${content}</body></html>`

export const listOnenoteNotebooks = async (params?: {
  fetch_all_pages?: boolean
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  if (params?.fetch_all_pages) {
    const result = await client.requestPaginated<GraphNotebook>("/me/onenote/notebooks")
    return result
      .mapLeft((error) => new UserError(`Failed to list notebooks: ${error.message}`))
      .map((items) => formatNotebookList(items))
  }

  const result = await client.listOnenoteNotebooks()
  return result
    .mapLeft((error) => new UserError(`Failed to list notebooks: ${error.message}`))
    .map((response) => formatNotebookList((response as ODataResponse<never>).value))
}

export const listOnenoteSections = async (params: {
  notebook_id: string
  fetch_all_pages?: boolean
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  if (params.fetch_all_pages) {
    const result = await client.requestPaginated<GraphSection>(`/me/onenote/notebooks/${params.notebook_id}/sections`)
    return result
      .mapLeft((error) => new UserError(`Failed to list sections: ${error.message}`))
      .map((items) => formatSectionList(items))
  }

  const result = await client.listOnenoteSections(params.notebook_id)
  return result
    .mapLeft((error) => new UserError(`Failed to list sections: ${error.message}`))
    .map((response) => formatSectionList((response as ODataResponse<never>).value))
}

export const listOnenotePages = async (params: {
  section_id: string
  fetch_all_pages?: boolean
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  if (params.fetch_all_pages) {
    const result = await client.requestPaginated<GraphPage>(`/me/onenote/sections/${params.section_id}/pages`)
    return result
      .mapLeft((error) => new UserError(`Failed to list pages: ${error.message}`))
      .map((items) => formatPageList(items))
  }

  const result = await client.listOnenotePages(params.section_id)
  return result
    .mapLeft((error) => new UserError(`Failed to list pages: ${error.message}`))
    .map((response) => formatPageList((response as ODataResponse<never>).value))
}

export const getOnenotePageContent = async (params: { page_id: string }): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.getOnenotePageContent(params.page_id)
  return result
    .mapLeft((error) => new UserError(`Failed to get page content: ${error.message}`))
    .map((content) => `# Page Content\n\n${content}`)
}

export const createOnenotePage = async (params: {
  section_id: string
  title: string
  content: string
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const html = buildPageHtml(params.title, params.content)
  const result = await client.createOnenotePage(params.section_id, html)
  return result
    .mapLeft((error) => new UserError(`Failed to create page: ${error.message}`))
    .map((page) => `Page created. ID: ${(page as { id: string }).id} — "${params.title}"`)
}

export const updateOnenotePageContent = async (params: {
  page_id: string
  content: string
  action?: string
  target?: string
  position?: string
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const command: Record<string, string> = {
    target: params.target ?? "body",
    action: params.action ?? "append",
    content: params.content,
  }
  if (params.position) command.position = params.position

  const result = await client.updateOnenotePageContent(params.page_id, [command])
  return result
    .mapLeft((error) => new UserError(`Failed to update page content: ${error.message}`))
    .map(() => `Page content updated (${command.action} on "${command.target}").`)
}

export const createOnenoteSection = async (params: {
  notebook_id: string
  display_name: string
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.createOnenoteSection(params.notebook_id, params.display_name)
  return result
    .mapLeft((error) => new UserError(`Failed to create section: ${error.message}`))
    .map((section) => `Section created. ID: ${(section as { id: string }).id} — "${params.display_name}"`)
}

export const createOnenoteNotebook = async (params: { display_name: string }): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.createOnenoteNotebook(params.display_name)
  return result
    .mapLeft((error) => new UserError(`Failed to create notebook: ${error.message}`))
    .map((notebook) => `Notebook created. ID: ${(notebook as { id: string }).id} — "${params.display_name}"`)
}

export const copyOnenotePage = async (params: {
  page_id: string
  section_id: string
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.copyOnenotePage(params.page_id, params.section_id)
  return result
    .mapLeft((error) => new UserError(`Failed to copy page: ${error.message}`))
    .map(() => `Copy initiated to section ${params.section_id}. Copy runs asynchronously in OneNote.`)
}

export const deleteOnenotePage = async (params: { page_id: string }): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.deleteOnenotePage(params.page_id)
  return result
    .mapLeft((error) => new UserError(`Failed to delete page: ${error.message}`))
    .map(() => `Page ${params.page_id} deleted.`)
}
