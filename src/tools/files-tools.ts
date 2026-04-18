import { readFile, stat } from "node:fs/promises"
import { basename } from "node:path"

import { UserError } from "fastmcp"
import type { Either } from "functype/either"
import { Left, Right } from "functype/either"

import { getAccessToken } from "../auth"
import { GRAPH_API_BASE } from "../auth/scopes"
import { getContextToken } from "../auth/token-context"
import { getGraphClient } from "../client/graph-client"
import type { GraphDriveItem, ODataResponse } from "../types"
import { MAX_UPLOAD_SIZE, sessionUpload, SIMPLE_UPLOAD_LIMIT, simpleUpload } from "../upload/upload"
import { formatDriveItemDetail, formatDriveItemList } from "../utils/formatters"
import { filenameFromPath, formatBytes, resolveUploadContentType } from "../utils/upload-helpers"

const requireClient = () => {
  const client = getGraphClient()
  if (client.isNone()) return null
  return client.orThrow()
}

const isENOENT = (error: unknown): boolean =>
  typeof error === "object" && error !== null && "code" in error && (error as { code: unknown }).code === "ENOENT"

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

const UPLOAD_FILE_MAX_BYTES = 4 * 1024 * 1024 // 4 MB — Graph simple-PUT ceiling

const BINARY_REDIRECT_MESSAGE =
  "Binary uploads are not supported via upload_file. Use get_upload_config (HTTP/SSE deployments) or upload_file_from_path (stdio/local) instead — both stream binary directly to Graph without round-tripping bytes through the LLM."

const TEXT_CONTENT_TYPE_ALLOWLIST = new Set([
  "application/json",
  "application/xml",
  "application/javascript",
  "application/x-www-form-urlencoded",
])

const isTextContentType = (contentType: string): boolean => {
  const lower = contentType.toLowerCase().split(";")[0]?.trim() ?? ""
  if (lower.startsWith("text/")) return true
  if (TEXT_CONTENT_TYPE_ALLOWLIST.has(lower)) return true
  if (lower.endsWith("+json") || lower.endsWith("+xml")) return true
  return false
}

export const uploadFile = async (params: {
  path: string
  content: string
  content_type?: string
  conflict_behavior?: "rename" | "replace" | "fail"
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const contentType = params.content_type ?? "text/plain"
  if (!isTextContentType(contentType)) {
    return Left(new UserError(BINARY_REDIRECT_MESSAGE))
  }

  const size = Buffer.byteLength(params.content, "utf8")
  if (size > UPLOAD_FILE_MAX_BYTES) {
    return Left(
      new UserError(
        `Text content is ${size} bytes (> 4 MB simple-PUT limit). Use get_upload_config (HTTP/SSE) or upload_file_from_path (stdio) for larger files.`,
      ),
    )
  }

  const result = await client.uploadFile(params.path, params.content, contentType, params.conflict_behavior ?? "rename")
  return result
    .mapLeft((error) => new UserError(`Failed to upload file: ${error.message}`))
    .map((item) => `File uploaded.\n\n${formatDriveItemDetail(item)}`)
}

export const getUploadConfig = async (params: {
  path: string
  localFile?: string
  contentType?: string
  conflictBehavior?: "rename" | "replace" | "fail"
}): Promise<Either<UserError, string>> => {
  await Promise.resolve()
  if (!/:\/content$/i.test(params.path)) {
    return Left(new UserError('path must end with ":/content" (e.g., /me/drive/root:/Documents/file.docx:/content)'))
  }

  const baseUrl = resolveUploadBaseUrl()
  if (!baseUrl) {
    return Left(
      new UserError(
        "Upload endpoint base URL not configured. Set MS365_PUBLIC_BASE_URL (or MS365_OAUTH_BASE_URL in oauth-proxy mode), or run with TRANSPORT_TYPE=httpStream.",
      ),
    )
  }

  const filename = filenameFromPath(params.path)
  const contentType = resolveUploadContentType(params.contentType, filename)
  const conflictBehavior = params.conflictBehavior ?? "rename"
  const localFile = params.localFile ?? "{local_file_path}"

  const query = new URLSearchParams({
    path: params.path,
    conflictBehavior,
    contentType,
    encoding: "base64",
  })
  const uploadUrl = `${baseUrl}/upload?${query.toString()}`

  const token = getContextToken() ?? process.env.MS365_UPLOAD_TOKEN
  const authHeader = token ? `Authorization: Bearer ${token}` : undefined

  const curlParts = [
    `base64 "${localFile}" | tr -d '\\n'`,
    "| curl -X POST",
    authHeader ? `-H "${authHeader}"` : undefined,
    `-H "Content-Type: text/plain"`,
    `--data-binary @-`,
    `"${uploadUrl}"`,
  ]
    .filter(Boolean)
    .join(" \\\n  ")

  const payload = {
    uploadUrl,
    method: "POST",
    contentType,
    conflictBehavior,
    maxSize: "250 MB",
    encoding: "base64",
    ...(authHeader ? { authHeader } : {}),
    curl: curlParts,
    notes: [
      "Pipe base64-encoded file bytes to uploadUrl via POST with --data-binary @-.",
      "Server decodes base64, then PUTs to Microsoft Graph (simple PUT up to 4 MB; chunked session upload above).",
      "Intermediate folders in the Graph path are auto-created.",
      "Known issue: curl from claude.ai code-execution sandbox may return HTTP 503 'DNS cache overflow' on binary bodies >~40KB — this is a sandbox egress-proxy bug, not a server error. If this occurs, the upload likely did not succeed; retry from a local shell (WSL/VM/Claude Code CLI). For text/markdown/JSON, prefer the upload_file MCP tool instead (different egress path).",
    ],
  }

  return Right(JSON.stringify(payload, null, 2))
}

export const uploadFileFromPath = async (params: {
  local_path: string
  path: string
  content_type?: string
  conflict_behavior?: "rename" | "replace" | "fail"
}): Promise<Either<UserError, string>> => {
  if (!/:\/content$/i.test(params.path)) {
    return Left(new UserError('path must end with ":/content" (e.g., /me/drive/root:/Documents/file.docx:/content)'))
  }

  const fileNotFoundMessage = `File not found: ${params.local_path}. If this file was generated in a cloud environment (e.g., claude.ai), use Desktop Commander's write_file to save it to the local filesystem first, then retry.`

  const stats = await stat(params.local_path).catch((error: unknown) => error as Error)
  if (stats instanceof Error) {
    if (isENOENT(stats)) return Left(new UserError(fileNotFoundMessage))
    return Left(new UserError(`Cannot read local file: ${stats.message}`))
  }
  if (!stats.isFile()) {
    return Left(new UserError(`local_path is not a regular file: ${params.local_path}`))
  }
  if (stats.size > MAX_UPLOAD_SIZE) {
    return Left(new UserError(`File too large (${formatBytes(stats.size)}, max ${formatBytes(MAX_UPLOAD_SIZE)})`))
  }

  const buffer = await readFile(params.local_path).catch((error: unknown) => error as Error)
  if (buffer instanceof Error) {
    if (isENOENT(buffer)) return Left(new UserError(fileNotFoundMessage))
    return Left(new UserError(`Failed to read local file: ${buffer.message}`))
  }

  const sessionToken = getContextToken()
  const resolvedToken =
    sessionToken ??
    (await (async () => {
      const result = await getAccessToken()
      return result.fold<string | undefined>(
        () => undefined,
        (value) => value as string,
      )
    })())
  if (!resolvedToken) {
    return Left(new UserError("No access token available. Check authentication mode."))
  }

  const filename = filenameFromPath(params.path) ?? basename(params.local_path)
  const contentType = resolveUploadContentType(params.content_type, filename)
  const conflictBehavior = params.conflict_behavior ?? "rename"
  const apiBase = `${GRAPH_API_BASE}/v1.0`

  console.error(
    `[Upload] local=${params.local_path} bytes=${buffer.length} contentType=${contentType} mode=${buffer.length <= SIMPLE_UPLOAD_LIMIT ? "simple" : "session"}`,
  )

  const result =
    buffer.length <= SIMPLE_UPLOAD_LIMIT
      ? await simpleUpload(apiBase, params.path, resolvedToken, buffer, contentType, conflictBehavior)
      : await sessionUpload(apiBase, params.path, resolvedToken, buffer, conflictBehavior)

  return result
    .mapLeft((error) => new UserError(`Failed to upload file: ${error.message}`))
    .map((item) => `File uploaded (${formatBytes(buffer.length)}).\n\n${formatDriveItemDetail(item)}`)
}

const resolveUploadBaseUrl = (): string | undefined => {
  const explicit = process.env.MS365_PUBLIC_BASE_URL ?? process.env.MS365_OAUTH_BASE_URL
  if (explicit) return explicit.replace(/\/$/, "")

  if (process.env.TRANSPORT_TYPE === "httpStream" || process.env.MS365_AUTH_MODE === "oauth-proxy") {
    const port = process.env.PORT ?? "3000"
    const host = process.env.HOST ?? "127.0.0.1"
    return `http://${host}:${port}`
  }

  return undefined
}
