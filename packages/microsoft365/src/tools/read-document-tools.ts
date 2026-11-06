import { filenameFromPath, formatBytes, resilientFetch } from "@sapientsai/ms-graph-core"
import type { Either } from "functype/either"
import { Left, Right } from "functype/either"
import { UserError } from "somamcp"

import { getAccessToken } from "../auth"
import { GRAPH_API_BASE } from "../auth/scopes"
import { getGraphClient } from "../client/graph-client"
import { defaultVersion } from "../client/graph-version"
import { resolveMailbox } from "../mail/mailbox"
import { mailboxOfUrl } from "../mail/mailbox-path"
import type { GraphApiVersion, GraphDriveItem } from "../types"

const requireClient = () => {
  const client = getGraphClient()
  if (client.isNone()) return null
  return client.orThrow()
}

const MB = 1024 * 1024
const PDF = "application/pdf"
const DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
const XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

// Per-format input caps. Memory cost is not uniform in file size: unpdf is roughly linear, mammoth
// ignores the embedded images that make up most of a large DOCX, and exceljs expands a workbook into
// a full object model at 10-20x file size — which is why XLSX is the tightest, not PDF.
// These bound container heap only. Output is bounded separately by max_chars.
const FORMAT_CAPS: Record<string, number> = {
  [PDF]: 100 * MB,
  [DOCX]: 50 * MB,
  [XLSX]: 25 * MB,
}
const DEFAULT_CAP = 25 * MB
const DEFAULT_MAX_CHARS = 50000

// MS365_MAX_EXTRACT_BYTES is a ceiling over the per-format defaults, never a raise: one variable
// tightens every format in a memory-constrained container. Unset or unparseable means no ceiling.
const envCeiling = (): number => {
  const parsed = Number.parseInt(process.env.MS365_MAX_EXTRACT_BYTES ?? "", 10)
  return Number.isFinite(parsed) && parsed > 0 ? parsed : Number.POSITIVE_INFINITY
}

const capForMimeType = (mimeType?: string): number => Math.min(FORMAT_CAPS[mimeType ?? ""] ?? DEFAULT_CAP, envCeiling())

// Two content-endpoint shapes reach this tool, and their metadata resources are NOT the same:
//
//   drive item       /me/drive/items/{id}/content              -> size, file.mimeType
//   mail attachment  /me/messages/{id}/attachments/{a}/$value  -> size, contentType (top level)
//
// The original strip handled only "/content", so an attachment path kept its "/$value" suffix and
// the $select was appended to an Edm.Stream resource — which Graph rejects outright:
// "The type 'Edm.Stream' is not valid for $select or $expand". Every mail-attachment read failed.
const metaPathFrom = (path: string): string => path.replace(/\/(content|\$value)\/?$/, "")

const isAttachmentPath = (path: string): boolean => /\/attachments\/[^/]+\/\$value\/?$/.test(path)

// A "/$value" that is not an attachment is the raw MIME of a whole message. Its metadata resource is
// a message, which has neither `name` nor `file`, so the probe below would 400 just as opaquely as
// the bug this replaces. Say so instead.
const isMessageMimePath = (path: string): boolean => /\/\$value\/?$/.test(path) && !isAttachmentPath(path)

const httpError = async (response: Response): Promise<UserError> => {
  if (response.headers.get("content-type")?.includes("application/json")) {
    const data = (await response.json()) as { error?: { message?: string } }
    return new UserError(data.error?.message ?? `HTTP ${response.status}: ${response.statusText}`)
  }
  return new UserError(`HTTP ${response.status}: ${response.statusText}`)
}

// read_document: fetch a Graph file and return its extracted text.
//
// Two things here are load-bearing and easy to get wrong:
//
// 1. Metadata comes FIRST. Checking size after arrayBuffer() is too late — for the very file the cap
//    exists to reject, the process dies inside arrayBuffer() and the check never runs. Reading size
//    from the drive item costs one round trip and makes the cap real.
// 2. The token must come from getAccessToken(), which resolves the per-request token that oauth-proxy
//    mode puts in async context. A credential-based path would work in stdio mode and silently break
//    the deployed oauth-proxy server.
export const readDocument = async (params: {
  path: string
  api_version?: GraphApiVersion
  format?: string
  max_chars?: number
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const version = params.api_version ?? defaultVersion()

  // list_attachments prints /users/{address}/... paths for other mailboxes, so this is the same
  // door the mail tools guard and has to apply the same mailbox check, before any request.
  const mailbox = mailboxOfUrl(`${GRAPH_API_BASE}/${version}${params.path}`)
  if (mailbox !== undefined) {
    const scope = await resolveMailbox(mailbox, client)
    if (scope.isLeft()) return Left(scope.value as UserError)
  }

  const tokenResult = await getAccessToken()
  if (tokenResult.isLeft()) {
    return Left(new UserError((tokenResult.value as { message: string }).message))
  }
  const token = tokenResult.value as string

  if (isMessageMimePath(params.path)) {
    return Left(
      new UserError(
        "Reading a whole message as MIME is not supported. Use get_message for the body, or point " +
          "read_document at one of its attachments (.../attachments/{id}/$value).",
      ),
    )
  }

  const attachment = isAttachmentPath(params.path)

  const metaResult = await client.graphQuery<GraphDriveItem & { contentType?: string }>(
    "GET",
    `${metaPathFrom(params.path)}?$select=${attachment ? "id,name,contentType,size" : "id,name,size,file"}`,
    undefined,
    version,
  )
  if (metaResult.isLeft()) {
    return Left(new UserError(`Failed to get file info: ${(metaResult.value as { message: string }).message}`))
  }
  const meta = metaResult.value as GraphDriveItem & { contentType?: string }
  const declaredMimeType = attachment ? meta.contentType : meta.file?.mimeType

  const cap = capForMimeType(declaredMimeType)
  const size = meta.size ?? 0
  if (size > cap) {
    return Left(
      new UserError(
        `File is ${formatBytes(size)}, over the ${formatBytes(cap)} extraction cap for this format. ` +
          `Use download_file to get a download URL instead.`,
      ),
    )
  }

  const query = params.format ? `?format=${params.format}` : ""
  const response = await resilientFetch(`${GRAPH_API_BASE}/${version}${params.path}${query}`, {
    headers: { Authorization: `Bearer ${token}` },
  })
  if (!response.ok) return Left(await httpError(response))

  const buffer = Buffer.from(await response.arrayBuffer())
  const contentType = response.headers.get("content-type") ?? declaredMimeType ?? "application/octet-stream"
  const filename = meta.name ?? filenameFromPath(params.path) ?? "download"

  // Lazy — keeps mammoth/unpdf/exceljs off the startup path of a server that mostly does mail and
  // calendar. A top-level import here would load all three in every deployment, since MS365_PRESETS
  // is unset by default and the preset filter therefore admits every tool.
  const { extractTextFromBuffer } = await import("@sapientsai/document-extract")
  const extracted = await extractTextFromBuffer(buffer, contentType, filename)
  if (extracted.isLeft()) {
    const error = extracted.value as { type: string; message: string }
    return Left(
      new UserError(
        error.type === "unsupported"
          ? `${error.message} Use download_file to get the raw bytes instead.`
          : error.message,
      ),
    )
  }

  const fullText = extracted.value as string
  const maxChars = params.max_chars ?? DEFAULT_MAX_CHARS
  const text =
    fullText.length > maxChars
      ? `${fullText.slice(0, maxChars)}\n\n[truncated at ${maxChars.toLocaleString()} chars — full document is ${fullText.length.toLocaleString()} chars]`
      : fullText

  return Right(`File: ${filename} (${formatBytes(size)})\n\n${text}`)
}
