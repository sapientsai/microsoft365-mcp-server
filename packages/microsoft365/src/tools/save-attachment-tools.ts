import { tmpdir } from "node:os"
import { extname, isAbsolute, join, resolve } from "node:path"

import { formatBytes, resilientFetch } from "@sapientsai/ms-graph-core"
import { UserError } from "fastmcp"
import type { Either } from "functype/either"
import { Left, Right } from "functype/either"

import { getAccessToken } from "../auth"
import { GRAPH_API_BASE } from "../auth/scopes"
import { getGraphClient } from "../client/graph-client"
import { defaultVersion } from "../client/graph-version"
import { resolveMailbox } from "../mail/mailbox"
import type { GraphAttachment } from "../types"
import { linkLocation } from "../utils/link-url"
import { writeNewFile } from "../utils/local-file"
import { describeRefFailure, resolveMessageIdOrRef } from "../utils/message-refs"

const requireClient = () => {
  const client = getGraphClient()
  if (client.isNone()) return null
  return client.orThrow()
}

// Attachments are written to disk rather than returned inline, so this cap is about disk and one
// arrayBuffer() in memory, not about what fits in a tool response. Graph itself refuses to serve
// fileAttachment bytes much beyond this.
const MAX_ATTACHMENT_SIZE = 150 * 1024 * 1024

// Content types Graph reports that do not match the extension people expect on disk.
const EXTENSION_FOR_TYPE: Record<string, string> = {
  "application/pdf": ".pdf",
  "image/jpeg": ".jpg",
  "image/png": ".png",
  "image/gif": ".gif",
  "image/webp": ".webp",
  "image/heic": ".heic",
  "image/tiff": ".tif",
}

const ITEM_ATTACHMENT = "#microsoft.graph.itemAttachment"

// Windows refuses these as a base name whatever follows the first dot, so "CON.pdf" fails too.
const WINDOWS_RESERVED = /^(con|prn|aux|nul|com[0-9]|lpt[0-9])$/i

// Most filesystems cap a name at 255 bytes, not characters. This leaves room for the " (99)" that
// writeNewFile may add.
const MAX_NAME_BYTES = 200
const MAX_EXTENSION_BYTES = 16

// Cuts on character boundaries so a multi-byte character is never split.
const truncateToBytes = (text: string, maxBytes: number): string => {
  const chars = Array.from(text)
  const overflowAt = chars.findIndex((_, i) => Buffer.byteLength(chars.slice(0, i + 1).join("")) > maxBytes)
  return overflowAt === -1 ? text : chars.slice(0, overflowAt).join("")
}

// Windows-illegal characters plus path separators and control codes. A Graph attachment name is
// attacker-influenced in the sense that anyone who emails you picks it, so it never reaches the
// filesystem unfiltered — "../../.ssh/authorized_keys" is a legal attachment name.
const safeName = (attachment: GraphAttachment): string => {
  const cleaned = (attachment.name ?? "")
    // eslint-disable-next-line no-control-regex
    .replace(/[\u0000-\u001f<>:"/\\|?*]/g, "_")
    .replace(/^\.+/, "")
    .trim()
  const name = cleaned || `attachment-${attachment.id.slice(0, 12)}`

  // An attached email's $value is MIME and its name is the subject, so a dot in the subject
  // ("Q3 v1.2") is not an extension.
  const isItem = attachment["@odata.type"] === ITEM_ATTACHMENT
  const ownExtension = extname(name)
  const keepsOwnExtension =
    ownExtension !== "" &&
    Buffer.byteLength(ownExtension) <= MAX_EXTENSION_BYTES &&
    (!isItem || ownExtension.toLowerCase() === ".eml")
  const extension = keepsOwnExtension
    ? ownExtension
    : isItem
      ? ".eml"
      : (EXTENSION_FOR_TYPE[attachment.contentType ?? ""] ?? (cleaned ? "" : ".bin"))
  const stem = keepsOwnExtension ? name.slice(0, -ownExtension.length) : name

  const allowedStem = WINDOWS_RESERVED.test(stem.split(".")[0]?.trim() ?? "") ? `_${stem}` : stem
  return `${truncateToBytes(allowedStem, MAX_NAME_BYTES - Buffer.byteLength(extension))}${extension}`
}

// Up to "name (99).ext": saving is additive, so a same-named attachment from another message gets a
// numbered name instead of replacing the earlier save.
const MAX_NAMES = 100
const REFERENCE_ATTACHMENT = "#microsoft.graph.referenceAttachment"

const httpError = async (response: Response): Promise<UserError> => {
  if (response.headers.get("content-type")?.includes("application/json")) {
    const data = (await response.json()) as { error?: { message?: string } }
    return new UserError(data.error?.message ?? `HTTP ${response.status}: ${response.statusText}`)
  }
  return new UserError(`HTTP ${response.status}: ${response.statusText}`)
}

// save_attachment: write a mail attachment to a local file and return its path.
//
// This exists because read_document is text extraction only. A scanned PDF, a photographed letter or
// any image comes back empty from it — not because the bytes are unreachable, but because there is no
// text layer to extract. Rather than grow this server an image pipeline (rasterising, OCR, resizing),
// it hands over the file and lets the caller use whatever already reads PDFs and images. Composition
// over a bigger Graph server.
//
// Note the deliberate omission: nothing here returns bytes inline. An MCP tool result carrying base64
// is not something a caller can look at, and it would blow the response budget for a large scan.
// A path is small, and the tools that read files are better at it than this server would be.
export const saveAttachment = async (params: {
  message_id: string
  attachment_id?: string
  out_dir?: string
  mailbox?: string
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const resolvedScope = await resolveMailbox(params.mailbox, client)
  if (resolvedScope.isLeft()) return Left(resolvedScope.value as UserError)
  const scope = resolvedScope.orThrow()

  const resolved = resolveMessageIdOrRef(params.message_id, scope.mailbox)
  if (resolved.kind !== "id") return Left(new UserError(describeRefFailure(resolved, scope.mailbox)))
  const messageId = resolved.id

  const listResult = await client.listAttachments(messageId, scope.prefix)
  if (listResult.isLeft()) {
    return Left(new UserError(`Failed to list attachments: ${(listResult.value as { message: string }).message}`))
  }
  const attachments = (listResult.value as { value?: ReadonlyArray<GraphAttachment> }).value ?? []

  // A referenceAttachment is a link to OneDrive/SharePoint/Dropbox; the mailbox holds no bytes for
  // it, so $value cannot serve one. It is excluded from what can be saved — but never silently:
  // dropping it would hide a document that exists, so its URL is reported instead.
  const references = attachments.filter((a) => a["@odata.type"] === REFERENCE_ATTACHMENT)
  const all = attachments.filter((a) => a["@odata.type"] !== REFERENCE_ATTACHMENT)

  const describeReferences = () =>
    references.map((r) => `  ${r.name ?? "(unnamed)"}${r.isFolder ? " (folder)" : ""} — ${linkLocation(r)}`).join("\n")

  if (all.length === 0) {
    return Left(
      new UserError(
        references.length === 0
          ? "This message has no attachments."
          : `This message has no downloadable file attachments, but it does carry ${references.length} cloud ` +
              `link${references.length === 1 ? "" : "s"} (reference attachment${references.length === 1 ? "" : "s"}). ` +
              `The mailbox holds no bytes for these:\n${describeReferences()}`,
      ),
    )
  }

  // No attachment_id: only unambiguous when there is exactly one. Guessing which of several a caller
  // meant is worse than making them look.
  const chosen = params.attachment_id
    ? all.find((a) => a.id === params.attachment_id)
    : all.length === 1
      ? all[0]
      : undefined

  if (!chosen) {
    if (params.attachment_id) {
      const asReference = references.find((r) => r.id === params.attachment_id)
      if (asReference) {
        return Left(
          new UserError(
            `"${asReference.name ?? params.attachment_id}" is a cloud link (reference attachment), not a file in ` +
              `the mailbox, so there are no bytes to save. Link: ${linkLocation(asReference)}`,
          ),
        )
      }
      return Left(new UserError(`No attachment "${params.attachment_id}" on this message. Use list_attachments.`))
    }
    const names = all.map((a) => `  ${a.id}  ${a.name ?? "(unnamed)"} (${formatBytes(a.size ?? 0)})`).join("\n")
    const alsoLinks =
      references.length === 0
        ? ""
        : `\n\nIt also carries ${references.length} cloud link${references.length === 1 ? "" : "s"}, which cannot ` +
          `be saved from the mailbox:\n${describeReferences()}`
    return Left(
      new UserError(
        `This message has ${all.length} attachments — pass attachment_id to choose one:\n${names}${alsoLinks}`,
      ),
    )
  }

  const size = chosen.size ?? 0
  if (size > MAX_ATTACHMENT_SIZE) {
    return Left(new UserError(`Attachment is ${formatBytes(size)}, over the ${formatBytes(MAX_ATTACHMENT_SIZE)} cap.`))
  }

  const tokenResult = await getAccessToken()
  if (tokenResult.isLeft()) return Left(new UserError((tokenResult.value as { message: string }).message))
  const token = tokenResult.value as string

  const response = await resilientFetch(
    `${GRAPH_API_BASE}/${defaultVersion()}${scope.prefix}/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(chosen.id)}/$value`,
    { headers: { Authorization: `Bearer ${token}` } },
  )
  if (!response.ok) return Left(await httpError(response))

  const buffer = Buffer.from(await response.arrayBuffer())

  const dir = params.out_dir ? (isAbsolute(params.out_dir) ? params.out_dir : resolve(params.out_dir)) : tmpdir()
  const filename = safeName(chosen)

  const written = await writeNewFile(dir, filename, buffer, MAX_NAMES)
  if (written.isLeft()) return Left(new UserError(`Failed to write ${join(dir, filename)}: ${written.value.message}`))
  const saved = written.value

  const type =
    chosen.contentType ?? (chosen["@odata.type"] === ITEM_ATTACHMENT ? "message/rfc822" : "application/octet-stream")
  return Right(
    `Saved: ${saved}\n` +
      `Type: ${type}\n` +
      `Size: ${formatBytes(buffer.length)}\n\n` +
      `Read the file directly — PDFs and images do not need text extraction to be viewed.`,
  )
}
