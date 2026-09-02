import { mkdir, writeFile } from "node:fs/promises"
import { tmpdir } from "node:os"
import { extname, isAbsolute, join, resolve } from "node:path"

import { formatBytes } from "@sapientsai/ms-graph-core"
import { UserError } from "fastmcp"
import type { Either } from "functype/either"
import { Left, Right } from "functype/either"

import { getAccessToken } from "../auth"
import { GRAPH_API_BASE } from "../auth/scopes"
import { getGraphClient } from "../client/graph-client"
import type { GraphAttachment } from "../types"
import { resolveMessageIdOrRef } from "../utils/message-refs"

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

// Windows-illegal characters plus path separators and control codes. A Graph attachment name is
// attacker-influenced in the sense that anyone who emails you picks it, so it never reaches the
// filesystem unfiltered — "../../.ssh/authorized_keys" is a legal attachment name.
const safeName = (name: string | undefined, contentType: string | undefined, id: string): string => {
  const cleaned = (name ?? "")
    // eslint-disable-next-line no-control-regex
    .replace(/[\u0000-\u001f<>:"/\\|?*]/g, "_")
    .replace(/^\.+/, "")
    .trim()
  if (cleaned) {
    return extname(cleaned) ? cleaned : `${cleaned}${EXTENSION_FOR_TYPE[contentType ?? ""] ?? ""}`
  }
  return `attachment-${id.slice(0, 12)}${EXTENSION_FOR_TYPE[contentType ?? ""] ?? ".bin"}`
}

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
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const messageId = resolveMessageIdOrRef(params.message_id)
  if (!messageId) {
    return Left(
      new UserError(
        `Unknown message ref "${params.message_id}". Refs come from scan_messages and last for the session — ` +
          `re-run the scan to refresh them.`,
      ),
    )
  }

  const listResult = await client.listAttachments(messageId)
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
    references
      .map((r) => `  ${r.name ?? "(unnamed)"}${r.isFolder ? " (folder)" : ""} — ${r.sourceUrl ?? "URL not returned"}`)
      .join("\n")

  if (all.length === 0) {
    return Left(
      new UserError(
        references.length === 0
          ? "This message has no attachments."
          : `This message has no downloadable file attachments, but it does carry ${references.length} cloud ` +
              `link${references.length === 1 ? "" : "s"} (reference attachment${references.length === 1 ? "" : "s"}). ` +
              `The mailbox holds no bytes for these — open the URL to get the content:\n${describeReferences()}`,
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
              `the mailbox, so there are no bytes to save. Open it directly: ${asReference.sourceUrl ?? "URL not returned by Graph"}`,
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
          `be saved from the mailbox — open the URL instead:\n${describeReferences()}`
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

  const response = await fetch(
    `${GRAPH_API_BASE}/v1.0/me/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(chosen.id)}/$value`,
    { headers: { Authorization: `Bearer ${token}` } },
  )
  if (!response.ok) return Left(await httpError(response))

  const buffer = Buffer.from(await response.arrayBuffer())

  const dir = params.out_dir ? (isAbsolute(params.out_dir) ? params.out_dir : resolve(params.out_dir)) : tmpdir()
  const filename = safeName(chosen.name, chosen.contentType, chosen.id)
  const target = join(dir, filename)

  // join() on a sanitised filename cannot escape dir, but assert it rather than trust the reasoning —
  // this is the one place in the server where remote input names a write path.
  if (!resolve(target).startsWith(resolve(dir))) {
    return Left(new UserError(`Refusing to write outside ${dir}.`))
  }

  try {
    await mkdir(dir, { recursive: true })
    await writeFile(target, buffer)
  } catch (err) {
    return Left(new UserError(`Failed to write ${target}: ${err instanceof Error ? err.message : String(err)}`))
  }

  const type = chosen.contentType ?? "application/octet-stream"
  return Right(
    `Saved: ${target}\n` +
      `Type: ${type}\n` +
      `Size: ${formatBytes(buffer.length)}\n\n` +
      `Read the file directly — PDFs and images do not need text extraction to be viewed.`,
  )
}
