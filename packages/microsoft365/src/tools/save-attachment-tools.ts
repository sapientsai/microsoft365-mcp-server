import { mkdir, writeFile } from "node:fs/promises"
import { tmpdir } from "node:os"
import { basename, extname, isAbsolute, join, resolve } from "node:path"

import { formatBytes, resilientFetch } from "@sapientsai/ms-graph-core"
import { UserError } from "fastmcp"
import type { Either } from "functype/either"
import { Left, Right } from "functype/either"

import { getAccessToken } from "../auth"
import { GRAPH_API_BASE } from "../auth/scopes"
import { getGraphClient } from "../client/graph-client"
import { defaultVersion } from "../client/graph-version"
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

const ITEM_ATTACHMENT = "#microsoft.graph.itemAttachment"

// Windows refuses these as a base name whatever follows the first dot, so "CON.pdf" fails too.
const WINDOWS_RESERVED = /^(con|prn|aux|nul|com[0-9]|lpt[0-9])$/i

// Most filesystems cap a name at 255 bytes, not characters. This leaves room for the " (99)" that
// writeWithoutOverwriting may add.
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

const MAX_NAME_ATTEMPTS = 100

// "wx" fails rather than replacing an existing file, so saving never destroys anything already on
// disk — including an earlier save of a same-named attachment from another message. That is what
// lets this tool be advertised as additive.
const numbered = (target: string, attempt: number): string => {
  if (attempt === 0) return target
  const ext = extname(target)
  return join(resolve(target, ".."), `${basename(target, ext)} (${attempt})${ext}`)
}

const writeWithoutOverwriting = async (target: string, data: Buffer, attempt = 0): Promise<Either<Error, string>> => {
  if (attempt >= MAX_NAME_ATTEMPTS) {
    return Left(new Error(`${MAX_NAME_ATTEMPTS} files named like ${basename(target)} already exist`))
  }
  const candidate = numbered(target, attempt)
  try {
    // The default folder is tmpdir(), which other users on the machine can list.
    await writeFile(candidate, data, { flag: "wx", mode: 0o600 })
    return Right(candidate)
  } catch (err) {
    if ((err as NodeJS.ErrnoException).code === "EEXIST") return writeWithoutOverwriting(target, data, attempt + 1)
    return Left(err instanceof Error ? err : new Error(String(err)))
  }
}

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
  const all = ((listResult.value as { value?: ReadonlyArray<GraphAttachment> }).value ?? []).filter(
    (a) => a["@odata.type"] !== "#microsoft.graph.referenceAttachment",
  )

  if (all.length === 0) return Left(new UserError("This message has no downloadable file attachments."))

  // No attachment_id: only unambiguous when there is exactly one. Guessing which of several a caller
  // meant is worse than making them look.
  const chosen = params.attachment_id
    ? all.find((a) => a.id === params.attachment_id)
    : all.length === 1
      ? all[0]
      : undefined

  if (!chosen) {
    if (params.attachment_id) {
      return Left(new UserError(`No attachment "${params.attachment_id}" on this message. Use list_attachments.`))
    }
    const names = all.map((a) => `  ${a.id}  ${a.name ?? "(unnamed)"} (${formatBytes(a.size ?? 0)})`).join("\n")
    return Left(
      new UserError(`This message has ${all.length} attachments — pass attachment_id to choose one:\n${names}`),
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
    `${GRAPH_API_BASE}/${defaultVersion()}/me/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(chosen.id)}/$value`,
    { headers: { Authorization: `Bearer ${token}` } },
  )
  if (!response.ok) return Left(await httpError(response))

  const buffer = Buffer.from(await response.arrayBuffer())

  const dir = params.out_dir ? (isAbsolute(params.out_dir) ? params.out_dir : resolve(params.out_dir)) : tmpdir()
  const filename = safeName(chosen)
  const target = join(dir, filename)

  // join() on a sanitised filename cannot escape dir, but assert it rather than trust the reasoning —
  // this is the one place in the server where remote input names a write path.
  if (!resolve(target).startsWith(resolve(dir))) {
    return Left(new UserError(`Refusing to write outside ${dir}.`))
  }

  const written = await mkdir(dir, { recursive: true }).then(
    () => writeWithoutOverwriting(target, buffer),
    (err: unknown) => Left<Error, string>(err instanceof Error ? err : new Error(String(err))),
  )
  if (written.isLeft()) return Left(new UserError(`Failed to write ${target}: ${written.value.message}`))
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
