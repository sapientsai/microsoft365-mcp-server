import { extname } from "node:path"

import { isTextMimeType, isUnknownMimeType, looksLikeText } from "@sapientsai/ms-graph-core"
import ExcelJS from "exceljs"
import { type Either, Left, Right } from "functype/either"
import mammoth from "mammoth"
import { extractText as extractPdfText, getDocumentProxy } from "unpdf"

// Binary document → text extraction. Heavy deps (mammoth/unpdf/exceljs) live here rather than in
// either server package, so the delegated server can reach them through a lazy import() and keep
// them off its startup path. This module knows nothing about Microsoft Graph — it takes a buffer
// and a content type and returns text — hence its own ExtractError rather than core's GraphApiError.

export type ExtractError = { readonly type: "parse" | "unsupported"; readonly message: string }

// Binary document formats, for files OneDrive stored as application/octet-stream. Text formats are
// deliberately absent: an extension cannot tell a real .log from a binary renamed to .log, so text
// is decided by looksLikeText on the bytes.
const DOCUMENT_TYPES: Record<string, string> = {
  ".pdf": "application/pdf",
  ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ".doc": "application/msword",
  ".xls": "application/vnd.ms-excel",
}

export const EXTRACTABLE_TYPES = [
  "application/pdf",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
] as const

export const resolveContentType = (contentType: string, filename: string): string => {
  const lower = contentType.toLowerCase()
  if (isUnknownMimeType(lower)) return DOCUMENT_TYPES[extname(filename).toLowerCase()] ?? lower
  return lower
}

const parseError = (message: string): ExtractError => ({ type: "parse", message })
const unsupportedError = (message: string): ExtractError => ({ type: "unsupported", message })

const extractPdf = async (buffer: Buffer): Promise<Either<ExtractError, string>> => {
  try {
    const pdf = await getDocumentProxy(new Uint8Array(buffer))
    try {
      const { totalPages, text } = await extractPdfText(pdf, { mergePages: true })
      return Right(`[PDF: ${totalPages} page${totalPages === 1 ? "" : "s"}]\n\n${text}`)
    } finally {
      await pdf.loadingTask.destroy()
    }
  } catch (err) {
    return Left(parseError(err instanceof Error ? err.message : "PDF extraction failed"))
  }
}

const extractDocx = async (buffer: Buffer): Promise<Either<ExtractError, string>> => {
  try {
    const result = await mammoth.extractRawText({ buffer })
    return Right(result.value)
  } catch (err) {
    return Left(parseError(err instanceof Error ? err.message : "DOCX extraction failed"))
  }
}

const extractXlsx = async (buffer: Buffer): Promise<Either<ExtractError, string>> => {
  try {
    const wb = new ExcelJS.Workbook()
    await wb.xlsx.load(buffer as unknown as ArrayBuffer)
    const parts: string[] = []
    wb.eachSheet((ws) => {
      const rows: string[] = []
      ws.eachRow((row) => {
        const cells = Array.isArray(row.values) ? row.values.slice(1) : []
        rows.push(cells.map((v) => (v == null ? "" : String(v))).join(","))
      })
      const csv = rows.join("\n")
      parts.push(wb.worksheets.length > 1 ? `[Sheet: ${ws.name}]\n${csv}` : csv)
    })
    return Right(parts.join("\n\n"))
  } catch (err) {
    return Left(parseError(err instanceof Error ? err.message : "XLSX extraction failed"))
  }
}

export const extractTextFromBuffer = async (
  buffer: Buffer,
  contentType: string,
  filename: string,
): Promise<Either<ExtractError, string>> => {
  const resolved = resolveContentType(contentType, filename)

  if (isTextMimeType(resolved)) return Right(buffer.toString("utf-8"))
  if (resolved === "application/pdf") return extractPdf(buffer)
  if (resolved === "application/vnd.openxmlformats-officedocument.wordprocessingml.document") return extractDocx(buffer)
  if (resolved === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") return extractXlsx(buffer)
  if (isUnknownMimeType(resolved) && looksLikeText(buffer, filename)) return Right(buffer.toString("utf-8"))

  const supported = [...EXTRACTABLE_TYPES, "text/*"].join(", ")
  return Left(
    unsupportedError(`Unsupported content type "${contentType}" for text extraction. Supported: ${supported}`),
  )
}
