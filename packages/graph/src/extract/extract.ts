import { extname } from "node:path"

import type { GraphApiError } from "@sapientsai/ms-graph-core"
import ExcelJS from "exceljs"
import { type Either, Left, Right } from "functype/either"
import mammoth from "mammoth"
import { extractText as extractPdfText, getDocumentProxy } from "unpdf"

// Binary document → text extraction. Heavy deps (mammoth/unpdf/exceljs) are scoped to this
// package, never the lean delegated server. Ported from microsoft-mcp-server/src/download/
// extract.ts onto core's GraphApiError (extraction failures → type "parse").

export const CONTENT_TYPE_MAP: Record<string, string> = {
  ".pdf": "application/pdf",
  ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ".doc": "application/msword",
  ".xls": "application/vnd.ms-excel",
  ".txt": "text/plain",
  ".csv": "text/csv",
  ".json": "application/json",
  ".xml": "application/xml",
  ".html": "text/html",
  ".htm": "text/html",
}

export const EXTRACTABLE_TYPES = [
  "application/pdf",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
] as const

const TEXT_MIME_PREFIXES = ["text/", "application/json", "application/xml", "application/csv"]
const TEXT_MIME_SUFFIXES = ["+xml", "+json"]

export const isTextContent = (contentType: string): boolean => {
  const lower = contentType.toLowerCase()
  return (
    TEXT_MIME_PREFIXES.some((prefix) => lower.startsWith(prefix)) ||
    TEXT_MIME_SUFFIXES.some((suffix) => lower.includes(suffix))
  )
}

export const resolveContentType = (contentType: string, filename: string): string => {
  const lower = contentType.toLowerCase()
  if (lower === "application/octet-stream" || lower === "") {
    const ext = extname(filename).toLowerCase()
    return CONTENT_TYPE_MAP[ext] ?? contentType
  }
  return lower
}

const parseError = (message: string): GraphApiError => ({ type: "parse", message })

const extractPdf = async (buffer: Buffer): Promise<Either<GraphApiError, string>> => {
  try {
    const pdf = await getDocumentProxy(new Uint8Array(buffer))
    try {
      const { totalPages, text } = await extractPdfText(pdf, { mergePages: true })
      return Right(`[PDF: ${totalPages} page${totalPages === 1 ? "" : "s"}]\n\n${text}`)
    } finally {
      await pdf.destroy()
    }
  } catch (err) {
    return Left(parseError(err instanceof Error ? err.message : "PDF extraction failed"))
  }
}

const extractDocx = async (buffer: Buffer): Promise<Either<GraphApiError, string>> => {
  try {
    const result = await mammoth.extractRawText({ buffer })
    return Right(result.value)
  } catch (err) {
    return Left(parseError(err instanceof Error ? err.message : "DOCX extraction failed"))
  }
}

const extractXlsx = async (buffer: Buffer): Promise<Either<GraphApiError, string>> => {
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
): Promise<Either<GraphApiError, string>> => {
  const resolved = resolveContentType(contentType, filename)

  if (isTextContent(resolved)) return Right(buffer.toString("utf-8"))
  if (resolved === "application/pdf") return extractPdf(buffer)
  if (resolved === "application/vnd.openxmlformats-officedocument.wordprocessingml.document") return extractDocx(buffer)
  if (resolved === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") return extractXlsx(buffer)

  const supported = [...EXTRACTABLE_TYPES, "text/*"].join(", ")
  return Left(parseError(`Unsupported content type "${contentType}" for text extraction. Supported: ${supported}`))
}
