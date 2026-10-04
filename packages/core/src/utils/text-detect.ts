import { extname } from "node:path"

// One answer to "is this text?" for download_file, upload_file and read_document. Each used to
// carry its own list, and they disagreed: a .log stored by OneDrive as application/octet-stream
// was shown by none of them, because all three trusted the MIME type alone.

const TEXT_MIME_TYPES: ReadonlySet<string> = new Set([
  "application/json",
  "application/xml",
  "application/javascript",
  "application/csv",
  "application/x-ndjson",
  "application/x-www-form-urlencoded",
])

// Extensions whose bytes may be text in a legacy encoding (Windows-1252 CSV exports, old logs).
// For these, no NUL bytes is enough; any other extension must also decode as UTF-8.
const TEXT_EXTENSIONS: ReadonlySet<string> = new Set([
  ".txt",
  ".log",
  ".md",
  ".csv",
  ".tsv",
  ".json",
  ".jsonl",
  ".ndjson",
  ".xml",
  ".html",
  ".htm",
  ".yaml",
  ".yml",
  ".ini",
  ".conf",
  ".vtt",
  ".srt",
])

const SNIFF_BYTES = 4096

const baseMimeType = (contentType?: string): string => (contentType ?? "").toLowerCase().split(";")[0]?.trim() ?? ""

export const isTextMimeType = (contentType?: string): boolean => {
  const base = baseMimeType(contentType)
  return base.startsWith("text/") || TEXT_MIME_TYPES.has(base) || base.endsWith("+json") || base.endsWith("+xml")
}

// OneDrive stores any extension it does not recognise as application/octet-stream, which says
// nothing about the bytes. Such files have to be judged by their content instead.
export const isUnknownMimeType = (contentType?: string): boolean => {
  const base = baseMimeType(contentType)
  return base === "" || base === "application/octet-stream"
}

const decodesAsUtf8 = (sample: Uint8Array): boolean => {
  try {
    // stream: true holds back a multi-byte character cut off at the end of the sample instead of
    // rejecting it, so only a genuinely invalid sequence fails.
    new TextDecoder("utf-8", { fatal: true }).decode(sample, { stream: true })
    return true
  } catch {
    return false
  }
}

// Judges the first 4 KB. Binary formats put NUL bytes there almost immediately, so a binary file
// renamed to .log is still caught.
export const looksLikeText = (bytes: Uint8Array, filename?: string): boolean => {
  const sample = bytes.subarray(0, SNIFF_BYTES)
  if (sample.includes(0)) return false
  if (filename && TEXT_EXTENSIONS.has(extname(filename).toLowerCase())) return true
  return decodesAsUtf8(sample)
}
