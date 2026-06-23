import { type AuthStrategy, filenameFromPath, formatBytes, GRAPH_API_BASE } from "@sapientsai/ms-graph-core"
import { z } from "zod"

import { extractTextFromBuffer } from "../extract/extract"

const MAX_FILE_SIZE = 10 * 1024 * 1024 // 10 MB
const API_VERSIONS = ["v1.0", "beta"] as const

// read_document: download a Graph file (binary, so it bypasses the JSON request() layer)
// and return its extracted text. Ported from microsoft-mcp-server.
export const buildReadDocumentTool = (auth: AuthStrategy, fetchImpl: typeof fetch = fetch) => ({
  name: "read_document",
  description:
    "Download a file from SharePoint or OneDrive and return its readable text content. Supports DOCX, PDF, XLSX, " +
    "and text-based files. Use instead of microsoft_graph when you need document contents. Construct the path as " +
    "/drives/{driveId}/items/{itemId}/content (or /me/drive/items/{id}/content).",
  parameters: z.object({
    path: z.string().describe("Graph path to the file content endpoint, ending in /content"),
    api_version: z.enum(API_VERSIONS).default("v1.0").describe("Graph API version"),
    format: z.string().optional().describe("Optional conversion format (e.g. 'pdf'), for supported types only"),
    max_chars: z
      .number()
      .int()
      .min(1000)
      .max(200000)
      .default(50000)
      .describe("Max characters to return (1000-200000); content beyond is truncated"),
  }),
  execute: async (args: {
    path: string
    api_version: (typeof API_VERSIONS)[number]
    format?: string
    max_chars: number
  }): Promise<string> => {
    const tokenResult = await auth.getAccessToken()
    if (tokenResult.isLeft()) throw new Error((tokenResult.value as { message: string }).message)
    const token = tokenResult.value as string

    const query = args.format ? `?format=${args.format}` : ""
    const url = `${GRAPH_API_BASE}/${args.api_version}${args.path}${query}`
    const response = await fetchImpl(url, { method: "GET", headers: { Authorization: `Bearer ${token}` } })

    if (!response.ok) {
      const contentType = response.headers.get("content-type")
      if (contentType?.includes("application/json")) {
        const errorData = (await response.json()) as { error?: { message?: string } }
        throw new Error(errorData.error?.message ?? `HTTP ${response.status}: ${response.statusText}`)
      }
      throw new Error(`HTTP ${response.status}: ${response.statusText}`)
    }

    const contentType = response.headers.get("content-type") ?? "application/octet-stream"
    const filename = filenameFromPath(args.path) ?? "download"
    const buffer = Buffer.from(await response.arrayBuffer())

    if (buffer.length > MAX_FILE_SIZE) {
      throw new Error(`File too large (${formatBytes(buffer.length)}). Maximum is ${formatBytes(MAX_FILE_SIZE)}.`)
    }

    const extracted = await extractTextFromBuffer(buffer, contentType, filename)
    const fullText = extracted.fold(
      (error) => {
        throw new Error(error.message)
      },
      (text) => text,
    )

    const text =
      fullText.length > args.max_chars
        ? `${fullText.slice(0, args.max_chars)}\n\n[truncated at ${args.max_chars.toLocaleString()} chars — full document is ${fullText.length.toLocaleString()} chars]`
        : fullText

    return `File: ${filename} (${formatBytes(buffer.length)})\n\n${text}`
  },
})
