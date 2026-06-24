import { filenameFromPath, mintUploadTicket, resolveUploadContentType } from "@sapientsai/ms-graph-core"
import { z } from "zod"

// get_upload_config: returns a ready-to-run curl for the /upload relay. The Authorization
// value is an opaque, short-lived UPLOAD TICKET — never the raw MCP_API_KEY. (The gateway
// inlined the raw key into tool output, leaking it into the transcript; this closes that.)
export const buildGetUploadConfigTool = (publicBaseUrl: string, apiKey: string | undefined) => ({
  name: "get_upload_config",
  description:
    "Get an authenticated upload URL + curl command for uploading a binary file to OneDrive/SharePoint. Pipe " +
    "base64 file bytes to the returned URL via POST; the server decodes and streams to Microsoft Graph (chunked " +
    "session upload for >4 MB, up to 250 MB). The Authorization value is a short-lived opaque ticket, not a secret.",
  parameters: z.object({
    path: z
      .string()
      .describe('Graph drive path ending in ":/content", e.g. /me/drive/root:/Documents/report.docx:/content'),
    content_type: z.string().optional().describe("Explicit MIME type; otherwise inferred from the filename"),
    conflict_behavior: z.enum(["rename", "replace", "fail"]).default("rename").describe("On name collision"),
  }),
  // eslint-disable-next-line @typescript-eslint/require-await -- somamcp tool execute is async
  execute: async (args: {
    path: string
    content_type?: string
    conflict_behavior: "rename" | "replace" | "fail"
  }): Promise<string> => {
    if (!/:\/content$/i.test(args.path)) {
      throw new Error('path must end with ":/content" (e.g. /me/drive/root:/Documents/file.docx:/content)')
    }

    const filename = filenameFromPath(args.path)
    const contentType = resolveUploadContentType(args.content_type, filename)
    const query = new URLSearchParams({
      path: args.path,
      conflictBehavior: args.conflict_behavior,
      contentType,
      encoding: "base64",
    })
    const uploadUrl = `${publicBaseUrl}/upload?${query.toString()}`

    // Mint an opaque ticket that resolves (server-side) to the api key; the raw key never
    // travels in the curl / tool output.
    const ticket = apiKey ? mintUploadTicket(apiKey) : undefined
    const authHeader = ticket ? `Authorization: Bearer ${ticket}` : undefined

    const curl = [
      `base64 "{local_file_path}" | tr -d '\\n'`,
      "| curl -X POST",
      authHeader ? `-H "${authHeader}"` : undefined,
      `-H "Content-Type: text/plain"`,
      `--data-binary @-`,
      `"${uploadUrl}"`,
    ]
      .filter(Boolean)
      .join(" \\\n  ")

    return JSON.stringify(
      {
        uploadUrl,
        method: "POST",
        contentType,
        conflictBehavior: args.conflict_behavior,
        encoding: "base64",
        maxSize: "250 MB",
        ...(authHeader ? { authHeader } : {}),
        curl,
        notes: [
          "Pipe base64-encoded file bytes to uploadUrl via POST with --data-binary @-.",
          "The Authorization value is an opaque, short-lived upload ticket (~10 min TTL) — not a secret.",
          ...(authHeader ? [] : ["No MCP_API_KEY configured: the /upload endpoint will refuse requests (503)."]),
        ],
      },
      null,
      2,
    )
  },
})
