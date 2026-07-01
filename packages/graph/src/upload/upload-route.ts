import {
  type AuthStrategy,
  decodeBase64Upload,
  describeFetchError,
  filenameFromPath,
  GRAPH_API_BASE,
  MAX_UPLOAD_SIZE,
  resolveUploadContentType,
  sessionUpload,
  SIMPLE_UPLOAD_LIMIT,
  simpleUpload,
} from "@sapientsai/ms-graph-core"
import { type Either, Left, Right } from "functype/either"
import type { RouteConfig } from "somamcp"

// Minimal structural shape for the parts of Hono's request we read, so we don't take a
// direct hono dependency just for this signature (the real `c.req` satisfies it).
export type UploadRequestContext = {
  header: (name: string) => string | undefined
  query: (name: string) => string | undefined
  arrayBuffer: () => Promise<ArrayBuffer>
}

const jsonResponse = (body: unknown, status: number): Response =>
  new Response(JSON.stringify(body), { status, headers: { "content-type": "application/json" } })

// Upload mechanics only. Caller authorization is handled upstream by somamcp's `protected`
// route gate (the server's `authenticate` callback), so this no longer re-checks the key —
// it always uploads with the server's own app-only Graph token (never the caller's).
export const handleUpload = async (
  req: UploadRequestContext,
  auth: AuthStrategy,
): Promise<{ status: number; body: unknown }> => {
  const path = req.query("path")
  if (!path) return { status: 400, body: { error: "path query parameter is required" } }
  if (!/:\/content$/i.test(path)) return { status: 400, body: { error: 'path must end with ":/content"' } }

  const tokenResult = await auth.getAccessToken()
  if (tokenResult.isLeft()) {
    return { status: 401, body: { error: (tokenResult.value as { message: string }).message } }
  }
  const token = tokenResult.value as string

  const rawBufferResult = await (async (): Promise<Either<string, Buffer>> => {
    try {
      return Right(Buffer.from(await req.arrayBuffer()))
    } catch (error) {
      return Left(`Failed to read request body: ${describeFetchError(error).message}`)
    }
  })()
  if (rawBufferResult.isLeft()) return { status: 400, body: { error: rawBufferResult.value as string } }
  const rawBuffer = rawBufferResult.value as Buffer
  if (rawBuffer.length === 0) return { status: 400, body: { error: "Empty request body" } }

  const buffer = req.query("encoding") === "base64" ? decodeBase64Upload(rawBuffer) : rawBuffer
  if (buffer.length === 0) return { status: 400, body: { error: "Invalid base64 content" } }
  if (buffer.length > MAX_UPLOAD_SIZE)
    return { status: 413, body: { error: `File too large (max ${MAX_UPLOAD_SIZE} bytes)` } }

  const apiVersion = req.query("apiVersion") ?? "v1.0"
  const conflictBehavior = req.query("conflictBehavior") ?? "rename"
  const filename = filenameFromPath(path)
  const contentType = resolveUploadContentType(req.query("contentType"), filename)
  const apiBase = `${GRAPH_API_BASE}/${apiVersion}`

  const result =
    buffer.length <= SIMPLE_UPLOAD_LIMIT
      ? await simpleUpload(apiBase, path, token, buffer, contentType, conflictBehavior)
      : await sessionUpload(apiBase, path, token, buffer, conflictBehavior)

  return result.fold<{ status: number; body: unknown }>(
    (error) => ({ status: (error as { status?: number }).status ?? 500, body: { error: error.message } }),
    (item) => ({ status: 200, body: item }),
  )
}

// The binary upload relay as a first-class somamcp protected route. `protected: true` runs
// the server's `authenticate` gate before the handler (401 on a bad key / ticket). When no
// MCP_API_KEY is configured the server registers no `authenticate`, so the gate rejects every
// request — `onUnauthorized` surfaces that as a 503 "not configured" instead of a bare 401,
// never serving a write endpoint (backed by the server's app-only token) unauthenticated.
export const buildUploadRoute = (auth: AuthStrategy, apiKey: string | undefined): RouteConfig => ({
  method: ["POST", "PUT"],
  path: "/upload",
  protected: true,
  onUnauthorized: () =>
    apiKey
      ? jsonResponse({ error: "Unauthorized" }, 401)
      : jsonResponse({ error: "Upload endpoint is not configured for authentication (set MCP_API_KEY)." }, 503),
  handler: async (c) => {
    try {
      const req: UploadRequestContext = {
        header: (name) => c.req.header(name),
        query: (name) => c.req.query(name),
        arrayBuffer: () => c.req.arrayBuffer(),
      }
      const result = await handleUpload(req, auth)
      return jsonResponse(result.body, result.status)
    } catch (err) {
      const { message } = describeFetchError(err)
      console.error("[Upload] unhandled error:", message)
      return jsonResponse({ error: message }, 500)
    }
  },
})
