import {
  type AuthStrategy,
  decodeBase64Upload,
  describeFetchError,
  filenameFromPath,
  GRAPH_API_BASE,
  MAX_UPLOAD_SIZE,
  resolveUploadContentType,
  resolveUploadTicket,
  sessionUpload,
  SIMPLE_UPLOAD_LIMIT,
  simpleUpload,
} from "@sapientsai/ms-graph-core"
import { type Either, Left, Right } from "functype/either"

// Minimal structural shapes for the somamcp/FastMCP Hono app, so we don't take a direct
// hono dependency just for these signatures.
export type UploadRequestContext = {
  header: (name: string) => string | undefined
  query: (name: string) => string | undefined
  arrayBuffer: () => Promise<ArrayBuffer>
}
type HonoContext = { req: UploadRequestContext; json: (body: unknown, status: number) => unknown }
type HonoLike = {
  post: (path: string, h: (c: HonoContext) => unknown) => void
  put: (path: string, h: (c: HonoContext) => unknown) => void
}

// Authorize the CALLER to the relay. somamcp does not auto-protect a hand-mounted POST
// route (only GET artifacts), so we self-apply the check here. Accepts an opaque upload
// ticket (resolving to the api key) or the raw api key. Refuses (503) when no api key is
// configured — never serve a write endpoint backed by the server's own app-only token to
// an unauthenticated caller.
const authorizeCaller = (
  bearer: string | undefined,
  apiKey: string | undefined,
): { ok: true } | { ok: false; status: number; error: string } => {
  if (!apiKey) {
    return { ok: false, status: 503, error: "Upload endpoint is not configured for authentication (set MCP_API_KEY)." }
  }
  const provided = bearer ? (resolveUploadTicket(bearer) ?? bearer) : undefined
  if (provided !== apiKey) return { ok: false, status: 401, error: "Unauthorized" }
  return { ok: true }
}

export const handleUpload = async (
  req: UploadRequestContext,
  auth: AuthStrategy,
  apiKey: string | undefined,
): Promise<{ status: number; body: unknown }> => {
  const path = req.query("path")
  if (!path) return { status: 400, body: { error: "path query parameter is required" } }
  if (!/:\/content$/i.test(path)) return { status: 400, body: { error: 'path must end with ":/content"' } }

  const bearer = (req.header("authorization") ?? req.header("Authorization"))?.replace(/^Bearer\s+/i, "")
  const caller = authorizeCaller(bearer, apiKey)
  if (!caller.ok) return { status: caller.status, body: { error: caller.error } }

  // The relay uploads with the server's own app-only Graph token (never the caller's).
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

export const mountUploadRoute = (app: HonoLike, auth: AuthStrategy, apiKey: string | undefined): void => {
  const handler = async (c: HonoContext): Promise<unknown> => {
    try {
      const result = await handleUpload(c.req, auth, apiKey)
      return c.json(result.body, result.status)
    } catch (err) {
      const { message } = describeFetchError(err)
      console.error("[Upload] unhandled error:", message)
      return c.json({ error: message }, 500)
    }
  }
  app.post("/upload", handler)
  app.put("/upload", handler)
}
