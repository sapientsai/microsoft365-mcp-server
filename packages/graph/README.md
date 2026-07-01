# @sapientsai/ms-graph-server

Minimal **app-only** Microsoft Graph MCP server, built on [somamcp](https://github.com/sapientsai/SomaMCP)
+ [`@sapientsai/ms-graph-core`](../core). The lean, headless counterpart to the delegated
`microsoft365-mcp-server` — it will replace `microsoft-mcp-server` at cutover.

- **Auth:** app-only (`client_credentials`) against a concrete tenant, via core's `AuthStrategy`.
- **Transport:** stdio or httpStream (somamcp). `/health`, `/info`, `/dashboard` come free from somamcp.
- **Server shell:** somamcp `createServer()` (telemetry/introspection included).

## Tools

| Tool | Notes |
|------|-------|
| `get_auth_status` | Reports whether an app-only token can be acquired (no token leak). |
| `microsoft_graph` | Generic Graph passthrough (any method/path/body, v1.0/beta). |
| `microsoft_graph_batch` | `$batch` — up to 20 requests, `dependsOn`. |
| `read_document` | Download + extract text (DOCX/PDF/XLSX/text), truncated to `max_chars`. |
| `sharepoint_search` | KQL document-library search via the Graph Search API (region-scoped). |
| `azure_ai_search` | *Optional* — only when `AZURE_AI_SEARCH_*` is configured. |
| `get_upload_config` | Returns a curl for `/upload`; Authorization is an opaque ticket, **not** the raw key. |
| `/upload` (HTTP) | Binary upload relay (POST/PUT) — self-applied `MCP_API_KEY` gate, uploads with the server's own app-only token. |

## Configuration

See [`.env.example`](.env.example). App-only auth (`MS_GRAPH_TENANT_ID`/`CLIENT_ID`/`CLIENT_SECRET`)
is required; `MCP_API_KEY` gates the transport + `/upload`; SharePoint search and Azure AI Search
have their own optional env.

## Parity with `microsoft-mcp-server` (the gateway it replaces)

All gateway capabilities are ported, with two deliberate scoping decisions:

- **SharePoint search:** only the Graph Search API path is ported (what app-only always uses; the
  region defaults to `NAM`). The gateway's drive-fan-out + site-cache path is unreachable for
  app-only and was intentionally not carried over.
- **`download_file`:** **not ported** — out of scope for this server's document-RAG purpose. It
  returns inline images / raw base64 / disk-saved binaries — a delegated/interactive UX feature
  that overlaps `read_document` (text extraction) and the `microsoft_graph` passthrough (raw GET).
  The somamcp-side blocker is resolved (`wrapTool` passes content-array/image returns through
  unchanged; `imageContent` is exported), so a later port is a straightforward scope call, not a
  dependency wait.
