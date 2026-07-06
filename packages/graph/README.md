# microsoft-mcp-server

**App-only** (`client_credentials`) Microsoft Graph MCP server for headless, tenant-wide
document-RAG — built on [somamcp](https://github.com/sapientsai/SomaMCP) +
[`@sapientsai/ms-graph-core`](../core). This is the successor to the archived
[`sapientsai/microsoft-mcp-server`](https://github.com/sapientsai/microsoft-mcp-server),
rebuilt in this monorepo over the shared core.

## Which server? (this vs `microsoft365-mcp-server`)

The monorepo ships **two** MCP servers over the same [`@sapientsai/ms-graph-core`](../core) — the
overlap was large, so they share roots but differ by auth model and purpose:

|              | `microsoft-mcp-server` (this package)                                       | [`microsoft365-mcp-server`](../microsoft365)                           |
| ------------ | --------------------------------------------------------------------------- | ---------------------------------------------------------------------- |
| **Auth**     | **app-only** (`client_credentials`) — no user present                       | **delegated** (per-user OAuth)                                         |
| **Use case** | headless, tenant-wide document-RAG / automation                             | interactive Microsoft 365 assistant                                    |
| **Surface**  | lean: `microsoft_graph` + `read_document` + `sharepoint_search` + `/upload` | full: 70+ tools across mail, calendar, files, Teams, Planner, OneNote… |
| **Deploy**   | Docker image (`ghcr.io/sapientsai/microsoft-mcp-server`)                    | npm (`microsoft365-mcp-server`)                                        |

Pick **this** one when there's **no user to consent** (a service/tenant credential); pick
`microsoft365-mcp-server` when a user is present to sign in.

- **Auth:** app-only (`client_credentials`) against a concrete tenant, via core's `AuthStrategy`.
- **Transport:** stdio or httpStream (somamcp). `/health`, `/info`, `/dashboard` come free from somamcp.
- **Server shell:** somamcp `createServer()` (telemetry/introspection included).

## Tools

| Tool                    | Notes                                                                                                                                            |
| ----------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------ |
| `get_auth_status`       | Reports whether an app-only token can be acquired (no token leak).                                                                               |
| `microsoft_graph`       | Generic Graph passthrough (any method/path/body, v1.0/beta).                                                                                     |
| `microsoft_graph_batch` | `$batch` — up to 20 requests, `dependsOn`.                                                                                                       |
| `read_document`         | Download + extract text (DOCX/PDF/XLSX/text), truncated to `max_chars`.                                                                          |
| `sharepoint_search`     | KQL document-library search via the Graph Search API (region-scoped).                                                                            |
| `azure_ai_search`       | _Optional_ — only when `AZURE_AI_SEARCH_*` is configured.                                                                                        |
| `get_upload_config`     | Returns a curl for `/upload`; Authorization is an opaque ticket, **not** the raw key.                                                            |
| `/upload` (HTTP)        | Binary upload relay (POST/PUT) — a somamcp `protected` route behind the shared `MCP_API_KEY` gate, uploads with the server's own app-only token. |

## Configuration

See [`.env.example`](.env.example). App-only auth (`MS_GRAPH_TENANT_ID`/`CLIENT_ID`/`CLIENT_SECRET`)
is required; `MCP_API_KEY` gates the transport + `/upload`; SharePoint search and Azure AI Search
have their own optional env.

## Deployment

- **Docker image:** `ghcr.io/sapientsai/microsoft-mcp-server` (published by CI on push to `main` and on
  `v*` tags — see [`.github/workflows/docker-graph.yml`](../../.github/workflows/docker-graph.yml)).
- **Dockerfile:** [`packages/graph/Dockerfile`](./Dockerfile) — multi-stage, `linux/amd64` +
  `linux/arm64`, health check on `/ping`.
- **Local run:** from the monorepo root,
  `docker compose -f packages/graph/docker-compose.yml up --build` (env from your shell / a `.env`
  at the invocation directory).

## Parity with the archived predecessor

This package carries forward the [archived
`sapientsai/microsoft-mcp-server`](https://github.com/sapientsai/microsoft-mcp-server) name and its
app-only capabilities. All of the old gateway's capabilities are ported, with two deliberate scoping
decisions:

- **SharePoint search:** only the Graph Search API path is ported (what app-only always uses; the
  region defaults to `NAM`). The gateway's drive-fan-out + site-cache path is unreachable for
  app-only and was intentionally not carried over.
- **`download_file`:** **not ported** — out of scope for this server's document-RAG purpose. It
  returns inline images / raw base64 / disk-saved binaries — a delegated/interactive UX feature
  that overlaps `read_document` (text extraction) and the `microsoft_graph` passthrough (raw GET).
  The somamcp-side blocker is resolved (`wrapTool` passes content-array/image returns through
  unchanged; `imageContent` is exported), so a later port is a straightforward scope call, not a
  dependency wait.
