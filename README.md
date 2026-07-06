# microsoft365-mcp monorepo

pnpm workspace for the Microsoft Graph MCP servers. Descends from the archived
[`sapientsai/microsoft-mcp-server`](https://github.com/sapientsai/microsoft-mcp-server) — split into
**two focused servers over a shared core**. The app-only server keeps the original
`microsoft-mcp-server` name; the delegated server is `microsoft365-mcp-server`.

## Packages

- **[`packages/microsoft365`](packages/microsoft365)** — `microsoft365-mcp-server`, the delegated
  (OAuth-proxy) MS365 gateway: 40+ tools across mail, calendar, contacts, files, Teams, Planner,
  OneNote, To Do, and more via Microsoft Graph. Published to
  [npm](https://www.npmjs.com/package/microsoft365-mcp-server). See its
  [README](packages/microsoft365/README.md) and the monorepo [DEPLOYMENT.md](DEPLOYMENT.md).
- **[`packages/graph`](packages/graph)** — `microsoft-mcp-server` (the original name), a lean
  **app-only** (`client_credentials`) server purpose-built for headless document-RAG:
  `microsoft_graph` passthrough + `microsoft_graph_batch` + `read_document` + `sharepoint_search` +
  optional `azure_ai_search` + a `/upload` relay. Ships as a Docker image at
  `ghcr.io/sapientsai/microsoft-mcp-server`. See its [README](packages/graph/README.md).
- **[`packages/core`](packages/core)** — `@sapientsai/ms-graph-core`, the shared plumbing (upload
  helpers, ticket tokens, OData/pagination, error and auth-strategy types) used by both servers.
  Internal workspace package, not published.

Pick by auth model: use `packages/microsoft365` when a user is present to consent (interactive
OAuth); use `packages/graph` for tenant-wide, headless, app-only deployments.

## Development

```bash
pnpm install
pnpm validate   # format + lint + test + build across all packages
pnpm build      # build all packages
pnpm test       # test all packages
```

Per-package commands run inside the package dir, e.g.:

```bash
pnpm --filter microsoft365-mcp-server build
pnpm --filter microsoft-mcp-server dev
```

## History

Design context for the split lives in
[`SAI_PLAN_ms-graph-monorepo_2026-06-20.md`](SAI_PLAN_ms-graph-monorepo_2026-06-20.md).
