# microsoft365-mcp monorepo

pnpm workspace for the Microsoft Graph MCP servers.

## Packages

- **[`packages/microsoft365`](packages/microsoft365)** — `microsoft365-mcp-server`, the published
  Microsoft 365 MCP server (mail, calendar, contacts, files, Teams, Planner, OneNote, To Do, and more
  via Microsoft Graph). See its [README](packages/microsoft365/README.md) and [DEPLOYMENT.md](DEPLOYMENT.md).

Planned (see [`SAI_PLAN_ms-graph-monorepo_2026-06-20.md`](SAI_PLAN_ms-graph-monorepo_2026-06-20.md)):
a shared `packages/core` and a minimal app-only `packages/graph` server.

## Development

```bash
pnpm install
pnpm validate   # format + lint + test + build across all packages
pnpm build      # build all packages
pnpm test       # test all packages
```

Per-package commands run inside the package dir (e.g. `pnpm --filter microsoft365-mcp-server build`).
