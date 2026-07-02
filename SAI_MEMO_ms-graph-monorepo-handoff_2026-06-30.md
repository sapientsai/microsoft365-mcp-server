# Handoff — ms-graph monorepo consolidation (2026-06-30)

Checkpoint for continuing the `microsoft365-mcp-server` → layered monorepo work after a compact.
**Master plan:** `SAI_PLAN_ms-graph-monorepo_2026-06-20.md` (in this repo). **Analysis:**
`SAI_ANLY_microsoft-mcp-servers-consolidation_2026-06-20.md` (personal KB). **Gateway input:**
`SAI_MEMO_ms-graph-consolidation-feedback_2026-06-20.md` (this repo).

## What this is

Consolidating two deployed Microsoft Graph MCP servers into a **pnpm monorepo over a shared core**,
keeping **two products** (delegated + app-only). Deployments:
- `microsoft365-mcp-server` (this repo, delegated/per-user) → `ms365.civala.ai`.
- `microsoft-mcp-server` (separate repo, app-only) → `ms-mcp-central.civala.ai` — **being replaced** by
  the new `packages/graph`.

## Current state — `main` @ latest is a 3-package monorepo

```
microsoft365-mcp-server/                (repo root = private workspace)
├── packages/core/          @sapientsai/ms-graph-core (private) — AuthStrategy seam, createGraphRequest
│                           (request/pagination/odata/mapHttpError), upload + upload-ticket, extract-free
│                           generic types (GraphApiError, GraphApiVersion, ODataResponse, GraphDriveItem),
│                           GRAPH_API_BASE, upload-helpers. Bundled into microsoft365's dist by tsdown.
├── packages/microsoft365/  published `microsoft365-mcp-server` — the delegated server, on core.
└── packages/graph/         @sapientsai/ms-graph-server (private, 0.0.0) — app-only server on somamcp+core.
                            Takes the `microsoft-mcp-server` npm identity at Phase-5 cutover.
```

**Tests:** core 27 + graph 65 + microsoft365 120 = **212**, all green. `pnpm validate` runs all packages.

## Phase progress

- ✅ **Phase 1** (#33) — workspacify.
- ✅ **Phase 2a** (#35) — extract core (auth-free plumbing + AuthStrategy).
- ✅ **Phase 2b** (#36) — invert graph-client onto AuthStrategy.
- ✅ **Phase 3** — packages/graph on somamcp+core. All gateway capabilities ported:
  - step 1 (#43) core `createGraphRequest`; step 2 (#44) `microsoft_graph` + `$batch`;
    step 3 (#45) `read_document` (mammoth/unpdf/exceljs); step 4 (#46) `azure_ai_search`;
    step 5 (#47) `/upload` + `get_upload_config` (leak fixed); step 5b (#48) `sharepoint_search`.
- ✅ Security fixes (#34): /upload gate, loopback bind, oauth key separation.
- ✅ somamcp spike (GO) + improvement spec/issue (see below).
- ✅ Step 6 part 2 (#50): Dockerfile + Docker publish workflow for `packages/graph`.
- ✅ somamcp 1.1.0 adoption (#52): `/upload` is now a somamcp `protected` `addRoute`;
  deleted `authorizeCaller`/`mountUploadRoute`/`extractAuthHeader`; `getRequestHeader` +
  new `authorizesWithApiKey` gate (`src/auth/api-key-gate.ts`) that resolves upload tickets.

## REMAINING WORK (priority order)

1. **Give `packages/graph` an npm identity + publish wiring at cutover** (currently private/0.0.0).
   Docker image + publish workflow already landed (#50).
2. **Phase 5 cutover/archive:** deploy `packages/graph` in place of `ms-mcp-central`, verify parity
   against the running deployment, then **archive** `sapientsai/microsoft-mcp-server` (don't delete).
3. **Deploy-side (parked, user action on `ms365.civala.ai`):** set `MS365_JWT_SIGNING_KEY` +
   `MS365_TOKEN_ENCRYPTION_KEY` (independent secrets) to activate the #34 key-separation; one-time
   re-auth window. Enforce fail-fast once set.

## DECIDED (no action)

- **`download_file`: NOT ported** — out of scope for the app-only RAG server (overlaps `read_document` +
  the `microsoft_graph` passthrough). The somamcp blocker was a false alarm (content-array/image returns
  pass through `wrapTool` unchanged; `imageContent` already exported), so it's a pure scope call, not a
  dependency. See `packages/graph/README.md` parity section.
- **SharePoint site-cache fan-out: NOT ported** — unreachable for app-only (always uses the Search API
  with a region, default `NAM`). Only the Search-API path was ported.

## Key facts / gotchas (don't relearn these)

- **functype pinned to `1.4.4`** via pnpm `overrides` in `pnpm-workspace.yaml`. somamcp pulls 1.4.4;
  our packages ask `^1.4.3`. TWO copies of functype's recursive `Either` type blow tsc's budget
  (**TS2589**). One version fixes it. Also: annotate `.fold<T>(...)` / `.request<unknown>` explicitly at
  call sites to keep Either inference shallow (pattern used throughout `packages/graph`).
- **core is bundled** into microsoft365's published `dist/` by tsdown because it's a **devDependency**
  (`workspace:*`) — verified no `@sapientsai/ms-graph-core` in dist and clean published deps. The
  cosmetic `workspace:*` in published devDependencies is npm-tolerated (consumers don't install devDeps).
- **Docker** is workspace-aware: `pnpm -r build` (core before microsoft365) then
  `pnpm --filter microsoft365-mcp-server deploy --prod /prod`. `injectWorkspacePackages: true` required.
  `.dockerignore` excludes `**/node_modules`/`**/dist`. Container sets `FASTMCP_HOST=0.0.0.0`.
- **Publishing:** CI publishes microsoft365 to npm **on tag push** (`v*`) via OIDC trusted publishing.
  `.nvmrc` = node 24 (required for OIDC). **NEVER run `npm publish`.** Release = `vbctp` skill (validate,
  `npm version`, push --follow-tags).
- **somamcp** (`somamcp@^1.1.0`, Jordan's own): app-only `graph` server's shell via `createServer()`.
  `/upload` is a first-class `server.addRoute({ protected: true })` behind the shared `authenticate`
  gate (no more `getApp()` self-applied auth); `getRequestHeader` normalizes the request shape. Pinned
  past pnpm's minimum-release-age gate via `minimumReleaseAgeExclude` in `pnpm-workspace.yaml`. fastmcp ^4.3.0.
- **Local MCP config `.mcp.json`** (runs the delegated microsoft365 server inside Claude Code; NOT the
  hosted `ms365.civala.ai` deploy, which sets its own Docker env):
  - `args` must be `./packages/microsoft365/dist/index.js` — the monorepo move left the old `./dist/index.js`
    path stale; running it crashes at startup (`ERR_MODULE_NOT_FOUND: dotenv`, no root runtime node_modules),
    which Claude Code surfaces as MCP `-32000`. Fixed `618f144`; stale root `dist/` (gitignored) removed.
  - `MS365_ORG_MODE` defaulted to `true` via `${MS365_ORG_MODE:-true}` (fix `413cdc2`). The chats/teams/
    groups/planner/sites/list_users tools are `orgOnly` (tool-registry.ts) and hidden unless org mode is on —
    that's why `send_chat_message` (self-chat via magic id **`48:notes`**) wasn't registered. The npm package's
    own default stays `false` (opt-in); override with `MS365_ORG_MODE=false`. After any env change, **reconnect
    the MCP** (restart) — env only applies at server start.

## somamcp improvement spec (separate repo `~/IdeaProjects/somamcp`)

✅ **DONE — closed end to end.** All 5 spec items shipped in **somamcp 1.1.0** (PR somamcp#35, commit
`7348ef1`), each with tests: #1 `addRoute` (protected routes), #2(b) `getRequestHeader`, #3 content-array/
image passthrough (`wrapTool` test), #4 httpStream widening (cors/stateless/eventStore/SSL), #5
`examples/protected-upload-server`. Adopted downstream here (#52). Tracking **issue somamcp#34 is closed**.
Spec: `~/IdeaProjects/somamcp/SAI_SPEC_somamcp-improvements_2026-06-30.md`. `download_file` stays a scope
call (not a dependency wait) — #3 confirmed content-array/image returns pass through.

## Workflow norms (this effort)

- One PR per step; squash-merge + delete branch; watch CI (`gh pr checks N --watch`) to green before merge.
- **Branch from `main` first** — I twice accidentally committed to local `main`; recover via
  `git branch <name> && git reset --hard origin/main && git checkout <name>`.
- Commit messages end `Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>`; PR bodies end with the
  Claude Code line. Commit/push/merge only when the user asks.
