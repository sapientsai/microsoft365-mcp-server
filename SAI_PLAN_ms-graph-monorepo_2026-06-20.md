# Master plan — Microsoft Graph MCP servers: layered monorepo

**Date:** 2026-06-20
**Status:** Agreed by both threads (driving = `microsoft365-mcp-server`; gateway = `microsoft-mcp-server`).
**This is the single master plan.** Inputs that feed it:
- `SAI_MEMO_ms-graph-consolidation-feedback_2026-06-20.md` (this repo) — gateway-side specifics & gotchas.
- `SAI_ANLY_microsoft-mcp-servers-consolidation_2026-06-20.md` (personal KB) — the analysis.
**Supersedes** the earlier "collapse to one server + `MS365_PRESETS=gateway` + archive specialist" position,
whose premise ("the gateway is experimental / not deployed") is false — **both servers are deployed and
working in similar capacities.**

---

## Decision

Keep **two deployable servers** in a pnpm monorepo over a shared core — **three packages**:

```
microsoft365-mcp-server/                 (this repo → monorepo root; mature, published, current deps)
├── pnpm-workspace.yaml                  (already a workspace-of-one — one line from a true monorepo)
└── packages/
    ├── core/                @sapientsai/ms-graph-core — PRIVATE/internal (not published, bundled by tsdown)
    │     request() · pagination/odata · upload (opaque-ticket, hardened) · upload-ticket ·
    │     GraphApiError · AuthStrategy { delegated, app-only } · doc-extraction · $batch · AI-Search client
    ├── microsoft365/        published `microsoft365-mcp-server` — ~73 delegated domain tools, on core
    └── graph/               minimal app-only headless server (re-homed gateway) — passthrough/batch/
                             extraction/AI-Search, optional read-only build profile; built on somamcp
```

Optionally also expose the gateway capabilities as `MS365_PRESETS=gateway` on the 365 server for
**delegated** users who want extraction / AI Search — a cheap additive bonus, **not** a replacement for
the second binary.

## Server-shell layer: somamcp (adopt bottom-up)

`somamcp` (your own published MCP framework — `createServer()`, FastMCP backend abstraction, transport,
telemetry/introspection, `/health`·`/info`·`/dashboard`, protected routes, gateway, feedback; same
toolchain: functype 1.4 / ts-builds 3.2 / pnpm 11 / Node 24 / FastMCP 4) is a **different layer** from
`core/`: it's the server shell ("how to be an MCP server"), where `core/` is Graph plumbing ("how to talk
to Graph"). They compose, they don't compete:

```
somamcp  →  @sapientsai/ms-graph-core  →  { microsoft365, graph }
```

**Adopt it bottom-up, not big-bang.** It does **not** reduce this migration's expensive work — the Graph
`AuthStrategy`, the tool-registry/presets system, and the 245-test port are all still ours (somamcp
provides *neither* a tool registry/presets *nor* Graph token management — its auth is a hook, not a token
manager). And retrofitting the live, published 365 server onto it is risk this migration shouldn't carry.
So:

- **`packages/graph` is built on `somamcp` from day one** (Phase 3) — greenfield, app-only, low-risk; the
  right place to prove the integration, and it inherits transport/telemetry/health/dashboard for free.
- **`packages/microsoft365` stays on its current FastMCP bootstrap** through this migration; it moves onto
  somamcp later as a **separate, parity-gated Phase 6**.

The rest of the plan is unchanged; somamcp simply slots in as the shell `graph` is built on and the
eventual target for `microsoft365`.

## Why two binaries (the security argument is load-bearing now)

Both servers being deployed flips attack-surface-by-construction from hypothetical to real:

- `packages/graph` runs **app-only / headless**. A binary that doesn't compile in `send_message` /
  `delete_event` cannot expose them by misconfiguration — least-privilege **by construction**.
- Folding it into the 70-tool binary would put those write tools one env-var flip away in the headless
  slot — least-privilege *by config* replacing *by construction*: a **regression** for a live autonomous
  deployment, not a simplification.
- Once `core/` exists (the expensive work, in every plan), a second thin server on top is cheap — a small
  entrypoint + a few tool registrations + the app-only `AuthStrategy` adapter.

> **Honest caveat to make the claim complete:** by-construction only removes the *named* write tools. The
> generic passthrough (`graph_query` / `microsoft_graph`) can still reach write endpoints
> (`POST /me/sendMail`, `DELETE …`). So the real least-privilege levers for `packages/graph` are
> **(a)** a read-only-scoped **app registration** (application permissions), and **(b)** optionally a
> read-only **build profile** that drops/guards the passthrough. Plan for both; don't treat "no named
> write tools" as the whole story.

## Guiding principles

1. **Modernize-on-arrival.** Ported code is rewritten *once*, as it lands in `core`/`graph` (functype 1.4
   subpaths, core graph-client, hardened upload). Never modernize `microsoft-mcp-server` in place.
2. **Behavior-preserving, test-gated.** The gateway's 245 tests are the parity contract; they travel with
   the code (re-expressed under functype 1.4) and must re-green before archive.
3. **Always green.** Each phase leaves both deployments building, tested, and shippable.

---

## Phase 0 — Safety net: the 245-test parity contract

The gateway repo has **245 tests / 13 spec files** (verified) covering exactly the ported surface:
`batch · extract-text · ai-search · ai-search-client · download · upload · sharepoint-search · site-cache ·
token-manager (app-only) · graph-client · errors · body-parameter · server`.

- Run them green and record as the **baseline contract**.
- **Fill thin spots** before porting: `token-manager` (6) — refresh-buffer expiry, `"common"` tenant
  rejection, api-key gate on `/upload` + transport; `site-cache` (8) — 1-hr TTL expiry + default-drive
  resolution; confirm extraction tests use **committed real fixtures** (.docx/.pdf/.xlsx, network-free).
- **Capture a tool-inventory snapshot** of both servers (name + params) for the final no-loss diff.
- **Nuance:** these are written against functype 0.57 + the gateway's structure. They are a parity
  *target*, re-expressed under functype 1.4 + `core/` during the port — same coverage, new imports/idioms,
  not a suite that runs unchanged.

Exit: old suite green; thin spots filled; fixtures portable; inventory snapshot saved.

## Phase 1 — Workspacify in place (plumbing only)

- Add `packages:` to `pnpm-workspace.yaml`. Move `src/`→`packages/microsoft365/src/`,
  `test/`→`packages/microsoft365/test/`.
- Preserve published name, `bin`, `exports`, `files`, Docker build, MCPB bundle, and the OIDC publish
  workflow (**Node 24** — npm 11.5.1+ for the OIDC handshake). Fix `manifest.json` / `.claude-plugin/`
  paths under `packages/`.
- **Riskiest plumbing step:** verify `pnpm validate` + a publish **dry-run** + `docker build` green from
  the new layout before proceeding.

Exit: 365 builds/tests/publishes (dry-run) identically from `packages/microsoft365`.

## Phase 2 — Extract `packages/core` + the `AuthStrategy` seam

- **Design `AuthStrategy` first** (see below) — the one genuine design task; the gate for everything.
- Lift shared plumbing into `core`: graph-client (`request()`/pagination/odata), upload + upload-ticket
  (hardened), `GraphApiError`, auth strategies. Build the **delegated adapter** first so the 365 server
  stays green. `packages/microsoft365` consumes `@sapientsai/ms-graph-core` via `workspace:*`.
- 365's existing 125 tests stay green unchanged.

Exit: `pnpm validate` green across packages; 365 tests unchanged and passing.

## Phase 3 — Re-home the gateway as `packages/graph` (on somamcp + core)

- Thin app-only server **built on `somamcp` (`createServer()`)** — inherits transport, telemetry,
  `/health`·`/info`·`/dashboard` for free — over `core` for Graph ops: passthrough (reuse `graph_query`),
  `$batch`, extraction, AI Search; the **app-only `AuthStrategy` adapter**. This greenfield server is the
  proving ground for the somamcp integration (low-risk: app-only, no oauth-proxy).
- Port the **three net-new capabilities** (`microsoft_graph_batch`, `read_document`, Azure AI Search) into
  `core`/`graph`, modernized on arrival.
- **Fix the upload-token leak in the process:** the gateway's `get_upload_config` currently returns a curl
  with `Authorization: Bearer ${config.apiKey}` (static `MCP_API_KEY` → transcript). Use `core`'s
  **opaque upload-ticket** pattern instead, and **rotate `MCP_API_KEY`** on the running deployment when
  the patch ships.
- Bring the 245 tests across (re-expressed under functype 1.4 + `core`) as the parity gate.

Exit: `packages/graph` re-greens the 245-test contract; tool-inventory diff shows no loss.

## Phase 4 — (Optional) gateway preset on the 365 server

- Add a `ToolDomain` `"gateway"` (or `"rag"`) to `tool-registry.ts`; register `microsoft_graph_batch`,
  `read_document`, AI Search under it; `PRESETS.gateway = ["query", "gateway"]`; document
  `MS365_PRESETS=gateway`. Cheap, since the capabilities already live in `core`.

## Phase 5 — Cut over & archive (only after parity)

- Deploy `packages/graph` in place of the gateway deployment; smoke-test; repoint; rotate `MCP_API_KEY`.
- **Archive** `sapientsai/microsoft-mcp-server` (do not delete) — **only after `packages/graph` is at test
  parity**.
- (Optional, separate, deliberate) rename the monorepo to a family name (`sapientsai/microsoft-mcp` /
  `ms-graph-mcp`) — touches OIDC trusted-publishing config + deploy webhooks.

## Phase 6 — (Separate, later) migrate `microsoft365` onto somamcp

Independent of cutover; do it only **after `packages/graph` has validated somamcp in production**, so the
mature, deployed, published server is never the guinea pig.

- **Spike first (hard gate).** Confirm somamcp can express everything 365 needs that `graph` doesn't
  exercise: the **oauth-proxy `AzureProvider`**, the Hono **`/upload` mount** (`getApp()` today), and the
  **MCPB bundle**. If any can't pass cleanly through somamcp's backend abstraction, fix it **in somamcp**
  before touching 365.
- Then rebase 365's entry layer onto `createServer()`, keeping the tool-registry/preset/filtering system
  on top (somamcp has none) and `core` for Graph ops. 365's 125 tests stay green; publish/Docker/MCPB
  dry-runs green.

Exit: 365 runs on somamcp at full feature + test parity → the whole fleet shares one server shell.

---

## `AuthStrategy` interface — the one hard design task

```ts
type AuthStrategy = {
  getAccessToken: () => Promise<Either<AuthError, string>>
  // + the hook the upload relay / token-context needs
}
```

- **Delegated adapter** (365): wraps oauth-proxy + `account-registry` + `token-context` +
  interactive/cert/secret/client-token + refresh. Build first.
- **App-only adapter** (graph): `client_credentials` against `login.microsoftonline.com`, `.default`
  scope, ~5-min refresh buffer, required tenant (reject `"common"`). Already supported here via cert/secret
  — this is *formalizing existing capability as a strategy*, not new auth code.

Design and review before Phase 2.

## Porting gotchas (from the gateway-side memo)

1. **🔴 `get_upload_config` secret leak — remediate (live bug).** Static `MCP_API_KEY` inlined into the
   curl → transcript (non-scrubbable). Adopt the opaque-ticket pattern in `packages/graph`; rotate the key
   on deploy. (Handled in Phase 3.)
2. **🟠 Rewrite functype 0.57 → 1.4 — don't paste.** Import style + `Either`/`Option`/`orThrow` differ;
   port logic, re-express in 1.4 + `core`.
3. **🟠 `read_document` deps are heavy** (`mammoth` + `unpdf` + `exceljs`). Scope them to `packages/graph`
   (and the optional 365 preset); lazy-load — do **not** bloat the default 365 install for users who never
   extract.
4. **🟡 Reconcile SharePoint search + `site-cache`, don't duplicate.** The gateway's app-only fan-out vs
   the 365 server's `search_site_files`. Keep only what the headless server needs in `core`/`graph`; don't
   fork.
5. **🟡 Namespace + identifiers.** Standardize `AZURE_*` / `GRAPH_*` / `SITE_URL` per package; **drop the
   hardcoded `DEFAULT_CLIENT_ID`**; keep every package free of customer-specific references (hostnames,
   tenant/drive IDs, principal names) — watch the ported AI Search + site-cache config paths.

## Verification

- **Per phase:** `pnpm validate` green in every package.
- **Contract:** the 245-test parity suite re-greens at Phases 3–4.
- **No loss:** tool-inventory diff vs the Phase-0 snapshot — zero dropped tools/params.
- **Smoke (end-to-end):** both deployables run; spot-check `read_document` on real DOCX/PDF/XLSX, a
  `graph_query` passthrough, an AI Search query, and a delegated mail/calendar op on 365.
- **Release safety:** 365 publish dry-run + Docker + MCPB builds succeed from the new layout.

## Risks & mitigations

- **Auth-strategy unification** → design the interface first (Phase 2 precondition).
- **Build/deploy path moves (Phase 1)** → isolate; require publish/docker dry-runs before proceeding.
- **functype 0.57→1.4 idiom gap** → modernize-on-arrival; the 245 tests catch regressions.
- **Extraction fixture portability** → committed real fixtures, network-free (Phase 0).
- **Live secret in transcript** → rotate `MCP_API_KEY` at cutover (Phases 3/5).
- **somamcp can't express 365's oauth-proxy / `/upload` mount / MCPB** → bottom-up adoption (prove on the
  greenfield `graph` server first) + the Phase 6 spike gate before 365 is touched.

## Out of scope / follow-ups

- Publishing `core` as a *public* package — only if a third external consumer appears (internal
  `workspace:*` for now; two private consumers don't trip rule-of-three).
- Repo rename (separate, deliberate — Phase 5 optional).
- CI guard failing on customer-specific strings in the public packages.
- Read-only build profile for `packages/graph` (least-privilege hardening; can follow Phase 3).
