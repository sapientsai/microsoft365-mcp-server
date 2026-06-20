# Feedback — Microsoft Graph MCP consolidation (gateway-side input)

**Date:** 2026-06-20
**For:** the thread driving consolidation in `microsoft365-mcp-server`
**Scope:** corrected position + the gateway-repo specifics the migration must carry. The
single merged master plan is owned by the driving thread; this is the input from the gateway
side.

---

## Correction (this supersedes the earlier "collapse to one server" position)

The earlier recommendation to fold the gateway into a preset and archive its repo was built
on the premise that **the gateway is experimental / not deployed.** That premise is **false** —
**both servers are deployed and working in similar capacities.** With the true facts:

**Keep two deployable servers. Three packages.**

```
ms-graph-mcp/
├── packages/
│   ├── core/                      # @sapientsai/ms-graph-core — PRIVATE/internal (not published)
│   │   └── request() · upload (opaque-ticket) · GraphApiError · AuthStrategy{delegated, app-only}
│   ├── microsoft365-mcp-server/   # delegated/per-user, published, ~70 tools
│   └── graph/                     # minimal app-only headless server (the re-homed gateway)
```

Optionally also expose the gateway capabilities as `MS365_PRESETS=gateway` on the 365 server
for **delegated** users who want extraction / AI Search — that's a cheap additive bonus, **not**
a replacement for the second binary.

## Why two binaries, now that both are deployed

The deciding reason is **attack-surface by construction**, which only became load-bearing once
the headless deployment was confirmed real:

- The `graph` server runs **app-only / headless**. A binary that physically contains only
  passthrough + batch + extraction + search **cannot** expose `send_message` / `delete_event` —
  the code isn't there.
- Folding it into the 70-tool binary would put those write tools one env-var flip away in the
  headless slot — least-privilege-*by-config* replacing least-privilege-*by-construction*. That
  is a **regression** for a live autonomous deployment, not a simplification.
- Once `core/` exists (the expensive work, present in every plan), a second thin server on top
  of it is cheap: a small entrypoint + a few tool registrations + the app-only `AuthStrategy`
  adapter. "One server is simpler" barely applies — it mostly just deletes a product you run.

## Adopt from the driving plan (no conflict)

pnpm monorepo; `microsoft365-mcp-server` repo as root (mature, published, current deps,
workspace-of-one already); **private** `core/` (internal workspace dep for two consumers — does
**not** trip rule-of-three, since nothing is published); Node 24 for the OIDC publish path; the
manifest path churn; only **three net-new capabilities** to port (`microsoft_graph_batch`,
`read_document`, Azure AI Search — the passthrough already exists here as `graph_query`).

## The migration contract: the gateway's 245 tests

The gateway repo has **245 test cases across 13 spec files** (verified), covering exactly the
ported surface:

```
batch · extract-text · ai-search · ai-search-client · download · upload · sharepoint-search
site-cache · token-manager (app-only) · graph-client · errors · body-parameter · server
```

These travel with the code as the parity gate. **Do not archive `microsoft-mcp-server` until
`packages/graph` re-greens them.**

> Nuance: they're written against functype **0.57** and the gateway's current structure. After
> porting to functype 1.4 + `core/`, they must be re-expressed and re-greened — same coverage,
> new imports/idioms. They are a parity *target*, not a suite that runs unchanged.

## Be careful of these in `microsoft-mcp-server` when re-homing it

1. **🔴 `get_upload_config` leaks a secret into the transcript — fix it (live bug).** In the
   gateway's `src/index.ts` (~L485–522), `get_upload_config` returns a curl string containing
   `Authorization: Bearer ${config.apiKey}` — the static `MCP_API_KEY` lands in tool output →
   conversation transcript (non-scrubbable, ~30-day API-log retention). Because the gateway
   *survives* as `packages/graph`, this is a remediation, not a "don't copy": adopt this
   server's **opaque upload-ticket** pattern in `packages/graph`, and **rotate `MCP_API_KEY`**
   on the running deployment when the patch ships.

2. **🟠 Rewrite functype 0.57 → 1.4 idioms — don't paste.** Import style and
   `Either`/`Option`/`orThrow` usage won't match. Port logic, re-express in 1.4 + `core/`.

3. **🟠 `read_document` drags in `mammoth` + `unpdf` + `exceljs` (heavy).** These belong to the
   `graph` package (and to the optional 365 preset). Don't let them bloat the default 365
   install for users who never touch extraction — lazy-load or scope to the package/preset.

4. **🟡 Reconcile SharePoint search + `site-cache`, don't duplicate.** The gateway's
   `buildSearchTool` + `site-cache.ts` fan-out is tuned for app-only. Keep what the headless
   server genuinely needs in `core/`/`graph`; don't fork the 365 server's `search_site_files`.

5. **🟡 Namespace + identifiers.** Gateway uses `AZURE_*` / `GRAPH_*` / `SITE_URL`; standardize
   per package. Drop the hardcoded `DEFAULT_CLIENT_ID` (`cf7d1f97-…`). Per the driving plan's
   naming note, keep every package free of customer-specific references (hostnames, tenant/drive
   IDs, principal names) — watch the ported AI Search + site-cache config paths especially.

## Sequencing note

`core/` extraction is the gate for everything. Suggested order: (1) stand up workspace + move
365 into `packages/`; (2) extract `core/` with the `AuthStrategy` seam (delegated adapter
first, 365 stays green); (3) re-home the gateway as `packages/graph` on `core/` with the
app-only adapter, porting the 245 tests as the parity contract and fixing the upload-token bug
in the process; (4) optional `gateway` preset on 365; (5) archive `microsoft-mcp-server` only
after `packages/graph` is at test parity.
