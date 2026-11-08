# Phase 6 spike: moving microsoft365 onto somamcp 2.1 — 2026-10-24

Phase 6 of `SAI_PLAN_ms-graph-monorepo_2026-06-20.md` moves `packages/microsoft365` off FastMCP and
onto somamcp. The plan makes a spike a hard gate. It must confirm somamcp can carry the OAuth proxy, the
`/upload` routes and the MCPB bundle before 365 changes.

> **Superseded 2026-11-06.** SomaMCP 2.3 shipped its own Azure OAuth server (option B below), and PR 178 moved
> microsoft365 onto it and removed fastmcp. Option A was never built. This memo is kept as the record of the spike.

## Bottom line

Everything but the OAuth proxy maps onto somamcp 2.1 today. The OAuth proxy needs a decision: somamcp
has none, and the MCP SDK 2.x server package ships no OAuth server either.

The move is worth doing now. **microsoft365 1.2.17 still answers every `2026-07-28` request with a 400**
once a client is past sign-in. That is the fault behind claude.ai connectors that sometimes load no
tools (sapientsai/SomaMCP#36).

## Why now: the 400 still ships

- The hosted deployment checks sign-in first. An unauthenticated `2026-07-28` `server/discover` gets a 401 with
  the right `WWW-Authenticate`, so a probe from outside cannot see the fault.
- A local 1.2.17 in `client-token` mode (same fastmcp 4.22.10 → mcp-proxy stack, no sign-in) answers an
  enveloped `2026-07-28` `server/discover` and `tools/list` with **HTTP 400**,
  `server_missing_modern_protocol_support`.
- `packages/graph` on somamcp 2.1 (PR #151) answers the same requests with 200.

## What maps cleanly

| Need | somamcp 2.1 | Notes |
|---|---|---|
| stdio + httpStream | `somamcp/node` `start()` | Same as graph in #151 |
| Tool shape | `addTool({ name, description, parameters, execute, annotations })` | Same fields `registerTools` passes today |
| `context.session.accessToken` | `authenticate` resolves the session tools see | OAuth mode reads it in `wrapExecute` |
| `UserError` (18 files) | `somamcp` exports `UserError` | Import swap. Check that the SDK backend renders it as an `isError` result the way FastMCP did |
| `/upload`, `/upload/attachment` | `addRoute({ method, path, handler })` | Keep `protected: false` and the existing `relayAuth`: these take upload tickets, not OAuth JWTs |
| `/ping` | `addRoute` GET `/ping` → `ok` | Dockerfile `HEALTHCHECK`, README and the hosted deployment's probes all use it. somamcp's own `/health` is extra |
| Server instructions | `createServer({ instructions })` | Unchanged |
| `info` tool | Added by default | Pass `enableIntrospection: false`, or accept one more tool and update the README counts |
| MCPB bundle | Not affected | `scripts/build-mcpb.sh` zips `manifest.json` only |
| functype | One copy, 1.19.0 | Done in #151 across the workspace |

## What changes behavior

- **`MCP_STATELESS=false` stops working.** somamcp 2 serves httpStream stateless only. #151 makes graph log a
  warning; 365 should do the same, and its README should say so.
- **`FASTMCP_HOST`** is set by the Dockerfile and documented. Keep reading it beside `HOST`.
- **`server-env.spec.ts`** mocks FastMCP's class and `getApp`. Rewrite it against `server.fetch`, which
  exercises the real HTTP path with no port.

## The blocker: the OAuth proxy

In `oauth-proxy` mode, FastMCP does three jobs that somamcp does not:

1. **`AzureProvider` / `OAuthProxy`** (from `fastmcp/auth`) run the OAuth server: client registration,
   PKCE, the Azure redirect, consent, code and refresh exchange, signed JWTs, and encrypted token storage
   in `DiskStore`. Roughly 2,800 lines of bundled JS, MIT, importing only Node built-ins (`crypto`,
   `fs/promises`, `path`, `util`) and zod.
2. **Route glue inside the `FastMCP` class**, not in `fastmcp/auth`: two `/.well-known` metadata
   endpoints and five `/oauth/*` routes (`register`, `authorize`, `callback`, `consent`, `token`). Each
   reads a body and calls one `OAuthProxy` method. About 300 lines of Node stream code. On Hono with web
   `Request`s it should be much shorter, since the body reading goes away (an estimate, not measured).
3. **The 401 challenge** on the MCP endpoint, with `WWW-Authenticate: Bearer resource_metadata=…`.
   somamcp's `authenticate` can reject with a `Response`, which it serves as-is, so this carries over.

One shim is needed whichever way we go: `AzureProvider.authenticate` reads `request.headers.authorization`,
which is the Node `IncomingMessage` shape. A web `Request` has `headers.get()`, so passed as-is every
call would fail. Call `proxy.loadUpstreamTokens(bearer)` directly instead.

### Options

| | What | Drops fastmcp? | Existing sign-ins | Effort |
|---|---|---|---|---|
| **A** | Keep `fastmcp/auth` as a library. Write the route glue on somamcp's `addRoute`, and shim `authenticate` | No. fastmcp, mcp-proxy and SDK 1.x stay installed, unused at runtime except `fastmcp/auth` | Survive: same library, same storage format | Days |
| **B** | Port the OAuth proxy into somamcp (`somamcp/node`, since it needs `fs` and `crypto`), then move 365 | Yes | Survive only if the port keeps the JWT, encryption and `DiskStore` formats. Otherwise every user signs in once | A SomaMCP project first, then the 365 move |
| **C** | A third-party Hono OAuth proxy | Yes | Every user signs in again | Unknown. SomaMCP#36 names mcp-use's `oauth/*` as a reference only |

### Recommendation

**A now, then B.** A ships the protocol fix to hosted users quickly and keeps every sign-in. Write A's
route glue as the seed of somamcp's version, so B moves code rather than rewriting it. B is the plan's
own rule ("fix it in somamcp before touching 365") and the only option that removes fastmcp.

## Done so far

- #151 (graph on somamcp 2.1): the plan's first step, in review.
- This spike. No microsoft365 code has changed.
