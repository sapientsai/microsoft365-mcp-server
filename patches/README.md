# fastmcp@4.0.0.patch

Local vendor patch against `fastmcp@4.0.0` to fix an `ERR_HTTP_HEADERS_SENT` crash
in the server's OAuth proxy. Applied automatically by pnpm via the
`patchedDependencies` entry in `pnpm-workspace.yaml`.

## The crash

```
Error [ERR_HTTP_HEADERS_SENT]: Cannot write headers after they are sent to the client
  at ServerResponse.writeHead (node:_http_server:354:11)
  at IncomingMessage.<anonymous> (fastmcp/dist/chunk-UVX47AE5.js:1831:19)
```

Triggered by every `POST /oauth/register` (DCR) call from the claude.ai connector.
Container crashes on the first auth attempt.

## Root cause — a contract mismatch between two transitive deps

fastmcp's OAuth proxy handlers (`/oauth/register`, `/oauth/consent`, `/oauth/token`)
use this pattern:

```js
req.on("data", (chunk) => (body += chunk))
req.on("end", async () => {
  /* write response */
})
return // handler returns before response is written
```

That works only under mcp-proxy's pre-6.4.5 contract, where `onUnhandledRequest` was
terminal: if set, its return ended the dispatch chain, no further writes to `res`.

[punkpeye/mcp-proxy#59](https://github.com/punkpeye/mcp-proxy/pull/59) (merged 2026-04-08,
released in mcp-proxy 6.4.5) changed that contract. The fix's intent was legitimate
— let `/health` and other custom routes skip API-key auth — but the implementation
turned `onUnhandledRequest` into just one link in the chain, with an unconditional
`res.writeHead(404).end()` tail fallback at the end of the request listener.

Under the new contract, fastmcp's async-listener handlers return with
`res.writableEnded === false`, mcp-proxy writes its tail 404, and fastmcp's async
handler wakes up later trying to write into an already-closed response → crash.

The bug surfaced via pnpm lockfile drift: fastmcp pins `mcp-proxy: ^6.4.0`, so any
lockfile regenerated after 2026-04-08 silently picks up 6.4.5+.

Full timeline: see commit `4945e4a` and the conversation in
`~/.claude/plans/concurrent-juggling-panda.md`.

## What the patch does

Three surgical edits to `dist/chunk-UVX47AE5.js`:

1. **Drop `outgoing: res` from the Hono pre-pass** — Hono's node adapter could write
   a 404 to `res` asynchronously, racing with OAuth handlers.
2. **Early-return the OAuth proxy block if `res.headersSent`** — defense in depth
   against anything that wrote upstream.
3. **Inline-`await` the body** on `/oauth/register`, `/oauth/consent`,
   `/oauth/token`. Replaces the `req.on("end", async () => {...}); return;` pattern
   with `const body = await new Promise(...); /* write response */; return;`. This
   is the real fix — it honors mcp-proxy's new contract regardless of whether
   downstream is pre-6.4.5 or post-6.4.5.

## When to retire this patch

Remove it when any of these lands:

- **fastmcp upstream fix**: the OAuth handlers await body inline (or equivalent).
  Check `node_modules/fastmcp/dist/chunk-*.js` after a version bump — if the
  `req.on("end", async () => ...)` pattern is gone from `/oauth/register`, patch
  is no longer needed.
- **mcp-proxy reverts terminal semantics**: if `startHTTPServer` skips the tail 404
  when `onUnhandledRequest` was called (or delays it across a microtask), the race
  disappears even without the fastmcp fix.
- **Publishing via `@jordanburke/fastmcp` fork**: cut a new fork release with the
  patch baked in, re-pin `package.json` to the fork, and delete `patches/`.

## Removing

```
pnpm patch-remove fastmcp
rm -rf patches/
# package.json + pnpm-workspace.yaml should auto-clean on the next install
```

## Verifying the patch is applied

After `pnpm install`:

```
grep -n "await new Promise" node_modules/fastmcp/dist/chunk-UVX47AE5.js
# should show the inline-await replacements around lines ~1821 and ~1888
```

Or just start the server and smoke-test:

```
curl -X POST http://localhost:3000/oauth/register \
  -H 'Content-Type: application/json' \
  -d '{"redirect_uris":["https://claude.ai/api/mcp/auth_callback"],"client_name":"t"}'
# expect 201 + DCR JSON, no crash in stderr
```

## Filing upstream

Not blocking, no rush. If/when filed, the clean writeup goes at `punkpeye/fastmcp`:

> fastmcp's OAuth proxy handlers for `/oauth/register`, `/oauth/consent`,
> `/oauth/token` attach `req.on("end", async () => ...)` listeners and return
> synchronously. This races any HTTP dispatcher that checks response state
> after handler return — currently surfaces as `ERR_HTTP_HEADERS_SENT` under
> mcp-proxy 6.4.5+. Fix: await the body inline before writing the response.
