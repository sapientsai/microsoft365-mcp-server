# Agent instructions

## Environment variable compatibility

When adding or changing tools, HTTP routes, scripts, authentication, or Graph requests:

- Read the affected package's README configuration sections before implementing the change. For the Microsoft 365 server, use [packages/microsoft365/README.md](packages/microsoft365/README.md#environment-variables) as the configuration contract.
- Apply every relevant documented control to each entry point. Check read-only mode, presets, tool filters, draft requirements, mailbox allowlists, transport restrictions, authentication, API version, extraction limits, and Graph resilience settings where applicable.
- Register tools through `TOOL_METADATA` and `filterTools`. Gate auxiliary HTTP routes with the same allowed-tool set; standalone scripts must enforce applicable restrictions before side effects.
- Reuse the existing authentication configuration, Graph version resolver, and shared resilience layer instead of hardcoding alternatives. Preserve operation safety when deciding which failures can be retried.
- Add regression coverage demonstrating that applicable restrictions block the operation and configured values reach the request. Include alternate entry points affected by the change.
- Update the package README when intentionally changing the configuration contract.

## Graph write safety

- Retry a write (POST, PATCH, PUT, DELETE) only on 429. On a 503, 504, timeout or network error the write may already have run, so a retry can repeat it or report a finished change as failed. This holds for `$batch` sub-requests too: judge each by its own method and status. GET is retried as the core resilience layer allows.
- Read `Retry-After` with core's `parseRetryAfterMs`, which handles seconds and HTTP-dates and caps the wait at 60 s. Don't parse it by hand or sleep for an uncapped value.
- Build retries on the shared `request` / `resilientFetch` path rather than writing a new loop. If a tool needs its own loop (for example over batch sub-responses), apply the same rules as `packages/core/src/resilience.ts`.
- Give `$batch` sub-requests positional ids (`"0"`, `"1"`, ...). Graph compares batch ids case-insensitively, and message ids are base64, so two ids differing only by case fail the whole batch.
- `$select` the fields a list needs. Some endpoints return large payloads by default; attachment lists, for example, include each file's base64 `contentBytes`.

## Untrusted text in tool output

Mail subjects, sender names and addresses, file and attachment names, and link URLs are chosen by whoever sent or shared the item. The model reads tool output as the server's own words, so:

- Collapse CR/LF to a space and cap the length before putting such text in output (and strip `|` in pipe-separated rows). See `cell` in `src/mail/sender-summary.ts` and `oneLine` in `src/tools/mail-tools.ts`.
- Never let it sit where it could pass for one of the tool's own lines, such as a preview's closing "Re-run with dry_run: false" instruction.
- Show a URL only if it parses as `http:` or `https:`, and print the parsed `url.href`, not the raw string.
- Add a test that feeds a newline-bearing value through and checks the output's line structure.

## Paths and URLs

- Check a Graph URL's target by parsing it with `new URL(url, base)` and comparing its `origin` before reading the path. Don't split the raw string. Leading whitespace, repeated or back slashes, and encoded `%2f` all change what Graph resolves. `mailboxOfUrl` in `src/mail/mailbox-path.ts` is the reference.

## Results

- When an operation changes nothing (every item failed, the batch call failed, the breaker was open), return a `Left` / `UserError` so the MCP result carries `isError`. A partial success stays `Right`, with the failures listed. Otherwise the caller reports work as done that never happened.
- The one exception is an idempotent repeat: a write that finds its own earlier success already in place (for example `create_todo_task` with an `external_id` a task already carries) returns a `Right` that names the existing item. The caller wanted that item to exist, and it does, so a retry must not read as a failure.

## Pull requests

- Every PR gets a pre-merge review, and the reviewed commit is what merges. After a review, add commits on top. Don't rebase or force-push, or the review no longer matches the head. If a maintainer has pushed to your branch, `git pull` before adding more.
