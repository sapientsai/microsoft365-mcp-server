# Change Request: To Do tools for the Civala Assistant

**Date:** 2026-10-30
**Status:** Draft, ready to build. Checked against the code at `4d059f5` and the Microsoft Graph docs. Revised 2026-10-30 after a review: the docs settle four "Verify first" items, and three gaps are closed (the duplicate check reads every page, the safe repeat is an intended `Right`, and a completed repeating occurrence is not a new task). Revised again after a second review: delta records can be partial, delta parameters are fixed by the first call, removed items get their own line, the narrow update's annotations follow MCP's definitions (DR-001 now has options), and PR 1 amends AGENTS.md for the safe repeat. DR-001 and DR-002, and the two product questions, were answered on 2026-10-30 with the recommendations (see Decision requests and item 7 below). PR 1 shipped as #164 and PR 2 as #165. Revised again after a third review: the person is told of every task the assistant completes, and PR 3 checks the extension expand on delta, sends the page-size header on every page, and treats the token as data.
**Author:** Jordan Burke
**Scope:** `packages/microsoft365` (the delegated server), the six To Do tools. The app-only `packages/graph` server is out of scope.
**Downstream:** the Civala Assistant (`civala-ai/claude-plugins`, `plugins/civala-assistant`), whose release 1.3.0 depends on this.

## Summary

The Civala Assistant keeps each person's top three to five items in a JSON file in their OneDrive. Civala wants those items in Microsoft To Do instead, where people already work and can check things off on any device. The assistant runs unattended, every half hour, as a scheduled Claude task.

The six To Do tools can't support that today:

- they can't tag a task with where it came from;
- they can't store the assistant's own data on a task;
- they can't report what changed since the last run;
- they hide importance in list output;
- every update is marked destructive, so an unattended run may stop for an approval.

This CR adds those capabilities, inside the limits Graph sets.

**The one hard limit:** Graph has no field for a task's position in a list. None of `todoTask`'s 16 properties or 5 relationships holds a manual order ([todoTask](https://learn.microsoft.com/en-us/graph/api/resources/todotask?view=graph-rest-1.0)). The assistant will keep its rank in an open extension on each task (REQ-TODO-003) and show the top items as high importance. Nothing in this CR asks Graph to sort.

## How the assistant will use the tools

So the builder knows what "done" means:

1. **Setup** (attended): creates one list, "Civala Assistant", with `create_todo_list`.
2. **Each run** (unattended):
   - reads what changed in that list since the last run (REQ-TODO-005);
   - reads the built-in Flagged email list as a signal of what the person thinks matters (REQ-TODO-006);
   - creates a task for each new top item, tagged with its source email or chat (REQ-TODO-001), and with its rank and tier in an extension (REQ-TODO-003);
   - sets importance on tasks entering or leaving the top (REQ-TODO-007).
3. **A task the person completes** is the strongest done signal the assistant gets. **A task the person adds** to the list becomes their own top item, unless it is already completed: completing a repeating task makes Graph split the finished occurrence into a new task with a new id, so an unseen task that is already completed is a done signal, never a new item (REQ-TODO-005).
4. **Never unattended:** delete, retitle, change notes, or touch the person's other lists.
5. **The fence is the caller's.** This is a general server, so nothing here can confine writes to one list. The assistant keeps to its own list by its own rule; the server does not enforce it.
6. **Rules only the caller can keep.** The server can't see across runs, so the assistant must:
   - **never run two wakes at once.** The duplicate check in REQ-TODO-001 reads, then creates, so two overlapping runs can both pass it;
   - **keep its own record of task id → `externalId`.** A removed task comes back from delta as an id only (REQ-TODO-005), with no link;
   - **treat a task the person deleted or moved away as dismissed,** and not create it again. A moved task gets a new id, and a deleted one is gone, so REQ-TODO-001's check can't see either one;
   - **tell its own writes from the person's.** Every change `set_todo_task_state` makes comes back through delta as a change;
   - **dedupe by id.** Graph may replay a change in a later response, and may list one task more than once in a response.
7. **What the assistant changes unattended** (decided 2026-10-30):
   - **status:** it completes only tasks it created itself, and only on evidence (the person replied, as triage already judges). Those never repeat, so `set_todo_task_state`'s refusal to complete a repeating task is a safety net, not a limit it meets in practice. **It always tells the person** (Jordan, 2026-10-30): the run's report names each task it completed and why, as the plugin already does for replies (`Marked done: "Reply to Jeff on pricing" (you replied at 10:42). Say if not.`), and the person can uncheck it in To Do. An uncheck then comes back through delta as a change, and the assistant must not complete that task again on the same evidence;
   - **importance:** it changes importance only on tasks that carry its extension, so it never overrides a setting the person chose.

   Both are caller rules. The server can't tell who created a task, and the tool description states them.

## Where things stand

Code references are to `packages/microsoft365/src/` unless noted.

- **Permission.** `Tasks.ReadWrite` is in `DEFAULT_INTERACTIVE_SCOPES` (`auth/scopes.ts:76`, `:90`). That list applies in OAuth proxy mode only; other modes request `.default` (see `SAI_SPEC_M365-connector-parity_2026-10-08.md`, "Where things stand"). The live Civala deployment reported the defaults on 2026-10-08, so no new consent is needed.
- **Tools.**
  - Defined at `index.ts:1755-1847`, implemented in `tools/todo-tools.ts`, Graph calls at `client/graph-client.ts:492-510`.
  - `create_todo_task` takes `title`, `body`, `due_date`, `importance` and `recurrence` only (`todo-tools.ts:99-133`).
  - `update_todo_task` takes title, status, due date, importance, body and recurrence (`todo-tools.ts:175` on).
- **Annotations.** `tools/tool-registry.ts:390-409`, rendered at `:448`. Reads are read-only. `create_todo_list` and `create_todo_task` are `additive` (`destructiveHint: false`). `update_todo_task` and `delete_todo_task` get `destructiveHint: true`.
- **Output.**
  - `formatTodoTaskSummary` (`utils/formatters.ts:626-641`) prints title, status, due date, recurrence and id. It does not print importance, categories, completion time or last-modified time.
  - It prints the title as is. AGENTS.md treats sender-chosen text as untrusted, and a flagged-email task's title is the email's subject.
- **Not exposed:** `categories`, `linkedResources`, open extensions, `startDateTime`, `reminderDateTime`, `checklistItems` and delta. Graph supports each one ([todoTask](https://learn.microsoft.com/en-us/graph/api/resources/todotask?view=graph-rest-1.0), [delta](https://learn.microsoft.com/en-us/graph/api/todotask-delta?view=graph-rest-1.0)).
- **Lists.** `todoTaskList.wellknownListName` is `none`, `defaultList`, `flaggedEmails` or `unknownFutureValue`. The docs say: "Tasks from flagged emails are present in this list." `formatTodoListSummary` already prints the well-known name (`utils/formatters.ts:613-621`).
- **Contract.** `test/plugin-contract.spec.ts` pins the line formats the Civala Assistant parses for mail and chats. To Do lines are not pinned yet.
- **Tests.** `test/todo-tools.spec.ts` has 38 tests.
- **Paging.** Core's `fetchAllPages` stops silently at 50 pages (`packages/core/src/utils/pagination.ts:6`) and does no origin check on `@odata.nextLink`. The origin-checked pager to copy is `nextMessageScanPage` (`client/graph-client.ts:88-95`).

## Requirements

Priority key: **Must** blocks Civala Assistant 1.3.0. **Should** can follow.

Every requirement follows `AGENTS.md`:

- register through `TOOL_METADATA` and `filterTools`;
- honor `MS365_READ_ONLY`, presets (`todo`) and `MS365_ENABLED_TOOLS`;
- retry writes only on 429;
- `$select` what a list needs (except To Do task lists, which refuse it; see REQ-TODO-004);
- clean untrusted text;
- check a Graph URL's origin before following it;
- return `Left` when nothing changed (REQ-TODO-001's safe repeat is the one exception, and PR 1 writes it into AGENTS.md);
- update the package README's To Do section.

### REQ-TODO-001: Tag a task with its source, and make create safe to repeat

- **Priority:** Must.
- **Statement:** `create_todo_task` takes an optional `linked_resource`: `{ external_id, web_url?, application_name, display_name? }`, stored as the task's `linkedResources`. When `external_id` is given and the list already has a task with a linked resource carrying that `externalId`, the tool returns that task and creates nothing.
- **Why:** A wake can die after Graph creates a task but before the caller records it. The next wake must not create a second copy. `externalId` is the key, for example `civala:mail:<message id>`. `web_url` puts a link back to the email in To Do.
- **Acceptance:**
  - A task created with `linked_resource` shows its `externalId` in list output, and its `externalId` and parsed `web_url` in detail output. The URL stays out of the list line, because a URL can hold `)` and break the line's parsing (REQ-TODO-004).
  - A second create with the same `external_id` in the same list returns "already exists" with the existing id, as a `Right`, and sends no POST. **This is an intended exception to AGENTS.md "Results"** (return `Left` when an operation changes nothing). That rule is about failures. A repeat of a create that already succeeded is a success, and the caller must be able to retry without treating it as an error. **PR 1 amends AGENTS.md "Results" to say so** ("an idempotent repeat of a write that already succeeded is a `Right` that names the existing item"), so the pre-merge review checks the code against the rule as amended.
  - **The duplicate check reads every page of the list, completed tasks included,** whatever `fetch_all_pages` says. A check that stops at the first page can miss the task and create the duplicate this requirement exists to stop. If any page fails to read, refuse to create, as `create_todo_list` does (`todo-tools.ts:49-56`).
  - The check pages with an origin-checked pager like `nextMessageScanPage`, not core's `fetchAllPages`, which stops silently at 50 pages. If the check hits its page cap, refuse to create rather than create blind.
  - **The ceiling.** PR 1 shipped `TODO_TASK_PAGE_SIZE = 100` and `MAX_TODO_TASK_PAGES = 50` (`client/graph-client.ts:46-47`), so the check covers 5,000 tasks, completed ones included. Above that, create refuses. Completed tasks stay in the list, so the plugin keeps its list under the ceiling, for example by deleting old completed tasks in an attended session.
  - The check can't see a task the person moved to another list or deleted. Not creating those again is a caller rule (How the assistant will use the tools, item 6).
  - `web_url` is shown only if it parses as `http:` or `https:` (AGENTS.md).
  - Tests cover both paths, and that `MS365_READ_ONLY` hides the tool.
- **Checked live after PR 1 (#164), 2026-10-30,** in a throwaway list since deleted: a tagged create stored `externalId`, `applicationName`, `displayName` and `webUrl` as sent; a second create with the same `external_id` returned the existing task and made no second one; the list line carried `(Link: …)`. Graph also accepted a linked resource with no `displayName`, so the server's default is a courtesy, not a requirement.
- **Settled by the docs:** `linkedResources` goes inside the create body. The [create task](https://learn.microsoft.com/en-us/graph/api/todotasklist-post-tasks?view=graph-rest-1.0) example "creates a **todoTask** in the specified task list, and includes a linkedResource" in the same POST. No second call is needed.
- **Settled live (2026-10-30):** Graph refuses `$filter=linkedResources/any(r: r/externalId eq '…')` with "CollectionNavigationNode operation is not supported in query filters." So the check pages the whole list with `$expand=linkedResources` and matches in code, as the acceptance above requires. `$expand=linkedResources` works on list tasks, and the Flagged email list returns `linkedResources` even without it.
- **Refuse an unsafe `external_id`.** It's a round-trip key, so output can't clean it without breaking the match. Refuse one that holds CR/LF, `|`, `(` or `)`, or is over 200 characters, before any call.

### REQ-TODO-002: Categories

- **Priority:** Should.
- **Statement:** `create_todo_task` and `update_todo_task` take `categories: string[]`, and list and detail output print them.
- **Acceptance:** Categories round-trip, and output cleans each one as untrusted text.
- **Settled by the docs:** "Each category corresponds to the **displayName** property of an outlookCategory that the user has defined" ([todoTask](https://learn.microsoft.com/en-us/graph/api/resources/todotask?view=graph-rest-1.0)).
- **Settled live (2026-10-30):** Graph accepts and stores a category that isn't in the person's master category list (`ClaudeTestCategoryNotInMasterList` round-tripped). How To Do shows such a category wasn't checked. Don't create master categories.
- **Not in `set_todo_task_state`.** Categories replace the person's own, so they stay out of the unattended tool (DR-001).

### REQ-TODO-003: Store a caller's data on a task (open extension)

- **Priority:** Must.
- **Statement:**
  - `create_todo_task` and the new tool in REQ-TODO-007 take `extension: { name, fields }`. `name` is a reverse-DNS extension name, for example `com.civala.assistant`. `fields` is a flat object of strings, numbers and booleans, at most 20 keys and at most 2 KB as JSON.
  - `list_todo_tasks` takes `extension_name`. When set, it expands that extension and prints its fields on each task line.
- **Why:** The assistant keeps rank, tier, source id and when the item was asked for. The person can't see or edit an extension, unlike the notes, so a stray edit can't break it.
- **Acceptance:**
  - Fields round-trip.
  - Updating the extension changes only the fields sent. **Graph doesn't do that for To Do:** a PATCH replaces the whole extension (live, 2026-10-30; see below), whatever the update page says for "Microsoft 365 resources like messages". So the server merges: it reads the extension, applies the fields sent, and writes the whole object back. A missing extension (404) is created with a POST to the task's `/extensions`. A `null` value is refused before the call.
  - The tool description tells callers to use one fixed name per app. The [open extension](https://learn.microsoft.com/en-us/graph/api/resources/opentypeextension?view=graph-rest-1.0) page warns: "Avoid naming extensions in a variable manner".
  - Names that aren't reverse-DNS, nested values, more than 20 keys and more than 2 KB are refused.
  - Output prints field values as cleaned single-line text.
  - Tests cover each.
- **Settled live (2026-10-30),** in a throwaway list since deleted:
  - **Create inline works.** A task POST with `"extensions": [{"@odata.type": "microsoft.graph.openTypeExtension", "extensionName": "…", …fields}]` stores the extension and returns it.
  - **The bare name filters.** `$expand=extensions($filter=id eq 'com.claude.test')` returns that extension, and a name with no match returns the task with no `extensions` key at all.
  - **A PATCH replaces.** PATCHing `{rank: 2}` onto `{rank, tier, pinned}` left only `rank`, in the response and on a fresh GET.
  - **Read shape.** An extension comes back as `extensionName`, `id` (`microsoft.graph.openTypeExtension.<name>`), the fields, and per-field annotations such as `"rank@odata.type": "#Int64"`. Output keeps only the fields: it drops `extensionName`, `id`, `@odata.type` and every key holding `@`. Match on `extensionName`, never on an `id` prefix.
  - **A missing extension** returns "Extension Id not found" (404), and a POST to the task's `/extensions` adds one.
  - **Paging keeps the expand.** With `$expand=linkedResources,extensions($filter=…)`, Graph's `@odata.nextLink` carries the same `$expand`, percent-encoded, which decodes to the string sent. So the PR 1 pager's `$expand` check holds.
- **Still open:**
  - Size limits. Graph documents none for To Do tasks. Its general limits (two open extensions per app per item, 2 KB each, per the extensibility overview) are footnoted to users, groups, devices and organizations. The 20-key and 2 KB caps above are chosen to sit inside them either way.

### REQ-TODO-004: List output an unattended caller can parse

- **Priority:** Must.
- **Statement:** Each task line in `list_todo_tasks` carries the fields a caller needs, in a fixed order, and the format is pinned in `test/plugin-contract.spec.ts`.
- **Line (as shipped in PR 1, #164):**
  ```
  - **<title>** [<status>] [<High|Low> importance] (Due: <date>) (Repeats: <pattern>) (Completed: <dateTime> <timeZone>) (Modified: <ISO time>) (Categories: a, b) (Link: <externalId>) (Ext: <JSON>) (ID: <id>)
  ```
  `TODO_LINE` in `test/plugin-contract.spec.ts` is the authority; copy it from there.
  - Optional parts are left out when empty, never printed blank.
  - Importance appears only when it isn't `normal`, labelled `[High importance]` or `[Low importance]` as mail lines are, so it can't be read as a status.
  - `(Repeats: …)` stays after Due: whether a task repeats is what a caller must know before completing it.
  - `(Due: …)` prints `dateTime` as today. `(Completed: …)` prints `completedDateTime`'s `dateTime` and `timeZone`, because it's a `dateTimeTimeZone`.
  - `(Ext: …)` is the extension's fields as one-line JSON. It sits last before the id, and the pattern anchors on ` (ID: <id>)$`, so a `)` inside a value can't end the line early.
  - **No defaults.** A field Graph didn't return is left out. Today `formatTodoTaskSummary` prints `Untitled` and `notStarted` for a missing title or status (`utils/formatters.ts:627,640`); a parsed line must never carry a value Graph didn't send.
- **Acceptance:**
  - Titles, categories and extension values are cleaned as `AGENTS.md` describes: CR/LF collapsed, length capped, `|` stripped.
  - A test feeds a title with a newline and checks the line count.
  - **No `$select`.** Graph refuses it on To Do task lists: `$select=id,title` returns "Invalid request" on both a default and the Flagged email list, while `$select` on `/me/messages` works (live, 2026-10-30). The list sends `$expand=linkedResources` and, when `extension_name` is set, the extension expand. The call site says why the AGENTS.md `$select` rule isn't followed there.
  - The contract test gains a `TODO_LINE` pattern covering the full line, Ext included, even though Ext is produced only from PR 2. The plugin then copies the pattern once.

### REQ-TODO-005: What changed since the last run (delta)

- **Priority:** Must.
- **Statement:** `list_todo_tasks` takes an optional `delta` mode (shape pending DR-002).
  - With `delta: true` and no token, it returns every task and a token.
  - With `delta_token`, it returns only the tasks added, changed or removed since that token, then the next token.
  - **Parameters are fixed by the first call.** Graph encodes every query parameter into the token: "If you use any query parameter (other than `$deltatoken` and `$skiptoken`), you must specify it in the initial **delta** request" ([todoTask delta](https://learn.microsoft.com/en-us/graph/api/todotask-delta?view=graph-rest-1.0)). So `$select` and `$expand` go only on the `delta: true` call. With `delta_token`, the delta request carries the token and nothing else. A token belongs to one list: the caller passes the same `list_id` it started with.
  - **`extension_name` with a token** is accepted, documented as "must match the first call", and used only for the GETs below. It never goes on the delta request.
  - **The token, not the URL.** The server parses each `@odata.deltaLink` with `new URL`, checks its origin, and prints only its `$deltatoken` value. On the next call it builds the URL itself from `list_id` and the token. A caller never hands the server a URL (DR-002).
  - **Each call reads to the end.** The server follows every `@odata.nextLink` until it reaches a `@odata.deltaLink`, origin-checking each, with the same page cap as REQ-TODO-001 (a named constant, given in the error text). Every page request sends `Prefer: odata.maxpagesize=…`, which the delta page lists as optional, so a large list such as Flagged email fits under the cap. It goes on every request, not only the first, because it is a header and the token carries query parameters only.
  - **The token is data.** Before building the URL, the server checks the token's characters (or `encodeURIComponent`s it), so a token holding `&` or `#` can't add query parameters. If the server hits the cap or a page fails, it returns a `Left` and prints no token, so the next run starts from the last good token.
  - **The pager.** `nextMessageScanPage` (`client/graph-client.ts:88-112` at `4d059f5`) checks one continuation link; it isn't a loop. Copy its checks (origin, no username, password or hash, a path regex for `/me/todo/lists/{id}/tasks` and `/tasks/delta`) into a capped loop. Don't wrap core's `fetchAllPages`.
  - **Changed records can be partial.** "Updated instances are represented by their **id** with _at least_ the updated properties" ([delta query overview](https://learn.microsoft.com/en-us/graph/delta-query-overview)), and a missing optional field can't be told from a cleared one. So:
    - **With a token,** every non-removed record is printed from one GET of that task, with REQ-TODO-004's `$select`, `$expand=linkedResources`, and the extension expand when `extension_name` is set. Changes per run are few, so this is cheap.
    - **On the first call** (`delta: true`), records are printed from the delta response, and only a record missing title, status or importance is filled by that GET.
    - If a GET returns 404, the task is listed as removed. If it fails otherwise, the call returns a `Left` and prints no token. Every printed line is then a full REQ-TODO-004 line.
  - **Removed tasks get their own line,** after the task lines: `Removed: <id> (<reason>)`, or `Removed: <id>` when `@removed` has no reason (the docs say it "might include" one). It can't match the plugin's item filter, `line.startsWith("- **")` (`test/plugin-contract.spec.ts:21-25`), and the contract test pins both forms as `TODO_REMOVED_LINE`. `<reason>` is whatever Graph sends; the docs explain `changed` and `deleted` only for directory objects, so the tool doesn't interpret them.
  - Delta lines carry the same fields as REQ-TODO-004, including Link and Ext: delta supports `$select`, `$top` and `$expand`. The docs say delta supports `$select`, but the task list refuses it (REQ-TODO-004), so PR 3 checks it live before relying on it. **Also check live** whether delta accepts `$expand=extensions($filter=…)`. If it refuses, the first call (`delta: true`) returns the token and lines without Ext, and the tool description says to read the list once with `list_todo_tasks` and `extension_name` for the full picture. Calls with a token are unaffected, because each record is already filled by a GET that carries the extension.
  - Within one call, a task listed more than once is printed once, from the GET above. Graph doesn't promise an order ("Don't assume a specific sequence"), so the last record isn't necessarily the newest.
  - Each tool description says that completing a repeating task makes the finished occurrence a new, already-completed task with a new id, so a caller must not read an unseen completed task as newly added work. This was seen live on 2026-10-28 (the comment in `todo-tools.ts`), not in Graph's docs, so the description doesn't cite the docs for it.
- **Why:** Each run needs only what the person did since the last run: tasks completed, added, or removed. Re-reading a whole list every half hour is slow and makes the caller diff it.
- **Acceptance:**
  - The token is printed on its own labelled line (`Delta token: …`) that a caller can pick out.
  - The server follows `@odata.nextLink` and `@odata.deltaLink` only after parsing each with `new URL` and checking its origin is the Graph host (AGENTS.md, Paths and URLs).
  - Both restart signals return a `Left` saying to start again with `delta: true`:
    - an expired token, an error with the code `syncStateNotFound`, matched on the code whatever the status (Outlook commonly sends it as `410`);
    - any `410 Gone`, with or without a `Location` header.
  - Tests use recorded responses for: first page, next page, a removed task, a partial changed record filled by a GET, a filling GET that 404s, an expired token, a `410 Gone`, a page cap hit, a task listed twice, a removed record with no reason, and a completed repeating occurrence appearing as a new task.
- **Settled by the docs** ([delta query overview](https://learn.microsoft.com/en-us/graph/delta-query-overview)):
  - Removed items carry `"@removed": {"reason": …}`, and "Possible **@removed** reasons can be `changed` or `deleted`."
  - "In case the token expires, the service should respond with a 40X-series error with error codes such as `syncStateNotFound`."
  - `410 Gone` with a `Location` header holding an empty `$deltatoken` means "the application must restart with a full synchronization".

### REQ-TODO-006: Read the Flagged email list

- **Priority:** Should.
- **Statement:** `list_todo_tasks` works on the `flaggedEmails` list, and each task's line includes its linked email (from `linkedResources`), so a caller can match it to a mail item.
- **Acceptance:**
  - A test with a recorded `flaggedEmails` response shows the link.
  - The tool description says the list is created by Outlook flags.
- **Verify first:** whether Graph refuses writes to this list. If it does, the write tools return a clear `Left` for it instead of Graph's raw error.

### REQ-TODO-007: A narrow update that is safe to run unattended

- **Priority:** Must, subject to DR-001.
- **Statement:** A new tool, `set_todo_task_state`, sets only `status`, `importance` and the extension from REQ-TODO-003 on one task. It can't change the title, notes, dates or recurrence, and it can't delete.
- **Why:** These are the only changes the assistant makes unattended. Each is reversible and touches only the person's own task.
- **Annotation:** `readOnlyHint: false`, `destructiveHint: true`, `idempotentHint: true` (DR-001, option B).
- **Order of calls, and what a failure returns:**
  1. If `status` is `completed`, read the task, and refuse if it repeats or the read fails. Nothing has changed yet.
  2. PATCH the task's `status` and `importance`, when either is given. If this fails, return a `Left`: nothing changed.
  3. Set the extension, when one is given: read it, merge the fields, PATCH the whole object, or POST it if it's missing. If this fails after step 2 succeeded, return a `Left` that names what was saved and what wasn't. The tool is idempotent, so the caller can repeat the whole call.
- **Acceptance:**
  - Any other field is refused, `categories` included (REQ-TODO-002).
  - Setting `status: completed` on a repeating task is refused, pointing to `update_todo_task`. Each such call completes another occurrence (`todo-tools.ts:290-308`), so it can't be idempotent. `recurrence` isn't an argument, so the tool reads the task first, as `update_todo_task` does (`todo-tools.ts:229-235`), and a failed read refuses the update.
  - Extension fields merge, as REQ-TODO-003 says, by read-merge-write.
  - At least one of `status`, `importance` or `extension` is required.
  - `MS365_READ_ONLY` hides it.
  - It retries on 429 only. The shared `request` path already does this for PATCH, so this is a test, not a new loop.
  - The registry gains an `idempotent` field, rendered as `idempotentHint`. `ToolAnnotations` (`tools/tool-registry.ts:432-437`) has no such field today. No `reversible` flag (DR-001 chose B). Only `set_todo_task_state` gets the new field in this CR; no existing tool changes.
- **Leave `update_todo_task` as it is:** still destructive, for attended use.

### REQ-TODO-008: Due dates land on the right day

- **Priority:** Should.
- **Statement:** A date-only `due_date` shows as that date in To Do for a person in any time zone.
- **Why:** `create_todo_task` sends `dueDateTime: { dateTime: <due_date>, timeZone: "UTC" }` (`todo-tools.ts:112`). For someone in America/New_York, midnight UTC is the evening before.
- **Acceptance:**
  - Create a task due `2026-11-02` for an Eastern-time user and confirm To Do shows November 2.
  - If it shows November 1, take a caller-supplied `time_zone`, or send noon UTC, which lands on the same day from UTC−12 to UTC+11. Don't read the mailbox's time zone: that needs `MailboxSettings.Read`, which isn't in `DEFAULT_INTERACTIVE_SCOPES` (`auth/scopes.ts:76-99`) and would break "no new consent is needed". Add a test.
- **This is a check first.** It may already be correct.

## Decision requests

### DR-001: How should `set_todo_task_state` be annotated?

- **The definitions.** MCP's `ToolAnnotations` (2025-06-18): `destructiveHint` "If true, the tool may perform destructive updates to its environment. If false, the tool performs only additive updates." `idempotentHint` "If true, calling the tool repeatedly with the same arguments will have no additional effect on its environment."
- **What REQ-TODO-007 already does to fit them:** it refuses `status: completed` on a repeating task, which would complete one more occurrence per call, and it leaves out `categories`, which would replace the person's own. What's left (status, importance and extension fields) gives the same result however many times it runs.
- **Option A:** `destructiveHint: false`, `idempotentHint: true`, with a new `reversible` registry flag. Setting status or importance overwrites a value rather than adding one, so this reads "only additive" as "nothing the person can't set back".
- **Option B:** `destructiveHint: true`, `idempotentHint: true`. This follows the letter of the definition. The tool is still narrower than `update_todo_task`.
- **Recommendation:** B, unless the plugin's gate G8 shows Claude's scheduled tasks stop for approval on `destructiveHint: true`. That test is untested so far. If they do stop, take A, because that prompt is the reason this tool exists.
- **Decision (2026-10-30):** B. Revisit A only if G8 shows the prompt.

### DR-002: Delta as a mode of `list_todo_tasks`, or its own tool? And a token or a URL?

- **Mode or tool. Recommendation:** a mode.
- **Token or URL.** The mail plan (`SAI_PLAN_tool-output-for-automated-callers_2026-10-05.md`, section 4.2) already chose a validated full `deltaLink` URL as the cursor: "the server must accept only a Graph delta URL under the caller's own `/me/mailFolders/.../messages/delta`". REQ-TODO-005 uses the bare `$deltatoken`. **Recommendation:** the token, for both. The server builds the URL, so a caller can't steer it anywhere, and there's no URL to validate. If you agree, the mail plan's section 4.2 changes to match.
- **Decision (2026-10-30):** a mode, and the bare token, for To Do and mail. The mail plan's section 4.2 now says so.

### DR-003: Fixed or caller-named extension?

- **Recommendation:** caller-named, validated reverse-DNS. This is a general Microsoft 365 server, so it shouldn't hard-code Civala.

## Delivery

Four PRs, each reviewed before merge (`AGENTS.md`, Pull requests). None bumps the package version, because the release (version, `server.json`, tag) is a maintainer step. See CLAUDE.md, Publishing.

1. **Output and tagging:** REQ-TODO-004, 001 and 002, with the contract test (`TODO_LINE`, full line) and the AGENTS.md "Results" amendment. Ship this first, because the plugin parses it.
2. **Caller data and the narrow update:** REQ-TODO-003 and 007, after DR-001 is answered. Answer DR-002 before this PR too, because PR 3 depends on it and mail delta will follow the same shape.
3. **Change feed:** REQ-TODO-005, with `TODO_REMOVED_LINE`.
4. **Flagged email and due dates:** REQ-TODO-006 and 008. Both are Should, so they can follow 1.3.0.

Run `pnpm validate` before each PR.

## Out of scope

- **Ranking or sorting.** Graph has no position field.
- **My Day.** Graph exposes no such property.
- **Sharing lists.** Graph can't; `create_todo_list` already says so.
- **Confining writes to one list.** A general server can't know which list is the caller's. The Civala Assistant enforces that itself.
- **Planner,** and the app-only `packages/graph` server.
- **JSON output** (`format: "json"`). That's the automated-callers plan. This CR keeps markdown, pinned by the contract test.
- **The plugin changes,** which ship as Civala Assistant 1.3.0.

## Open questions

- Each "Verify first" item above, answered during the build, with the fallback named in each. Still open, with nothing in the docs:
  - To Do's extension size limits (REQ-TODO-003);
  - which `@removed` reasons To Do sends, and the status `syncStateNotFound` arrives with (REQ-TODO-005);
  - whether Graph refuses writes to the Flagged email list (REQ-TODO-006);
  - how To Do shows a UTC-midnight due date (REQ-TODO-008).
- Whether to also expose `checklistItems`, `startDateTime` and `reminderDateTime`. Nobody needs them yet.
