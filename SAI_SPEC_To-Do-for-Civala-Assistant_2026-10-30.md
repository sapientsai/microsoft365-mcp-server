# Change Request: To Do tools for the Civala Assistant

**Date:** 2026-10-30
**Status:** Draft, ready to build. Checked against the code at `4d059f5` and the Microsoft Graph docs. Revised 2026-10-30 after a review: the docs settle four "Verify first" items, and three gaps are closed (the duplicate check reads every page, the safe repeat is an intended `Right`, and a completed repeating occurrence is not a new task).
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

## Where things stand

Code references are to `packages/microsoft365/src/` unless noted.

- **Permission.** `Tasks.ReadWrite` is in `DEFAULT_INTERACTIVE_SCOPES` (`auth/scopes.ts:76`, `:90`). That list applies in OAuth proxy mode only; other modes request `.default` (see `SAI_SPEC_M365-connector-parity_2026-10-08.md`, "Where things stand"). The live Civala deployment reported the defaults on 2026-10-08, so no new consent is needed.
- **Tools.**
  - Defined at `index.ts:1755-1845`, implemented in `tools/todo-tools.ts`, Graph calls at `client/graph-client.ts:492-510`.
  - `create_todo_task` takes `title`, `body`, `due_date`, `importance` and `recurrence` only (`todo-tools.ts:99-132`).
  - `update_todo_task` takes title, status, due date, importance, body and recurrence (`todo-tools.ts:175` on).
- **Annotations.** `tools/tool-registry.ts:390-409`, rendered at `:448`. Reads are read-only. `create_todo_list` and `create_todo_task` are `additive` (`destructiveHint: false`). `update_todo_task` and `delete_todo_task` get `destructiveHint: true`.
- **Output.**
  - `formatTodoTaskSummary` (`utils/formatters.ts:626-641`) prints title, status, due date, recurrence and id. It does not print importance, categories, completion time or last-modified time.
  - It prints the title as is. AGENTS.md treats sender-chosen text as untrusted, and a flagged-email task's title is the email's subject.
- **Not exposed:** `categories`, `linkedResources`, open extensions, `startDateTime`, `reminderDateTime`, `checklistItems` and delta. Graph supports each one ([todoTask](https://learn.microsoft.com/en-us/graph/api/resources/todotask?view=graph-rest-1.0), [delta](https://learn.microsoft.com/en-us/graph/api/todotask-delta?view=graph-rest-1.0)).
- **Lists.** `todoTaskList.wellknownListName` is `none`, `defaultList`, `flaggedEmails` or `unknownFutureValue`. The docs say: "Tasks from flagged emails are present in this list." `formatTodoListSummary` already prints the well-known name (`utils/formatters.ts:613-621`).
- **Contract.** `test/plugin-contract.spec.ts` pins the line formats the Civala Assistant parses for mail and chats. To Do lines are not pinned yet.
- **Tests.** `test/todo-tools.spec.ts` has 38 tests.

## Requirements

Priority key: **Must** blocks Civala Assistant 1.3.0. **Should** can follow.

Every requirement follows `AGENTS.md`:

- register through `TOOL_METADATA` and `filterTools`;
- honor `MS365_READ_ONLY`, presets (`todo`) and `MS365_ENABLED_TOOLS`;
- retry writes only on 429;
- `$select` what a list needs;
- clean untrusted text;
- check a Graph URL's origin before following it;
- return `Left` when nothing changed;
- update the package README's To Do section.

### REQ-TODO-001: Tag a task with its source, and make create safe to repeat

- **Priority:** Must.
- **Statement:** `create_todo_task` takes an optional `linked_resource`: `{ external_id, web_url?, application_name, display_name? }`, stored as the task's `linkedResources`. When `external_id` is given and the list already has a task with a linked resource carrying that `externalId`, the tool returns that task and creates nothing.
- **Why:** A wake can die after Graph creates a task but before the caller records it. The next wake must not create a second copy. `externalId` is the key, for example `civala:mail:<message id>`. `web_url` puts a link back to the email in To Do.
- **Acceptance:**
  - A task created with `linked_resource` shows its `externalId` and a parsed `https:` link in list and detail output.
  - A second create with the same `external_id` in the same list returns "already exists" with the existing id, as a `Right`, and sends no POST. **This is an intended exception to AGENTS.md "Results"** (return `Left` when an operation changes nothing). That rule is about failures. A repeat of a create that already succeeded is a success, and the caller must be able to retry without treating it as an error.
  - **The duplicate check reads every page of the list, completed tasks included,** whatever `fetch_all_pages` says. A check that stops at the first page can miss the task and create the duplicate this requirement exists to stop. If any page fails to read, refuse to create, as `create_todo_list` does (`todo-tools.ts:35-41`).
  - `web_url` is shown only if it parses as `http:` or `https:` (AGENTS.md).
  - Tests cover both paths, and that `MS365_READ_ONLY` hides the tool.
- **Settled by the docs:** `linkedResources` goes inside the create body. The [create task](https://learn.microsoft.com/en-us/graph/api/todotasklist-post-tasks?view=graph-rest-1.0) example "creates a **todoTask** in the specified task list, and includes a linkedResource" in the same POST. No second call is needed.
- **Verify first:** whether Graph supports `$filter=linkedResources/any(r: r/externalId eq '…')`. The docs don't mention it. If it doesn't work, page the whole list with `$expand=linkedResources` and match in code, as the acceptance above requires.

### REQ-TODO-002: Categories

- **Priority:** Should.
- **Statement:** `create_todo_task` and `update_todo_task` take `categories: string[]`, and list and detail output print them.
- **Acceptance:** Categories round-trip, and output cleans each one as untrusted text.
- **Settled by the docs:** "Each category corresponds to the **displayName** property of an outlookCategory that the user has defined" ([todoTask](https://learn.microsoft.com/en-us/graph/api/resources/todotask?view=graph-rest-1.0)).
- **Verify first:** whether Graph refuses a category that isn't in the person's master category list. The docs don't say. If it refuses, say so in the tool description. Don't create master categories.

### REQ-TODO-003: Store a caller's data on a task (open extension)

- **Priority:** Must.
- **Statement:**
  - `create_todo_task` and the new tool in REQ-TODO-007 take `extension: { name, fields }`. `name` is a reverse-DNS extension name, for example `com.civala.assistant`. `fields` is a flat object of strings, numbers and booleans, at most 20 keys and at most 2 KB as JSON.
  - `list_todo_tasks` takes `extension_name`. When set, it expands that extension and prints its fields on each task line.
- **Why:** The assistant keeps rank, tier, source id and when the item was asked for. The person can't see or edit an extension, unlike the notes, so a stray edit can't break it.
- **Acceptance:**
  - Fields round-trip.
  - Updating the extension changes only the fields sent; Graph merges an update, keeping fields left out. A `null` value is refused before the call, because Graph refuses it.
  - Names that aren't reverse-DNS, nested values, more than 20 keys and more than 2 KB are refused.
  - Output prints field values as cleaned single-line text.
  - Tests cover each.
- **Settled by the docs:**
  - An extension can be created with the task in one call. The [open extension](https://learn.microsoft.com/en-us/graph/api/opentypeextension-post-opentypeextension?view=graph-rest-1.0) page lists `POST /users/{id|userPrincipalName}/todo/lists/{id}/tasks` under "Create an extension in a new resource instance".
  - Naming: `extensionName` is the bare name (`com.civala.assistant`). `id` is "a fully qualified name which concatenates the extension type and unique text identifier" ([get open extension](https://learn.microsoft.com/en-us/graph/api/opentypeextension-get?view=graph-rest-1.0)).
- **Verify first:**
  - Which name `$expand=extensions($filter=id eq '…')` matches. The docs filter by the full `id` in one example and the bare name in another.
  - The `id` prefix on read. A create response in the docs shows `microsoft.graph.openTypeExtension.Com.Contoso.Referral`; a read example shows `Microsoft.OutlookServices.OpenTypeExtension.Com.Contoso.Referral`. To Do is stored in Outlook, so either may come back. Match an extension on read by `extensionName`, never by an `id` prefix.
  - Size limits. Graph documents none for To Do tasks. Its general limits (two open extensions per app per item, 2 KB each, per the extensibility overview) are footnoted to users, groups, devices and organizations. The 20-key and 2 KB caps above are chosen to sit inside them either way.

### REQ-TODO-004: List output an unattended caller can parse

- **Priority:** Must.
- **Statement:** Each task line in `list_todo_tasks` carries the fields a caller needs, in a fixed order, and the format is pinned in `test/plugin-contract.spec.ts`.
- **Proposed line:**
  ```
  - **<title>** [<status>] [<importance>] (Due: <date>) (Completed: <ISO time>) (Modified: <ISO time>) (Categories: a, b) (Link: <externalId>) (Ext: k=v; k=v) (ID: <id>)
  ```
  Optional parts are left out when empty, never printed blank.
- **Acceptance:**
  - Titles, categories and extension values are cleaned as `AGENTS.md` describes: CR/LF collapsed, length capped, `|` stripped.
  - A test feeds a title with a newline and checks the line count.
  - `$select` covers exactly the printed fields.
  - The contract test gains a `TODO_LINE` pattern. The plugin will copy that pattern.

### REQ-TODO-005: What changed since the last run (delta)

- **Priority:** Must.
- **Statement:** `list_todo_tasks` takes an optional `delta` mode (shape pending DR-002).
  - With `delta: true` and no token, it returns every task and a token.
  - With `delta_token`, it returns only the tasks added, changed or removed since that token, then the next token.
  - Removed tasks are listed separately, by id, with Graph's reason (`changed` or `deleted`).
  - Delta lines carry the same fields as REQ-TODO-004, including Link and Ext: delta supports `$select`, `$top` and `$expand`.
  - Each tool description says that completing a repeating task makes Graph split off the finished occurrence as a new, already-completed task with a new id, so a caller must not read an unseen completed task as newly added work.
- **Why:** Each run needs only what the person did since the last run: tasks completed, added, or removed. Re-reading a whole list every half hour is slow and makes the caller diff it.
- **Acceptance:**
  - The token is printed on its own labelled line (`Delta token: …`) that a caller can pick out.
  - The server follows `@odata.nextLink` and `@odata.deltaLink` only after parsing each with `new URL` and checking its origin is the Graph host (AGENTS.md, Paths and URLs).
  - Both restart signals return a `Left` saying to start again with `delta: true`:
    - an expired token, a 40X error with a code such as `syncStateNotFound`;
    - a sync reset, `410 Gone` with a `Location` header, which Graph sends after tenant maintenance or migration.
  - Tests use recorded responses for: first page, next page, a removed task, an expired token, a `410 Gone`, and a completed repeating occurrence appearing as a new task.
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
- **Statement:** A new tool, `set_todo_task_state`, sets only `status`, `importance`, `categories` and the extension from REQ-TODO-003 on one task. It can't change the title, notes, dates or recurrence, and it can't delete.
- **Why:** These are the only changes the assistant makes unattended. Each is reversible and touches only the person's own task.
- **Annotation:** `readOnlyHint: false`, `destructiveHint: false`, `idempotentHint: true`.
- **Acceptance:**
  - Any other field is refused.
  - Setting `status: completed` on a repeating task reports the series rolling forward, as `update_todo_task` already does (`todo-tools.ts:293-308`).
  - `MS365_READ_ONLY` hides it.
  - It retries on 429 only.
  - The registry gains a flag for "reversible change to the user's own item" (DR-001), so `tool-registry.ts:448` doesn't overload `additive`.
- **Leave `update_todo_task` as it is:** still destructive, for attended use.

### REQ-TODO-008: Due dates land on the right day

- **Priority:** Should.
- **Statement:** A date-only `due_date` shows as that date in To Do for a person in any time zone.
- **Why:** `create_todo_task` sends `dueDateTime: { dateTime: <due_date>, timeZone: "UTC" }` (`todo-tools.ts:112`). For someone in America/New_York, midnight UTC is the evening before.
- **Acceptance:**
  - Create a task due `2026-11-02` for an Eastern-time user and confirm To Do shows November 2.
  - If it shows November 1, send the date in the mailbox's time zone, or at a time To Do treats as date-only, and add a test.
- **This is a check first.** It may already be correct.

## Decision requests

### DR-001: May a reversible update be marked non-destructive?

- **Question:** Can `set_todo_task_state` carry `destructiveHint: false`? And what should the registry flag be called?
- **Recommendation:** Yes, with a new flag, `reversible: true`, separate from `additive`. MCP defines `destructiveHint` as "may perform destructive updates". Setting a status or importance the person can set back isn't one.
- **Open:** whether Claude's scheduled tasks prompt on `destructiveHint: true` at all. That's the plugin's gate G8, untested. If they don't prompt, this tool is still worth having, because it's narrower.

### DR-002: Delta as a mode of `list_todo_tasks`, or its own tool?

- **Recommendation:** a mode. Mail delta (`SAI_PLAN_tool-output-for-automated-callers_2026-10-05.md`, section 4.2) should follow whatever this picks, so choose it once for both.

### DR-003: Fixed or caller-named extension?

- **Recommendation:** caller-named, validated reverse-DNS. This is a general Microsoft 365 server, so it shouldn't hard-code Civala.

## Delivery

Three PRs, each reviewed before merge (`AGENTS.md`, Pull requests). None bumps the package version, because the release (version, `server.json`, tag) is a maintainer step. See CLAUDE.md, Publishing.

1. **Output and tagging:** REQ-TODO-004, 001 and 002, with the contract test. Ship this first, because the plugin parses it.
2. **Caller data and the narrow update:** REQ-TODO-003 and 007, after DR-001 is answered. Answer DR-002 before this PR too, because PR 3 depends on it and mail delta will follow the same shape.
3. **Change feed and Flagged email:** REQ-TODO-005, 006 and 008.

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
  - filtering a list by `linkedResources` (REQ-TODO-001);
  - whether Graph refuses a category missing from the master list (REQ-TODO-002);
  - which name the extension `$filter` matches, and To Do's extension limits (REQ-TODO-003);
  - whether Graph refuses writes to the Flagged email list (REQ-TODO-006);
  - how To Do shows a UTC-midnight due date (REQ-TODO-008).
- Whether to also expose `checklistItems`, `startDateTime` and `reminderDateTime`. Nobody needs them yet.
