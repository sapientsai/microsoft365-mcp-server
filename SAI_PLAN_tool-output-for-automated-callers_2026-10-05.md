# Plan — tool output for automated callers

**Date:** 2026-10-05
**Status:** Proposed. Nothing here is scheduled; v1.2.12 shipped without it.
**Scope:** `packages/microsoft365` (the delegated server). The app-only `packages/graph` server is
out of scope except where noted.

## Summary

The listing tools serve two kinds of reader with one output. A person or an LLM reads the markdown
for meaning. An automated caller, such as an unattended triage agent, needs exact fields: IDs,
times, senders and flags. Today the automated caller gets the same markdown and has to recover the
fields from it.

The plan keeps markdown as the default, and makes it one rendering of a typed record per item. The
same record is available as versioned JSON through a `format: "json"` parameter. Mail gains a
delta-based change feed, so an unattended caller stops relying on timestamps. The output contract
moves out of conversations between maintainers and into checked-in sample outputs that both sides
test against.

Graph's quirks stay in this server's tested code. Callers never have to learn them.

## 1. The problem

### How it showed up

In the week before v1.2.12, a downstream triage agent needed fields that the markdown did not carry
cleanly. The line formats changed three times to suit it:

- #95 (list one mail folder, with sender addresses) moved the Message-ID before the Graph ID.
- #106 (connector report fixes) reshaped `list_chats` lines.
- #107 (chat triage) reshaped `list_chat_messages` lines.

Each change was agreed in messages between maintainers' sessions. The only durable record is
`test/plugin-contract.spec.ts`. It is a hand-copied pin of the downstream agent's three line
patterns, taken from one commit of that project.

### Why it matters

- **Two readers, one format.** Markdown suits an LLM. A caller that needs exact fields must parse
  prose. Every wording change is a potential break, and nothing in this repo can see which wordings
  a caller depends on.
- **Silent loss.** A parser that skips lines it cannot match loses items with no error. The
  downstream agent has since added a guard: it counts unparsed item lines and reports format drift.
  That guard lives in one caller, not in this server.
- **The contract has no home.** Server and caller deploy independently. Neither repo holds the
  agreed format, except as a copy.
- **Timestamps are a weak watermark.** `since` (#107) and date filters miss or repeat items around
  edits, moves, deletions and late indexing. The downstream agent already works around late
  indexing.

### What is not the problem

The text itself. Markdown is the right default for people and for LLM callers, which adapt to
wording changes. The goal is to give automated callers a second, exact channel. It is not to remove
the readable one.

## 2. Where things stand (v1.2.12)

- **Text formatters** in `src/utils/formatters.ts` build every listing line directly from Graph
  objects. There is no intermediate record.
- **Graph knowledge is already server-side.** It is tested here and verified against real Graph by
  `pnpm smoke:live`:
  - `list_messages` puts an always-true `receivedDateTime` condition in front of any filter that
    does not lead with it. Graph rejects such filters otherwise (`InefficientFilter`). That is #97
    (let `list_messages` take any filter).
  - `list_chat_messages` pairs its `lastModifiedDateTime` filter with the matching `$orderby`.
    Graph ignores the filter otherwise (#107).
  - Chat HTML is reduced to text, `[You]` comes from a per-call `/me` lookup, and system and
    deleted messages are dropped (#107).
- **`pnpm smoke:live`** drives the built server against real Graph before each release. It needs
  an interactive sign-in, so it is a manual step, not CI.
- **`test/plugin-contract.spec.ts`** pins one caller's line patterns.

## 3. Goals and non-goals

**Goals**

1. An automated caller can read every listing as exact fields, without parsing prose.
2. The markdown and the fields cannot drift apart.
3. An unattended caller can ask "what changed in my inbox since last time" without timestamp
   edge cases.
4. The output contract is checked in, versioned and tested on both sides.

**Non-goals**

- Removing or degrading the markdown output.
- A raw Graph passthrough for automation (section 4.5).
- Chat delta sync on the delegated server, which Graph does not offer (section 4.2).

## 4. Design

### 4.1 One record per item, two renderings

Each listing tool builds a typed record per item, then renders it.

- **Markdown (default).** Today's lines are produced from the record. Existing callers see no
  change, and a snapshot test proves the output is byte-identical.
- **JSON (`format: "json"`).** The same records as a versioned envelope in the text content:

  ```json
  {
    "schema": "ms365.chatMessages/1",
    "items": [
      {
        "id": "1616964509832",
        "created": "2026-10-04T09:00:00.832Z",
        "sender": { "name": "Gregg Smith", "kind": "user" },
        "fromMe": false,
        "importance": "urgent",
        "mentionsMe": true,
        "text": "Can you check this?"
      }
    ],
    "more": false,
    "selfResolved": true
  }
  ```

  The notes that end today's markdown become fields: "more messages match" becomes `more`, and
  "the signed-in user could not be resolved" becomes `selfResolved: false`.

Because both renderings come from one record, a field cannot appear in one and not the other.

The JSON goes in the ordinary text content, not in MCP structured output (section 4.4), so it
reaches the caller through every MCP client unchanged.

**Tools, in order:** `list_messages` and `search_messages`; `list_chats`; `list_chat_messages`.
Others follow only when a caller needs them.

**Versioning.** `schema` carries a name and a major version. Adding a field is not a version
change. Removing or redefining one is, and the old major stays available until callers move.

### 4.2 Change feeds instead of timestamps

**Mail: a `mail_changes` tool on Graph's delta query.**
[message: delta](https://learn.microsoft.com/en-us/graph/api/message-delta?view=graph-rest-1.0)
works with delegated permissions (`Mail.ReadBasic` or above) on one mail folder at a time:
`/me/mailFolders/{id}/messages/delta`. Each round ends with an `@odata.deltaLink`. The next round
starting from that link returns only what changed.

- **Input:** `folder` (default `inbox`), and an optional `cursor` from the previous call.
- **Output:** the changed messages as records (4.1), the removed IDs, and a new `cursor`. Graph
  returns removals as `@removed` entries with `"reason": "deleted"`, which covers deletes and moves
  out of the folder.
- **The server stays stateless.** The caller stores the cursor. The cursor is Graph's bare
  `$deltatoken`, not the `deltaLink` URL (decided 2026-10-30 with To Do delta, in
  `SAI_SPEC_To-Do-for-Civala-Assistant_2026-10-30.md`, DR-002). The server origin-checks each
  `deltaLink`, prints only its token, and builds the next URL itself from the folder and the
  token, so a caller never hands it a URL. Query options go only on the first call, because Graph
  encodes them into the token.
- **Documented quirks to handle in code:**
  - Delta can return events that do not match the initial filter, including read/unread changes
    and `@removed` entries.
  - `$filter` supports only `receivedDateTime ge|gt`, and `$orderby` only `receivedDateTime desc`.
  - Page size is set with `Prefer: odata.maxpagesize`.

**Chats: keep `since`.**
[chats-getAllMessages: delta](https://learn.microsoft.com/en-us/graph/api/chatmessage-delta?view=graph-rest-1.0)
lists "Delegated (work or school account): Not supported". Its only permissions are application
(`Chat.Read.All`), and it returns only the last eight months. The delegated server therefore keeps
#107's `since`, which filters on `lastModifiedDateTime` per chat. Callers must still treat chat
`since` as "created or changed after". The app-only `packages/graph` server could offer a chat
change feed. That is a separate decision with its own permission and tenant-consent questions.

### 4.3 The contract as checked-in sample outputs

- **Fixtures generated in this repo.** A test renders a fixed set of inputs through every listing
  tool, in both formats, into `packages/microsoft365/contract/`. CI fails if the checked-in files
  differ from what the code produces. A format change is then a visible diff in review.
- **Callers test against those files.** A downstream project copies or vendors the fixtures and
  runs its own parser over them. A change on either side shows up as a test failure, not as a
  message between maintainers.
- **Retire the hand-copied pin.** `plugin-contract.spec.ts` goes once its one caller consumes the
  fixtures.
- **Release notes name schema changes.** Any change to a fixture or a `schema` version is listed in
  the release commit.

### 4.4 MCP structured output: later

MCP's `outputSchema` and `structuredContent` are the protocol's own channel for typed results, and
fastmcp 4.22.1 supports both. Clients disagree today about what the model sees:

- Claude Code forwards `structuredContent` and drops the text blocks.
- claude.ai and Claude Desktop forward the text and drop `structuredContent`.
- Cowork is unverified.

These reports come from public GitHub issues collected during review. They have not been reproduced
here. Adopting structured output now would change what one client shows for every listing tool.
When clients converge, the records from 4.1 become `structuredContent` with no new data model.

### 4.5 What stays as it is

- **Graph quirks live in tested server code.** Callers do not need to know the filter-order rule,
  the chat filter-and-orderby pairing, HTML reduction or the per-user `/me` lookup.
- **No raw `graph_get` for automation.** Considered and rejected for these reasons:
  - It moves every quirk above into each caller.
  - Raw chat JSON carries full HTML bodies, which costs tokens.
  - In read-only deployments it widens what a caller can read to any GET within the delegated
    scopes.
  - Client approval settings apply per tool, so a new read-only tool is not automatically allowed
    in unattended runs.

  `graph_query` remains the attended escape hatch.

- **Markdown stays the default.**

## 5. Phases

Each phase stands alone and is useful without the next.

| Phase | Work                                                                                                    | Done when                                                                                                                                  |
| ----- | ------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------ |
| 1     | Records and `format: "json"` for `list_messages`, `search_messages`, `list_chats`, `list_chat_messages` | Markdown output is byte-identical to v1.2.12 (snapshot tests); JSON matches a published schema; `smoke:live` checks the JSON for each tool |
| 2     | Contract fixtures in `packages/microsoft365/contract/`, generated and checked in CI                     | A caller's tests run against them; `plugin-contract.spec.ts` is removed                                                                    |
| 3     | `mail_changes` on Graph delta, with a validated cursor                                                  | `smoke:live` runs an initial sync and an incremental round; unit tests cover `@removed`, read-state events and cursor validation           |
| 4     | Revisit MCP structured output                                                                           | Claude Code, claude.ai and Cowork are observed to pass `structuredContent` to the model consistently                                       |

Caller-side work, outside this repo: switch to `format: "json"`, then `mail_changes`, and run any
deterministic rules as code over the JSON. Today, rules written as prose and applied by an LLM are
the less reliable half.

## 6. Risks and open questions

- **Token cost of JSON.** A JSON listing is larger than its markdown. Records should stay lean:
  clipped text, no raw HTML. Measure before and after on a real mailbox.
- **Cursor handling.** The caller must store the cursor and recover when it expires; the error
  Graph returns for an expired token needs checking. The server must validate cursors strictly
  (4.2).
- **Delta semantics.** Read-state changes and moves come back as changes. Callers must treat the
  feed as "something about this item changed", not "new mail".
- **Chat change feeds stay timestamp-based** on this server (4.2).
- **Unverified:**
  - Cowork's handling of `structuredContent`.
  - How reliably caller-side scripts run during unattended sessions.
  - The exact expired-cursor error for mail delta.

## 7. Options considered

| Option                                                              | Verdict                     | Reason                                                                   |
| ------------------------------------------------------------------- | --------------------------- | ------------------------------------------------------------------------ |
| Keep text only, with a contract test and caller-side drift counting | Interim (current)           | Works, but the contract stays implicit and each caller re-derives fields |
| MCP structured output now                                           | Later (4.4)                 | Clients disagree about what reaches the model                            |
| `format: "json"` in the text content                                | **Chosen (4.1)**            | Reaches every client; one record behind both renderings                  |
| Raw Graph via a read-only `graph_get`                               | Rejected (4.5)              | Moves quirks to callers; tokens; wider read surface                      |
| Graph delta for mail                                                | **Chosen (4.2)**            | Delegated-capable; covers edits, moves and deletes                       |
| Graph delta for chats                                               | Not possible on this server | Application permissions only                                             |

## 8. References

- #95, list one mail folder with sender addresses (mail line format)
- #97, let `list_messages` take any filter (filter-order rule; `smoke:live` introduced)
- #106, connector report fixes (chat names and order, search limit, event text)
- #107, chat triage (`since`, chat message format, contract pin)
- [message: delta](https://learn.microsoft.com/en-us/graph/api/message-delta?view=graph-rest-1.0)
- [chats-getAllMessages: delta](https://learn.microsoft.com/en-us/graph/api/chatmessage-delta?view=graph-rest-1.0)
- [List messages in a chat](https://learn.microsoft.com/en-us/graph/api/chat-list-messages?view=graph-rest-1.0)
- [MCP specification, 2025-06-18: tools and structured content](https://modelcontextprotocol.io/specification/2025-06-18/server/tools)
