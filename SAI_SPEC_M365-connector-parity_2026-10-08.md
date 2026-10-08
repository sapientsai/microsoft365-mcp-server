# Change Request: parity with Anthropic's Microsoft 365 connector

**Date:** 2026-10-08
**Status:** Draft. Revised 2026-10-08 after a review against the code (`97bbf35`) and the Microsoft Graph docs.
**Author:** Jordan Burke
**Scope:** `packages/microsoft365` (the delegated server). The app-only `packages/graph` server is out of scope.
**Routes to:** the maintainers of this repo. Engineering responds with a disposition for each requirement and an ADR for each decision request.

## Summary

Civala plans to turn off Anthropic's built-in Microsoft 365 connector for Claude and standardize on this server. Running both confuses Claude about which one to call. This server already covers more of Microsoft 365 (Planner, To Do, OneNote, contacts, groups, meeting creation, binary upload, raw Graph). Anthropic's connector still does about six things this server can't. This CR lists those gaps as requirements so the cutover doesn't take anything away from users.

## Why now

The cutover email is drafted for a one-week notice. Until the Must items below ship, a user who asks Claude to set their out-of-office or read a shared mailbox gets "Access denied," where Anthropic's connector would have done it.

## Where things stand

These facts were checked against the repo and the live Civala deployment on 2026-10-08. Code references are to `packages/microsoft365/src/` unless noted.

- **Default scopes.** As-built: `DEFAULT_INTERACTIVE_SCOPES` (`auth/scopes.ts:76`) has no `MailboxSettings.*`, no `Mail.*.Shared` and no `Files.*.All`.
  - This list only applies in OAuth proxy mode (`auth/oauth-provider.ts:103`). The interactive, certificate and client-secret modes request `https://graph.microsoft.com/.default`, which means "whatever the app registration has been granted" (`auth/auth-manager.ts:131,161`).
  - `MS365_EXTRA_SCOPES` adds scopes on top, without a release, and is also OAuth-proxy only.
- **Live Civala deployment.** `get_auth_status` reports the defaults plus `OnlineMeetings.ReadWrite` and `OnlineMeetingTranscript.Read.All`.
- **Live test calls (read-only).** All three ran through `graph_query`:
  - `GET /me/mailboxSettings/automaticRepliesSetting` returned "Access denied."
  - `GET /me/mailFolders/inbox/messageRules` returned "Access denied."
  - `GET /users/it@civala.com/messages` returned "Access denied." This one is suggestive, not conclusive, because it also depends on how the mailbox is delegated.
- **Document extraction.** As-built: the extract path (`packages/extract/src/extract.ts:93-96`) reads PDF, .docx, .xlsx and text only.
  - `.doc` and `.xls` are mapped to content types (`extract.ts:23-24`) but have no extractor, so they return "Unsupported."
  - Nothing in `packages/extract` handles PowerPoint. `pptx` appears only in `packages/core/src/utils/upload-helpers.ts`.
  - One lossy path already exists. `read_document` takes a `format` parameter (`index.ts:805`) that it passes to Graph as `?format=` (`tools/read-document-tools.ts:128`). With `format=pdf`, Graph converts a OneDrive or SharePoint file to PDF, which the PDF extractor then reads.
- **Reachable today only through `graph_query`.** These need no new scopes but have no named tool:
  - editing or deleting a draft (`Mail.ReadWrite`)
  - accepting or declining an invite (`Calendars.ReadWrite`)
  - replying in a channel thread (`ChannelMessage.Send`; `send_channel_message` takes only team, channel and content, `index.ts:1089-1093`)
  - starting a chat, and listing chat members (`Chat.ReadWrite` covers both; `send_chat_message` needs an existing `chat_id`, `index.ts:1039`)
  - tagging a message with a category (`Mail.ReadWrite`; no tool reads or sets `categories`)
- **Existing write controls.** Each environment flag applies to the whole deployment (`tools/tool-registry.ts:205-224`):
  - `MS365_READ_ONLY` hides every write tool, including `graph_query`.
  - `MS365_REQUIRE_DRAFT` hides `send_message`, `send_reply`, `send_reply_all` and `send_forward`. `send_draft`, `send_chat_message`, `send_channel_message` and `graph_query` stay available, so a `POST /me/sendMail` through `graph_query` still sends.
  - `MS365_ENABLED_TOOLS` (a name pattern), `MS365_PRESETS` (tool domains) and `MS365_ORG_MODE` (organization-wide tools) choose which tools are exposed.
- **Audit log.** `utils/audit.ts` writes each tool name and its parameters to stderr. It doesn't record which user made the call. It redacts only `access_token`, `password`, `secret` and `content_type`, so mail bodies are logged in full.
- **From session notes, not verifiable in the repo:**
  - The live Civala server runs `MS365_REQUIRE_DRAFT=true`.
  - Civala's app registration requires users to be assigned to it.
  - Its delegated permissions were granted by an admin. Adding a scope means updating the app registration and patching the existing grant directly, because re-running admin consent doesn't update it.
  - The recorded grant list doesn't include `Sites.ReadWrite.All` as a delegated permission. That record may be stale.
  - Whether the live server runs in OAuth proxy mode isn't recorded.
- **Reference.** Anthropic's capabilities come from the [Set up the Microsoft 365 connector](https://support.claude.com/en/articles/12542951-set-up-the-microsoft-365-connector) help article as of 2026-10-08.

## Requirements

Priority key: **Must** blocks the cutover. **Should** is parity, soon after. **Later** goes beyond parity.

### REQ-MS365-001 (provisional): Out-of-office replies

- **Status:** Proposed. **Priority:** Must.
- **Statement:** A user can ask Claude to read, turn on, change or turn off their automatic replies.
- **Who needs it:** Any user before leave or travel.
- **Rationale:** This is a common, low-risk request, and Anthropic's connector handles it. Today ours returns "Access denied," which reads as a regression right after the cutover.
- **Acceptance signal:** In a test mailbox, Claude sets an out-of-office with a start and end time, reads it back, and clears it.
- **Change Proposal against:** `packages/microsoft365` scopes (`MailboxSettings.ReadWrite`) and mail tools.

### REQ-MS365-002 (provisional): Shared mailboxes, read

- **Status:** Proposed. **Priority:** Must.
- **Statement:** A user can search and read any shared mailbox they already have delegate access to in Microsoft 365, and can't read any mailbox they lack access to.
- **Who needs it:** IT staff working the it@civala.com mailbox, and anyone with a team mailbox.
- **Rationale:** Anthropic's connector reads shared mailboxes. Ours can't, so team inbox work would stop at the cutover.
- **Acceptance signal:** A delegate lists and reads messages in it@civala.com. A non-delegate gets a clear "no access" message.
- **Change Proposal against:** `packages/microsoft365` scopes (`Mail.Read.Shared`) and mail tools.

### REQ-MS365-003 (provisional): Named tools for common writes

- **Status:** Proposed. **Priority:** Must.
- **Statement:** These actions are available as named tools, not only through raw Graph:
  - editing a draft
  - deleting a draft
  - accepting, declining or tentatively accepting an invite
  - replying in a channel thread
  - starting a new chat
  - listing chat members
- **Who needs it:** Every user. Claude finds named tools more reliably than `graph_query`.
- **Rationale:** The permissions already exist, so this is tool work only. Without it, Claude often says it can't do things it can.
- **Acceptance signal:** Each action succeeds in a fresh session from a plain-English request, with no hint to use `graph_query`.
- **Change Proposal against:** mail, calendar, Teams and chat tools.
- **Ordering note:** this adds write tools before REQ-008 (guardrails) ships. `MS365_REQUIRE_DRAFT` already makes outgoing mail go through a draft on the Civala server, but chat and channel posts send directly, as they do today.

### REQ-MS365-004 (provisional): Inbox rules

- **Status:** Proposed. **Priority:** Should.
- **Statement:** A user can ask Claude to list, create, change or delete their inbox rules.
- **Who needs it:** Users cleaning up or automating their mail.
- **Rationale:** This is parity with Anthropic's connector. A wrong rule can hide mail, so the user must be able to see it and undo it.
- **Acceptance signal:**
  - Claude creates a rule, shows it, and deletes it on request.
  - A rule that forwards mail outside Civala is refused or flagged.
- **Change Proposal against:** `packages/microsoft365` mail tools. Uses the same `MailboxSettings.ReadWrite` scope as REQ-001.

### REQ-MS365-005 (provisional): Mail categories

- **Status:** Proposed. **Priority:** Should.
- **Statement:** A user can create, rename and delete their mail categories, and tag or untag messages with them.
- **Rationale:** Parity. Triage workflows that sort by category stall without it. Today no tool reads or sets a message's categories; only `graph_query` can.
- **Acceptance signal:** Claude creates a category, tags a message with it, and removes the category.

### REQ-MS365-006 (provisional): Files shared from other people's OneDrive

- **Status:** Proposed. **Priority:** Should.
- **Statement:** A user can read and update files that colleagues share with them from their own OneDrive, within the sharing permissions they already have.
- **Rationale:** Users expect to work on files colleagues share with them. Whether this fails today is unverified:
  - Microsoft lists `Sites.Read.All` and `Sites.ReadWrite.All` as able to read drive items, and `Sites.ReadWrite.All` is in the default list. So the permission may already reach a colleague's OneDrive.
  - The default list only applies in OAuth proxy mode, and the session notes don't show `Sites.ReadWrite.All` granted to Civala's app as a delegated permission.
  - Before adding `Files.ReadWrite.All`, run a live `graph_query` read of a file shared from a colleague's OneDrive.
  - Don't build on `/me/drive/sharedWithMe`. Microsoft has marked it deprecated, in a degraded state until November 2026.
- **Acceptance signal:** A file shared from another user's OneDrive can be read and, when shared with edit rights, updated.
- **Change Proposal against:** `packages/microsoft365` scopes, if the live test fails.

### REQ-MS365-007 (provisional): PowerPoint and older Office formats

- **Status:** Proposed. **Priority:** Should.
- **Statement:** Document reading covers PowerPoint (.pptx) and the older .doc, .xls and .ppt formats, wherever reading already works today.
- **Who needs it:** Anyone asking Claude about a deck. Decks are a large share of Civala's SharePoint content.
- **Rationale:** Anthropic's connector reads these. Ours can't by default, so "summarize this deck" fails unless Claude knows to pass `format=pdf`.
- **Implementation notes:**
  - Graph's `?format=pdf` conversion supports doc, ppt, pptx and xls, and the PDF extractor already exists. Using it by default for these types is a cheap first step.
  - The conversion only works for OneDrive and SharePoint files, not mail attachments.
  - The PDF probably drops speaker notes (unverified), so .pptx likely needs its own parser to meet the acceptance signal.
- **Acceptance signal:** A sample of real Civala decks and older files returns readable text, including slide text and speaker notes.
- **Change Proposal against:** the extract path used by `read_document`.

### REQ-MS365-008 (provisional): Admin guardrails on write actions

- **Status:** Proposed. **Priority:** Should.
- **Statement:** An administrator can turn individual write actions off, or require confirmation, for the whole organization or for a group of users. Repeated sends are capped per user.
- **Who needs it:** IT, before a broad rollout of write tools.
- **Rationale:** Anthropic's connector offers per-tool Ask, Allow and Blocked settings, role-based access and send limits.
  - This server already has deployment-wide controls (see "Existing write controls" above). It has no per-group controls and no send caps.
  - `graph_query` can send mail even when `MS365_REQUIRE_DRAFT` is on. Only `MS365_READ_ONLY` removes it.
  - Without per-group controls and caps, one runaway agent session could mass-send mail as a user, and we couldn't contain it short of turning the whole server off.
- **Acceptance signal:**
  - With sending turned off for a test group, a send attempt is refused with a clear message, while drafts still work. This covers `send_draft` and `graph_query`, not just the direct-send tools.
  - A burst of sends past the cap is refused.

### REQ-MS365-009 (provisional): Mark agent-sent mail

- **Status:** Proposed. **Priority:** Should.
- **Statement:** Mail sent through this server is identifiable afterward as sent by an AI agent on the user's behalf.
- **Rationale:** GxP (good practice rules for regulated pharma work) and the audit trail require knowing who, or what, sent a message. Anthropic's connector tags its emails this way.
- **Implementation notes:**
  - The audit log isn't an audit trail yet. It doesn't record the user, and the deployment decides how long stderr is kept.
  - The audit log also writes full mail bodies, which may not be acceptable for regulated content.
  - Graph lets a custom `x-` header be set only when a message is created (`POST /me/messages`), not by a later update. The docs imply the header survives sending but don't say so outright.
  - `create_draft` doesn't expose headers today. The reply and forward draft tools would need a different approach, because Graph's JSON reply creation doesn't document header support.
- **Acceptance signal:** A sent message can be told apart from a hand-sent one, using the message itself or a log record.

### REQ-MS365-010 (provisional): Meeting recordings and AI insights

- **Status:** Proposed. **Priority:** Later.
- **Statement:** A user can find recordings of meetings they attended, and read Microsoft's AI meeting insights where their license includes them.
- **Rationale:** Parity. Transcripts already cover most needs, so this is lower priority. Both scopes need admin consent (see DR-001).
- **Acceptance signal:**
  - A recorded meeting the user attended returns its recording link.
  - Insights return for a licensed user and give a clear "not available" message otherwise.

### REQ-MS365-011 (provisional): Attachments on outbound mail

- **Status:** Proposed. **Priority:** Later.
- **Statement:** A user can attach a file (from OneDrive, SharePoint or an upload) to a draft, reply or forward.
- **Rationale:** Neither connector can do this today. It's the most common reason users still finish email by hand, so it would put us ahead of Anthropic's connector.
- **Acceptance signal:** A draft with a SharePoint file attached opens in Outlook with the attachment intact.

## Decision requests

### DR-001: How should the new permissions reach users?

- **Status:** Open.
- **Needed by:** before REQ-001 and REQ-002 ship.
- **Context:**
  - `scopes.ts` notes that adding an admin-consent scope to the defaults breaks sign-in for tenants that haven't granted it.
  - The same happens for any new default scope in an app that requires user assignment, or in a tenant that blocks user consent. Microsoft: "Applications that require users to be assigned to the application must have their permissions consented by an administrator, even if the user consent policies ... would otherwise allow a user to consent."
  - This only affects OAuth proxy mode. The `.default` modes take whatever the app registration has been granted.
  - `MS365_EXTRA_SCOPES` adds scopes per deployment with no release, also in OAuth proxy mode only.
- **Constraints:** No existing deployment loses sign-in. Civala gets the scopes before the cutover.
- **Admin consent required (delegated), per the [Graph permissions reference](https://learn.microsoft.com/en-us/graph/permissions-reference):**

  | Permission | Admin consent | Needed by |
  | --- | --- | --- |
  | `MailboxSettings.ReadWrite` | No | REQ-001, REQ-004 |
  | `Mail.Read.Shared` | No | REQ-002 |
  | `Files.ReadWrite.All` | No | REQ-006, only if the live test fails |
  | `OnlineMeetingRecording.Read.All` | Yes | REQ-010 |
  | `OnlineMeetingAiInsight.Read.All` | Yes | REQ-010 |

  REQ-003 needs no new scopes. `Chat.ReadWrite`, already in the defaults, covers creating a chat and listing its members.
- **Candidates (illustrative; engineering selects):**
  1. Add user-consentable scopes to the defaults, and keep admin-consent ones in `MS365_EXTRA_SCOPES`.
  2. Use `MS365_EXTRA_SCOPES` only, for the Civala deployment.
  3. Both, in phases.
- **Review recommendation:** option 3. Unblock Civala with `MS365_EXTRA_SCOPES` first, then move the user-consentable scopes into the defaults in a release whose notes warn about assignment-required apps and tenants that block user consent. This holds only if the Civala server runs in OAuth proxy mode. Either way, an admin still has to grant the new scopes on Civala's app registration first. `MS365_EXTRA_SCOPES` only saves the release.
- **What decides it:** the risk to other deployments, and how fast Civala can be unblocked.
- **Resolution path:** engineering ADR.

### DR-002: Where should admin guardrails live?

- **Status:** Open.
- **Needed by:** before write tools roll out beyond IT.
- **Context:** REQ-008 needs per-action controls and send caps. The server's deployment-wide flags already exist (see "Existing write controls"). Claude's own per-tool approval settings already exist. Entra (Microsoft's identity service) can revoke individual scopes.
- **Constraints:** IT can change a control without a release. Users see a clear reason when an action is refused.
- **Candidates (illustrative; engineering selects):**
  - server-side policy configuration
  - relying on Claude's per-tool settings plus Entra scope revocation
  - a mix of the two
- **What decides it:** coverage across Claude clients (claude.ai, Cowork, Claude Code), how hard it is for IT to run, and whether it can be audited.
- **Resolution path:** engineering ADR.

### DR-003: How should shared mailboxes be addressed?

- **Status:** Open.
- **Context:** REQ-002 can be met by adding a mailbox option to the existing mail tools, or by adding separate tools.
- **Constraints:** Claude must not mix up the user's own mailbox and a shared one. Results show which mailbox each item came from.
- **Candidates (illustrative; engineering selects):** a `mailbox` parameter on the existing tools, or dedicated shared-mailbox tools.
- **What decides it:** how reliably Claude picks the right mailbox in testing, and the tool-count budget.
- **Resolution path:** engineering ADR.

## Out of scope

- Changing Teams settings or permissions. Neither connector does this.
- The Online Archive (In-Place Archive) mailbox. Anthropic's connector excludes it too.

## Open questions

1. Is it@civala.com a shared mailbox with delegate access, or a group or distribution list? This decides whether REQ-002 covers it.
2. Do Civala users have Copilot licenses? Without them, the AI-insights part of REQ-010 does nothing.
3. Does the live Civala server run in OAuth proxy mode? This decides whether `MS365_EXTRA_SCOPES` and the default scope list apply to it at all.
4. Is `Sites.ReadWrite.All` granted to Civala's app as a delegated permission? This decides whether REQ-006 needs a new scope.
