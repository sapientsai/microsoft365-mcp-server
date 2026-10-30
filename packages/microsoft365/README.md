## microsoft365-mcp-server

[![Node.js CI](https://github.com/sapientsai/microsoft365-mcp-server/actions/workflows/node.js.yml/badge.svg)](https://github.com/sapientsai/microsoft365-mcp-server/actions/workflows/node.js.yml)
[![npm version](https://img.shields.io/npm/v/microsoft365-mcp-server.svg)](https://www.npmjs.com/package/microsoft365-mcp-server)

A Model Context Protocol (MCP) server for Microsoft 365 — manage email, calendar, contacts, files, Teams chats, channels, Planner, OneNote, To Do, users, and groups via Microsoft Graph API.

> **Which server?** This is the **delegated** (per-user OAuth) server — the full interactive M365
> assistant. Its sibling in this monorepo, [`microsoft-mcp-server`](../graph), is the **app-only**
> (`client_credentials`) server for headless, tenant-wide document-RAG. Both sit on the shared
> [`@sapientsai/ms-graph-core`](../core). Use this one when a user is present to sign in; use
> `microsoft-mcp-server` when there's no user to consent (a service/tenant credential).

## Features

- **79 Tools** across 12 Microsoft 365 domains + generic Graph API escape hatch
- **5 Auth Modes**: Interactive, certificate, client secret, client-provided token, OAuth proxy
- **Draft Workflow**: Create drafts for user review in Outlook, then send when approved
- **Tool Filtering**: Presets, regex patterns, read-only mode, and org-mode gating
- **Auto-Pagination**: `fetch_all_pages` parameter on all list tools (max 50 pages)
- **Multi-Account**: Register and switch between multiple authenticated accounts
- **Functional Programming**: [functype](https://github.com/jordanburke/functype) patterns — `Either`, `Option`, `Try`, `Brand` types
- **Type-Safe**: Branded IDs, Zod parameter schemas, strict TypeScript
- **Modern Build System**: [ts-builds](https://github.com/jordanburke/ts-builds) + [tsdown](https://tsdown.dev/)
- **Dual Transport**: stdio (default) and HTTP stream
- **Persistent OAuth Sessions**: DiskStore-backed token persistence survives server restarts
- **SharePoint Sites**: Browse, search, and access files across SharePoint sites the user has permissions on

## Quick Start

```bash
# Install globally
npm install -g microsoft365-mcp-server

# Or run directly
npx microsoft365-mcp-server
```

### Claude Desktop / VS Code Configuration

Add to your `claude_desktop_config.json` or MCP settings:

```json
{
  "mcpServers": {
    "microsoft365": {
      "command": "npx",
      "args": ["-y", "microsoft365-mcp-server"],
      "env": {
        "MS365_AUTH_MODE": "interactive",
        "MS365_CLIENT_ID": "your-azure-app-client-id",
        "MS365_TENANT_ID": "common"
      }
    }
  }
}
```

## Authentication

### Interactive (Browser/Device Code)

Simplest setup — opens a browser or displays a device code for headless environments.

```bash
MS365_AUTH_MODE=interactive
MS365_CLIENT_ID=your-client-id
MS365_TENANT_ID=common          # "common" for multi-tenant
```

For headless hosts, SSH sessions, or a machine where the browser cannot launch,
select device-code authentication explicitly:

```bash
MS365_USE_DEVICE_CODE=true
```

This skips the browser and prints a URL and code to stderr. Open the URL on a
browser-capable device and enter the code to sign in.

Either way, the token cache and the signed-in account are saved to
`~/.config/microsoft365-mcp-server/token-cache/` (or `$XDG_CONFIG_HOME/...`), so a
restart reuses them instead of asking you to sign in again. The files are written
with mode `0600`, and a cache file readable by anyone else is discarded rather than
used. Set `MS365_TOKEN_CACHE=false` to keep tokens in memory only, or
`MS365_TOKEN_CACHE_PATH` to put them somewhere else. In a container, point
`MS365_TOKEN_CACHE_PATH` at a mounted volume, or the cache is lost when the
container is rebuilt.

A cached access token is dropped when the server's configuration changes: any
`MS365_*` or `GRAPH_*` variable (secrets aside), the tenant, or the client ID. The
refresh token is kept, so the next call gets a new access token silently, under
the app registration's current permissions. This covers the common case of
granting a permission in Azure, such as `Mail.ReadWrite.Shared`, together with a
change here, such as a new `MS365_ALLOWED_MAILBOXES` entry. A permission granted
in Azure with no change here still reaches the server only when the cached token
expires, within about an hour, or after the cache directory is deleted.

When the server is launched unattended, by an agent gateway or a scheduled job,
nobody is there to enter a device code, and a sign-in prompt just blocks until the
host gives up. Set `MS365_DISABLE_SIGN_IN=true` and the server exits at startup with
an `[Auth] Sign-in required` error instead whenever the cached sign-in cannot be
reused. Sign in once by running the server without it, then put it back. It needs
the token cache, so it cannot be combined with `MS365_TOKEN_CACHE=false`.

### Client Secret

For service accounts and automation.

```bash
MS365_AUTH_MODE=client-secret
MS365_TENANT_ID=your-tenant-id
MS365_CLIENT_ID=your-client-id
MS365_CLIENT_SECRET=your-secret
```

### Certificate

For production service principals with certificate-based auth.

```bash
MS365_AUTH_MODE=certificate
MS365_TENANT_ID=your-tenant-id
MS365_CLIENT_ID=your-client-id
MS365_CERT_PATH=/path/to/cert.pem
MS365_CERT_PASSWORD=optional-password
```

### Client-Provided Token

For external token management — the MCP client supplies tokens.

```bash
MS365_AUTH_MODE=client-token
MS365_ACCESS_TOKEN=optional-initial-token
```

Use the `set_access_token` tool to update tokens at runtime.

### OAuth Proxy

Full OAuth 2.0 authorization server mode using FastMCP's built-in AzureProvider. Handles PKCE, consent screens, JWT issuance, and token refresh automatically. Requires HTTP transport.

```bash
MS365_AUTH_MODE=oauth-proxy
MS365_TENANT_ID=your-tenant-id
MS365_CLIENT_ID=your-client-id
MS365_CLIENT_SECRET=your-client-secret
MS365_OAUTH_BASE_URL=http://localhost:3000
PORT=3000
```

Endpoints provided automatically:

- `GET /.well-known/oauth-authorization-server` — OAuth metadata
- `GET /authorize` — Redirect to Microsoft auth
- `POST /token` — Token exchange
- `GET/POST /mcp` — MCP protocol (with bearer auth)
- `GET /ping` — Pre-auth health check (for Docker/Kubernetes liveness probes)

OAuth tokens are persisted to disk via FastMCP's `DiskStore`, so users stay authenticated across server restarts. Refresh tokens last ~90 days. Set `TOKEN_STORAGE_PATH` to customize the storage directory (default: `/tmp/ms365-tokens`). For persistence across container recreations, mount a Docker volume at that path.

### Azure AD App Registration

You need an Azure AD (Entra ID) app registration:

1. Go to [Azure Portal](https://portal.azure.com) > App registrations > New registration
2. Set supported account types based on your needs (single tenant, multi-tenant, or personal)
3. Add redirect URIs:
   - **Mobile/Desktop platform**: `http://localhost` (for interactive mode — allows any port)
   - **Web platform**: `http://localhost:3000/oauth/callback` (for OAuth proxy mode)
4. Add Microsoft Graph **delegated** permissions:

| Permission                                     | Domain                                         |
| ---------------------------------------------- | ---------------------------------------------- |
| `User.Read`                                    | User profile                                   |
| `Mail.Read`, `Mail.Send`                       | Email                                          |
| `Calendars.ReadWrite`                          | Calendar                                       |
| `Calendars.Read.Shared`                        | Calendar free/busy (find_meeting_availability) |
| `Contacts.Read`                                | Contacts                                       |
| `Files.ReadWrite`                              | OneDrive                                       |
| `Sites.Read.All`, `Sites.ReadWrite.All`        | SharePoint sites                               |
| `Chat.ReadWrite`                               | Teams chats                                    |
| `ChatMessage.Read`, `ChatMessage.Send`         | Chat messages                                  |
| `Team.ReadBasic.All`                           | Teams                                          |
| `Channel.ReadBasic.All`, `ChannelMessage.Send` | Channels                                       |
| `Tasks.ReadWrite`                              | Planner & To Do                                |
| `Notes.ReadWrite`                              | OneNote                                        |

5. Grant admin consent (for org tenants)
6. Create a client secret (for client-secret and OAuth proxy modes)

### Restricting Access to Specific Users

By default, any user in your tenant can authenticate. To restrict to specific users:

```bash
# Enable user assignment requirement
az ad sp update --id <app-id> --set appRoleAssignmentRequired=true

# Assign a user
SP_ID=$(az ad sp show --id <app-id> --query id -o tsv)
USER_ID=$(az ad user show --id user@example.com --query id -o tsv)
az rest --method POST \
  --url "https://graph.microsoft.com/v1.0/servicePrincipals/${SP_ID}/appRoleAssignments" \
  --body "{\"principalId\":\"${USER_ID}\",\"resourceId\":\"${SP_ID}\",\"appRoleId\":\"00000000-0000-0000-0000-000000000000\"}"
```

Or in Azure Portal: Enterprise Applications > your app > Users and groups > Add user.

### Other Mailboxes (Delegated and Shared)

Mail tools take an optional `mailbox` parameter — an email address — so one server can
work across a household or a team: your own mail by default, someone else's when named.

What `MS365_ALLOWED_MAILBOXES` means depends on who the server signs in as:

| Sign-in                                           | List empty (default)                   | List set                                 |
| ------------------------------------------------- | -------------------------------------- | ---------------------------------------- |
| As a user (interactive, device code, OAuth proxy) | Any mailbox your account can reach     | Only the listed mailboxes, plus your own |
| As an application (client secret, certificate)    | No mailbox; every mail call is refused | Only the listed mailboxes                |

Signed in as yourself, Microsoft already limits you to the mailboxes you have been given
access to, so the list is an optional narrowing. Signed in as an application, Graph would
reach any mailbox the tenant's ApplicationAccessPolicy allows, so the list is required.
Client token mode can carry either kind of token and goes by the token itself: a token
with delegated scopes (`scp`) counts as a user, one with app roles (`roles`) as an
application. A token it cannot read (personal-account tokens are opaque) keeps your own
mailbox but reaches no other until it is listed.

```bash
MS365_ALLOWED_MAILBOXES="bel@example.com,household@example.com"
```

Omitting `mailbox` means your own mailbox, so existing setups need no configuration and
behave exactly as before. Naming your own address (primary, UPN or an alias) also works
without listing it and is treated the same as omitting it. This applies whenever you sign
in as yourself (interactive, device code, client token, OAuth proxy), not app-only, where
there is no own mailbox. Your addresses are looked up from `/me` with the token of the
request itself and cached per token, so in OAuth proxy mode each user only ever matches
their own.

Two boundaries apply, and both matter:

| Boundary                  | Set by                                                 | What it does                                                                              |
| ------------------------- | ------------------------------------------------------ | ----------------------------------------------------------------------------------------- |
| Azure                     | Delegation, or an ApplicationAccessPolicy for app-only | Caps what the credential can reach at all. The real enforcement.                          |
| `MS365_ALLOWED_MAILBOXES` | This server                                            | Narrows one deployment within that cap, and returns a clear error instead of a Graph 403. |

The env var is configuration, not security: it lets a single app registration serve
several agents with different reach — a triage bot limited to one mailbox, an assistant
given two — without minting a registration per agent. It cannot widen what Azure allows.
Signed in as yourself, set it when you have Exchange access to mailboxes (a partner's,
say) that you do not want this particular server or agent to touch. The list scopes the
agent; it does not authorise anything.

**Which permission you need depends on how the server authenticates:**

- **OAuth proxy (delegated).** Reaching another mailbox needs a `.Shared` scope; the
  non-shared one does not grant it. The proxy requests its scopes itself, so add it
  explicitly:

  ```bash
  MS365_EXTRA_SCOPES="Mail.ReadWrite.Shared"   # or Mail.Read.Shared for read-only
  ```

- **Interactive / device code / client token (delegated).** These request `.default`, so
  `MS365_EXTRA_SCOPES` is not read. Add `Mail.ReadWrite.Shared` (or `Mail.Read.Shared`) as
  a delegated permission on the app registration, consent to it, and sign in again.

  In both delegated cases, the mailbox's owner must also have granted you delegate access (in Outlook, or via
  `Add-MailboxPermission`). Actions are performed as you: Exchange audit records
  distinguish delegate access from owner access, but not one delegate tool from another.

- **Client secret / certificate (app-only).** These request `.default`, so permissions
  come from the app registration — nothing to add here. Each mailbox must be listed in
  `MS365_ALLOWED_MAILBOXES`. There is no `/me` for an app-only token, so `mailbox` is
  **required** on every mail call; a call without it is refused
  with an error saying to pass one, before anything reaches Graph. Scope the app to specific
  mailboxes with an ApplicationAccessPolicy, otherwise `Mail.ReadWrite` reaches every
  mailbox in the tenant:

  ```powershell
  New-ApplicationAccessPolicy -AppId <app-id> `
    -PolicyScopeGroupId mailbox-group@example.com -AccessRight RestrictAccess
  ```

`read_document` and `graph_query` take raw Graph paths, so they apply the same check to
any path under `/users/{address}/...` (or `/users('{address}')/...`), reads included, and
`graph_query` to every sub-request of a `$batch`, refusing the whole batch if one fails.
Directory reads about a user are not mailbox data and are not checked: the user itself
(`/users/{address}`), `photo`, `photos`, `manager`, `directReports`, `memberOf`,
`transitiveMemberOf`, `presence`, and OneDrive (`drive`, `drives`). Anything else under a
user (messages, calendars, contacts, mailbox settings, To Do, and any resource not listed
here) is treated as mailbox data. `/me/...`, `/drives/...` and `/sites/...` are unaffected.

Note that `scan_messages` refs are per-mailbox: a ref from one mailbox passed to a call
addressing another is refused with an explanation rather than resolving to the wrong
message.

### Cleaning a large inbox

A mailbox with tens of thousands of messages is mostly a few hundred senders repeating.
Two tools make that tractable without paging headers through the model:

1. `summarize_senders` reads the whole folder server-side and returns one row per
   sender — count, unread, first and last date, latest subject — most frequent first.
2. `move_messages_matching` moves everything from chosen senders (or matching an OData
   filter) to a destination, using Graph JSON batching (20 moves per round-trip, with
   throttling retried). It is a dry run unless `dry_run: false` is passed, and a live
   run refuses to move more than `limit` (default 1000) messages, so a mis-scoped
   filter cannot empty a folder. `destination: deleteditems` is the bulk-delete: it is
   reversible from Deleted Items, and a hard delete is deliberately not offered.

Both fetch without `$orderby` — Graph rejects a `from/emailAddress` filter combined with
a `receivedDateTime` sort ("The restriction or sort order is too complex").

### Safety Layers

| Layer                    | Protection                                                   | Default                                                       |
| ------------------------ | ------------------------------------------------------------ | ------------------------------------------------------------- |
| **User assignment**      | Only assigned users can authenticate                         | Off (enable above)                                            |
| **Platform governance**  | Per-tool allow/confirm/deny in Claude Desktop Enterprise     | Platform-level                                                |
| **Tool filtering**       | Presets, read-only, org-mode gating                          | All tools                                                     |
| **Mailbox allowlist**    | `MS365_ALLOWED_MAILBOXES` caps which mailboxes are reachable | Delegated: whatever Microsoft allows the user; app-only: none |
| **Tenant restriction**   | `MS365_TENANT_ID` locks to one org                           | `common`                                                      |
| **M365 native recovery** | Recycle bins, version history                                | Built-in                                                      |

**Recovery by domain:**

- Mail, Calendar, OneDrive, SharePoint: Deleted Items / Recycle Bin (30-93 days), version history
- Teams messages: Immutable (can't be deleted via API)
- Contacts, Planner tasks, To Do tasks: No native recovery — use `MS365_READ_ONLY=true` or platform governance to restrict writes

## Tool Filtering

### Presets

Named bundles of tool domains:

| Preset          | Domains                                        |
| --------------- | ---------------------------------------------- |
| `personal`      | mail, calendar, contacts, todo, files, onenote |
| `collaboration` | chats, teams, meetings, planner, groups        |
| `productivity`  | mail, calendar, todo                           |
| `all`           | everything                                     |

```bash
MS365_PRESETS=personal                    # just personal tools
MS365_PRESETS=personal,collaboration      # personal + team tools
MS365_PRESETS=rag                         # read_document + files + graph_query
```

If not set, all tools are registered.

### Other Filters

```bash
MS365_ENABLED_TOOLS="mail|calendar"   # regex pattern — only matching tools registered
MS365_READ_ONLY=true                  # hide all write tools (send, create, update, delete)
MS365_ORG_MODE=true                   # enable org-only tools (teams, chats, meetings, groups, planner, list_users)
MS365_REQUIRE_DRAFT=true              # hide all send_* mail tools; force the create_*_draft + send_draft flow
```

The HTTP `/upload` endpoint is mounted only when `get_upload_config` survives these
filters. Read-only mode, a preset without files, or a tool pattern that excludes
`get_upload_config` also disables the upload endpoint. The `/upload/attachment`
endpoint follows `get_attachment_upload_config` the same way: read-only mode, a
preset without mail, or a pattern that excludes the tool disables it.

Both endpoints accept the same upload ticket: an opaque stand-in for the caller's own
Graph token that expires after about 10 minutes. It is not tied to one file or draft,
so it carries the caller's own access for that window, no more. The attachment
endpoint re-checks on every request that the target is a draft.

Org mode is required for Teams, Chats, Meetings, Groups, Planner, and user listing. Without it, these tools are hidden to prevent 403 errors on personal accounts.

## Available Tools

### Mail (23 tools)

| Tool                           | Description                                                               |
| ------------------------------ | ------------------------------------------------------------------------- |
| `list_messages`                | List messages, all folders or one (`folder`), with optional filtering     |
| `scan_messages`                | Scan compact headers with session refs and nextLink cursors               |
| `get_message`                  | Get a specific message with full body                                     |
| `search_messages`              | Search messages by query                                                  |
| `send_message`                 | Send a new email                                                          |
| `send_reply`                   | Reply to the sender and send now (threaded, original quoted)              |
| `send_reply_all`               | Reply to all recipients and send now (threaded, original quoted)          |
| `send_forward`                 | Forward a message to new recipients and send now (original quoted)        |
| `create_draft`                 | Create a new email draft in the Drafts folder                             |
| `create_reply_draft`           | Create a reply draft — threaded, with the original quoted underneath      |
| `create_reply_all_draft`       | Create a reply-all draft — threaded, with the original quoted underneath  |
| `create_forward_draft`         | Create a forward draft — original quoted, recipients you specify          |
| `send_draft`                   | Send an existing email draft                                              |
| `update_draft`                 | Edit a draft's subject, body or recipients (refuses non-drafts)           |
| `delete_draft`                 | Delete a draft (refuses non-drafts)                                       |
| `get_attachment_upload_config` | Relay URL to attach a file to a draft, up to 150 MB (HTTP transport only) |
| `list_attachments`             | List a message's attachments, with a read_document path for file ones     |
| `save_attachment`              | Save an attachment to a local file and return its path (stdio only)       |
| `list_mail_folders`            | List top-level mail folders with item, unread and subfolder counts        |
| `move_message`                 | Move a message to a well-known folder, a folder name, or a folder ID      |
| `batch_move_messages`          | Move up to 50 messages to one folder in a single call                     |
| `summarize_senders`            | Count a folder by sender, server-side — find what to sweep                |
| `move_messages_matching`       | Move everything matching senders/filter, batched; dry-run by default      |

`scan_messages` returns a compact list of message headers. If there’s another page,
it includes a `cursor` — pass that back on its own to continue the same scan, even if
the last page was short or empty. Cursors and message refs work until the server restarts.

Search stops at 1,000 results. To check every matching message, use a filter such as
`hasAttachments eq true` and keep paging until there’s no cursor left.

> The `create_*_draft` tools produce a properly threaded draft (same conversation, full
> quoted history) for review, then send via `send_draft`. They remain available under
> `MS365_REQUIRE_DRAFT=true`; the `send_*` tools are hidden in that mode, and `graph_query`
> refuses requests that send mail without a draft (`sendMail`, and `reply`, `replyAll` or
> `forward` on a message, including inside a `$batch`).

> **Reading attachments.** `read_document` extracts _text_, so a scanned PDF or a
> photographed letter comes back empty — there is no text layer to extract. Use
> `save_attachment` for those: it writes the file locally and returns the path, leaving
> the client to read the PDF or image with whatever it already has. That keeps
> rasterising and OCR out of this server.

### Calendar (8 tools)

| Tool                        | Description                                                     |
| --------------------------- | --------------------------------------------------------------- |
| `list_events`               | List calendar events                                            |
| `list_calendar_view`        | List event instances in a date range (expands recurring series) |
| `find_meeting_availability` | Suggest meeting times where all participants are free           |
| `get_event`                 | Get event details (body as text by default)                     |
| `create_event`              | Create a new event                                              |
| `update_event`              | Update an existing event                                        |
| `delete_event`              | Delete an event                                                 |
| `respond_to_event`          | Accept, decline or tentatively accept an invitation             |

### Contacts (7 tools)

| Tool              | Description                                                   |
| ----------------- | ------------------------------------------------------------- |
| `list_contacts`   | List contacts                                                 |
| `get_contact`     | Get contact details                                           |
| `create_contact`  | Create a new contact                                          |
| `search_contacts` | Search contacts                                               |
| `update_contact`  | Update a contact; list fields replace the whole list          |
| `delete_contact`  | Delete a contact, returning its details                       |
| `export_contacts` | Export every contact folder to a local JSON file (stdio only) |

### Files / OneDrive (7 tools)

| Tool               | Description                                                                   |
| ------------------ | ----------------------------------------------------------------------------- |
| `list_drive_items` | List files and folders (supports `folder_id` or `folder_path` for navigation) |
| `get_drive_item`   | Get file/folder metadata                                                      |
| `search_files`     | Search OneDrive/SharePoint (25 results by default)                            |
| `download_file`    | Download a file — returns content inline for text files under 100KB           |
| `create_folder`    | Create a new folder                                                           |
| `upload_file`      | Upload a file to OneDrive (text or base64-encoded binary, max ~4MB)           |

### SharePoint Sites (5 tools, org mode)

| Tool                | Description                                                             |
| ------------------- | ----------------------------------------------------------------------- |
| `list_sites`        | List followed sites, or search all sites by query                       |
| `get_site`          | Get SharePoint site details                                             |
| `list_site_drives`  | List document libraries (drives) in a site                              |
| `list_site_items`   | List files/folders in a site drive (supports `folder_id`/`folder_path`) |
| `search_site_files` | Search a SharePoint site (25 results by default)                        |

SharePoint tools use **delegated permissions** — users see only the sites and files they have access to. Private channel sites are properly isolated; access requires channel membership.

### Chats (3 tools, org mode)

| Tool                 | Description                                                                |
| -------------------- | -------------------------------------------------------------------------- |
| `list_chats`         | List Teams chats by latest message, named                                  |
| `list_chat_messages` | List chat messages as text, with flags and `since`                         |
| `send_chat_message`  | Send in a chat, or to people by email (`to`). Use `48:notes` for self-chat |

> With `to`, one person reuses (or starts) your one-on-one chat with them. Two or more people
> start a new group chat on every call, because Graph doesn't deduplicate group chats, so reuse
> the returned chat ID for follow-ups.

### Teams (4 tools, org mode)

| Tool                    | Description                                                               |
| ----------------------- | ------------------------------------------------------------------------- |
| `list_teams`            | List joined teams                                                         |
| `list_channels`         | List channels in a team                                                   |
| `list_channel_messages` | List recent channel messages, with their IDs                              |
| `send_channel_message`  | Send a message to a channel, or reply in a thread (`reply_to_message_id`) |

### Meeting Transcripts (2 tools, org mode)

Requires opt-in scopes that are **not** requested by default — see [Meeting transcripts](#meeting-transcripts).

| Tool                       | Description                                      |
| -------------------------- | ------------------------------------------------ |
| `list_meeting_transcripts` | List a Teams meeting's transcripts and their IDs |
| `get_meeting_transcript`   | Get one transcript's text                        |

### Users & Groups (6 tools, org mode except get_me)

| Tool                 | Description                      |
| -------------------- | -------------------------------- |
| `get_me`             | Get authenticated user's profile |
| `list_users`         | List organization users          |
| `get_user`           | Get a specific user's profile    |
| `list_groups`        | List organization groups         |
| `get_group`          | Get group details                |
| `list_group_members` | List group members               |

### Planner (5 tools, org mode)

| Tool                  | Description          |
| --------------------- | -------------------- |
| `list_plans`          | List Planner plans   |
| `list_planner_tasks`  | List tasks in a plan |
| `get_planner_task`    | Get task details     |
| `create_planner_task` | Create a new task    |
| `update_planner_task` | Update a task        |

### OneNote (10 tools)

| Tool                          | Description                                              |
| ----------------------------- | -------------------------------------------------------- |
| `list_onenote_notebooks`      | List notebooks                                           |
| `list_onenote_sections`       | List sections in a notebook                              |
| `list_onenote_pages`          | List pages in a section                                  |
| `get_onenote_page_content`    | Get page content as HTML                                 |
| `create_onenote_page`         | Create a page from HTML (title + body)                   |
| `update_onenote_page_content` | Append to / modify a page's content without rewriting it |
| `create_onenote_section`      | Create a section in a notebook                           |
| `create_onenote_notebook`     | Create a notebook                                        |
| `copy_onenote_page`           | Copy a page to another section (async)                   |
| `delete_onenote_page`         | Delete a page (destructive)                              |

> `create_onenote_page` and `update_onenote_page_content` take HTML. OneNote supports a
> constrained HTML subset and silently drops unsupported CSS/tags, so a page can post
> successfully yet render differently than the source.

### To Do (7 tools)

| Tool                  | Description                                                   |
| --------------------- | ------------------------------------------------------------- |
| `list_todo_lists`     | List task lists                                               |
| `create_todo_list`    | Create a task list                                            |
| `list_todo_tasks`     | List tasks in a list                                          |
| `create_todo_task`    | Create a task, optionally repeating                           |
| `update_todo_task`    | Update a task, or change its repeat                           |
| `set_todo_task_state` | Set status, importance and your own fields, safely repeatable |
| `delete_todo_task`    | Delete a task (no API undo)                                   |

`create_todo_list` refuses a name an existing list already uses. Graph accepts a
duplicate `displayName` and returns a second list indistinguishable from the first in
every client, which scatters tasks across two lists nobody can tell apart.

`create_todo_task` and `update_todo_task` take an optional `recurrence`, mapping to
Graph's `patternedRecurrence`. Pass a `pattern` of `daily`, `weekly`,
`absoluteMonthly`, `relativeMonthly`, `absoluteYearly` or `relativeYearly`, with the
fields that pattern needs — `days_of_week` for weekly, `day_of_month` for
`absoluteMonthly`, and so on. Quarterly is `absoluteMonthly` with an `interval` of 3.

The repeat runs forever unless `range_type` is `endDate` or `numbered`. A recurring
task needs a `due_date`: To Do rolls the task forward from it, so without one the
task repeats but never appears in Today. Pass `clear_recurrence` on an update to
make a repeating task one-off again.

**Completing a recurring task does not mark that task completed.** Graph rolls the
same task id forward to the next due date and returns it as `notStarted`, while the
completed occurrence becomes a _separate_ task with a new id. The task id you hold
therefore refers to the live series, not one occurrence, and `update_todo_task` says
so in its result rather than reporting a bare "Task updated" over a `notStarted` task.

**Moving a repeating task's due date goes through its recurrence.** Graph reads a plain
due-date change on a repeating task as completing the current occurrence: the series
rolls forward one period and a second task appears on the new date. So on a repeating
task `update_todo_task` sends the new date as the recurrence's start date instead, and
the same task moves. To Do puts it on the first date the pattern allows from there — a
weekly-on-Saturday task moved to a Monday lands on the next Saturday — and the result
says so when that happens. Passing `due_date` with a different `recurrence.start_date`
is refused, since both set when the series starts. So is moving a task with an `endDate`
range past its end date: pass `recurrence` with a later `end_date`, or `clear_recurrence`.

**Tagging a task with its source makes a create safe to repeat.** `create_todo_task`
takes an optional `linked_resource` (`external_id`, `application_name`, and optionally
`web_url` and `display_name`), stored as the task's `linkedResources`. When a task in
the list already carries that `external_id`, the tool returns it and creates nothing, so
a caller that died between Graph's create and its own record of it can simply retry.
That check reads every page of the list first, completed tasks included, because Graph
can't filter on `linkedResources`; if the list can't be read in full, the create is
refused. An `external_id` is a lookup key, so one with a line break, `|`, `(`, `)` or a
comma, or over 200 characters, is refused, and a `web_url` must be `http:` or `https:`.
`display_name` defaults to `application_name`, since Graph lists it as required.

`create_todo_task` and `update_todo_task` also take `categories`, the names of the person's
Outlook categories. On an
update they replace the task's categories, and `[]` clears them.

**Task lines have a fixed shape** for callers that parse them:

```
- **<title>** [<status>] [<High|Low> importance] (Due: …) (Repeats: …) (Completed: <dateTime> <timeZone>) (Modified: …) (Categories: a, b) (Link: <externalId>) (ID: <id>)
```

A part Graph didn't return is left out, never filled with a default. Importance shows only
when it isn't `normal`, as `[High importance]` or `[Low importance]`, the way mail lines show it. Titles, categories and link ids are kept to one line, and lose `|`
and any parentheses that could end a `(…)` group early. The pattern is pinned in
`test/plugin-contract.spec.ts`. Graph refuses `$select` on To Do task lists, so these
reads fetch whole tasks, with `linkedResources` expanded. `list_todo_tasks` with
`fetch_all_pages`, and the duplicate check, read at most 50 pages of 100 tasks and fail past
that rather than return part of the list.

**A caller can keep its own data on a task** in an open extension, which the person can't
see or edit. `create_todo_task` and `set_todo_task_state` take `extension: { name, fields }`:
a reverse-DNS `name` (one fixed name per app, such as `com.example.assistant`) and at most 20
flat string, number or boolean fields, at most 2 KB as JSON. `list_todo_tasks` with
`extension_name` prints those fields at the end of each line as `(Ext: {...})`, leaving out
Graph's own name, id and type annotations. Graph replaces the whole extension on an update, so
`set_todo_task_state` reads it, merges the fields sent, and writes the whole set back; fields
it isn't sent keep their values.

**`set_todo_task_state` is the update for unattended use.** It changes only `status`,
`importance` and extension fields on one task, never the title, notes, dates or repeat
pattern, and it can't delete. The same call twice leaves the task as one call does, so it is
advertised as idempotent, and a failed call is safe to repeat. It is still marked destructive,
because it overwrites values. Completing a repeating task is refused, since each completion
rolls the series on by one more occurrence; use `update_todo_task` for that. When the task
updates but the extension doesn't, the error says which part was saved.

`delete_todo_task` has no undo through the API. Microsoft documents restoring a deleted
task by hand from Outlook's Deleted Items folder, but that hasn't been checked for tasks
deleted through Graph, so treat a delete as permanent. Because an id addresses the series, deleting a repeating task ends the whole
series; that requires `force: true`, and the refusal points at `clear_recurrence` for
the usual case of wanting the task to stop repeating without losing it.

### Auth & Utility (5 tools)

| Tool               | Description                            |
| ------------------ | -------------------------------------- |
| `get_auth_status`  | Check authentication status and scopes |
| `set_access_token` | Update token (client-token mode)       |
| `list_accounts`    | List registered accounts               |
| `switch_account`   | Switch default account                 |
| `graph_query`      | Execute arbitrary Graph API queries    |

### Auto-Pagination

All list tools support `fetch_all_pages: true` to automatically follow `@odata.nextLink` pagination (max 50 pages):

```json
{ "name": "list_messages", "arguments": { "fetch_all_pages": true } }
```

## Environment Variables

| Variable                  | Description                                                                                                                     | Default             |
| ------------------------- | ------------------------------------------------------------------------------------------------------------------------------- | ------------------- |
| `MS365_AUTH_MODE`         | Auth mode: `interactive`, `certificate`, `client-secret`, `client-token`, `oauth-proxy`                                         | `interactive`       |
| `MS365_TENANT_ID`         | Azure AD tenant ID                                                                                                              | `common`            |
| `MS365_CLIENT_ID`         | Azure AD application (client) ID                                                                                                | --                  |
| `MS365_CLIENT_SECRET`     | Client secret (for `client-secret` and `oauth-proxy` modes)                                                                     | --                  |
| `MS365_CERT_PATH`         | Certificate path (for `certificate` mode)                                                                                       | --                  |
| `MS365_CERT_PASSWORD`     | Certificate password (optional)                                                                                                 | --                  |
| `MS365_ACCESS_TOKEN`      | Initial access token (for `client-token` mode)                                                                                  | --                  |
| `MS365_USE_DEVICE_CODE`   | Skip the browser in `interactive` mode and use device code                                                                      | `false`             |
| `MS365_DISABLE_SIGN_IN`   | In `interactive` mode, fail instead of prompting when the cached sign-in cannot be reused                                       | `false`             |
| `MS365_TOKEN_CACHE`       | Persist the `interactive` token cache to disk; `false` keeps it in memory only                                                  | `true`              |
| `MS365_TOKEN_CACHE_PATH`  | Directory for the persisted token cache and account record                                                                      | `~/.config/...`     |
| `MS365_OAUTH_BASE_URL`    | Base URL for OAuth proxy mode                                                                                                   | --                  |
| `MS365_GRAPH_VERSION`     | Graph API version: `v1.0` or `beta`                                                                                             | `v1.0`              |
| `TRANSPORT_TYPE`          | Transport: `stdio` or `httpStream`                                                                                              | `stdio`             |
| `PORT`                    | HTTP server port                                                                                                                | `3000`              |
| `HOST`                    | HTTP server host                                                                                                                | `127.0.0.1`         |
| `MS365_PRESETS`           | Comma-separated presets: `personal`, `collaboration`, `productivity`, `rag`, `all`                                              | -- (all tools)      |
| `MS365_EXTRA_SCOPES`      | Comma-separated Graph scopes added to the requested set (OAuth proxy mode)                                                      | --                  |
| `MS365_MAX_EXTRACT_BYTES` | Ceiling over `read_document`'s per-format input caps, in bytes. Never raises them.                                              | -- (per-format)     |
| `MS365_ALLOWED_MAILBOXES` | Comma-separated addresses the `mailbox` parameter may target. Unset = any the user can reach when delegated, none when app-only | --                  |
| `MS365_ENABLED_TOOLS`     | Regex pattern to filter tools                                                                                                   | --                  |
| `MS365_READ_ONLY`         | Hide write tools                                                                                                                | `false`             |
| `MS365_ORG_MODE`          | Enable org-only tools (teams, chats, groups, planner)                                                                           | `false`             |
| `MS365_REQUIRE_DRAFT`     | Hide all `send_*` mail tools and block direct sends through `graph_query`; force the `create_*_draft` + `send_draft` flow       | `false`             |
| `TOKEN_STORAGE_PATH`      | Directory for persistent OAuth token storage                                                                                    | `/tmp/ms365-tokens` |
| `FASTMCP_HOST`            | Bind address for HTTP server (set `0.0.0.0` in containers)                                                                      | `localhost`         |

### Graph resilience

Every Graph call goes through a retry / timeout / circuit-breaker layer. The defaults suit
normal use; these knobs exist for tuning a deployment without a code change.

| Variable                          | Description                                                          | Default  |
| --------------------------------- | -------------------------------------------------------------------- | -------- |
| `MS365_GRAPH_MAX_RETRIES`         | Retries for a throttled (429) or transient (503/504/network) call    | `3`      |
| `MS365_GRAPH_TIMEOUT_MS`          | Per-attempt fetch timeout. Sized for slow large uploads              | `100000` |
| `MS365_GRAPH_BASE_BACKOFF_MS`     | Base for exponential backoff with full jitter                        | `200`    |
| `MS365_GRAPH_MAX_BACKOFF_MS`      | Backoff ceiling. A 429's `Retry-After` overrides it (capped at 60 s) | `5000`   |
| `MS365_GRAPH_CIRCUIT_THRESHOLD`   | Consecutive failures before the breaker opens                        | `5`      |
| `MS365_GRAPH_CIRCUIT_COOLDOWN_MS` | How long the breaker stays open before allowing a probe              | `30000`  |
| `MS365_GRAPH_CIRCUIT_DISABLED`    | Disable the breaker entirely                                         | `false`  |

Throttling (429) never counts toward the process-wide circuit breaker. Only
503/504 responses, timeouts and network errors count as breaker failures.

A 429 is retried on every method — Graph decides to throttle before it executes the
operation, so nothing has landed server-side. A 503/504/network failure is retried only
for idempotent methods (GET/HEAD/PUT/DELETE); retrying a POST or PATCH there could
duplicate a side effect that already succeeded, so those surface to the caller instead.
Uploads also avoid retries on 503/504/network failures: a renamed file may already
exist, or a chunk may already have advanced its upload session. They still honour
timeouts, circuit-breaker settings, and retries on 429.

When the breaker is open, calls fail fast as a `throttle` error carrying `retryAfter`,
rather than adding load to an upstream that is already failing.

## Claude Desktop (Local Installation)

### Option A: Desktop Extension (.mcpb)

1. Download the latest `.mcpb` file from [Releases](https://github.com/sapientsai/microsoft365-mcp-server/releases)
2. In Claude Desktop: **Settings → Extensions → Install Extension**
3. Select the `.mcpb` file
4. Enter your Azure App Client ID and Tenant ID when prompted
5. Authenticate in browser when first tool is used

### Option B: Manual Configuration

Add to your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "microsoft365": {
      "command": "npx",
      "args": ["-y", "microsoft365-mcp-server"],
      "env": {
        "MS365_AUTH_MODE": "interactive",
        "MS365_CLIENT_ID": "your-azure-app-client-id",
        "MS365_TENANT_ID": "your-tenant-id"
      }
    }
  }
}
```

Config file locations:

- macOS: `~/Library/Application Support/Claude/claude_desktop_config.json`
- Windows: `%APPDATA%\Claude\claude_desktop_config.json`
- Linux: `~/.config/Claude/claude_desktop_config.json`

## Docker / Remote Deployment

Deploy as a remote MCP server with per-user OAuth authentication:

```bash
docker compose up -d
```

Connect from Claude Desktop or any MCP client:

```json
{
  "mcpServers": {
    "microsoft365": {
      "url": "https://your-domain.example.com/mcp"
    }
  }
}
```

The Dockerfile sets `FASTMCP_HOST=0.0.0.0` (binds to all interfaces) and uses `/ping` for health checks (pre-auth endpoint). The `/health` endpoint provided by FastMCP is currently unreachable when auth is enabled due to a [known issue](https://github.com/punkpeye/mcp-proxy) in mcp-proxy's auth middleware ordering.

See **[DEPLOYMENT.md](DEPLOYMENT.md)** for the full guide — Docker, Azure AD app setup, Dokploy, reverse proxy, and security configuration.

## Development

```bash
pnpm install
pnpm validate        # format + lint + typecheck + test + build
pnpm dev             # development build with watch mode
pnpm inspect         # build and open MCP Inspector
```

## Architecture

- **[FastMCP](https://github.com/punkpeye/fastmcp)** — MCP server framework with Zod schema validation and built-in OAuth (AzureProvider)
- **[functype](https://github.com/jordanburke/functype)** — Functional programming: `Either` for error handling, `Option` for nullable fields, `Brand` for type-safe IDs
- **[ts-builds](https://github.com/jordanburke/ts-builds)** — Standardized TypeScript build toolchain
- **[@azure/identity](https://github.com/Azure/azure-sdk-for-js)** — Azure AD authentication
- Raw `fetch` with `Either`-based error handling (no Microsoft Graph SDK dependency)
- Data-driven tool registration with domain metadata, filtering, and MCP annotations
- `AsyncLocalStorage` for per-request token injection in OAuth proxy mode

## License

MIT

---

**Sponsored by <a href="https://sapientsai.com/"><img src="https://sapientsai.com/images/logo.svg" alt="SapientsAI" width="20" style="vertical-align: middle;"> SapientsAI</a>** — Building agentic AI for businesses

## Reading documents

`read_document` returns the readable text of a SharePoint or OneDrive file — PDF, DOCX, XLSX, or
anything text-based. Pair it with `search_site_files` (SharePoint, takes a `site_id`) or
`search_files` (your own OneDrive) to get an item ID, then pass
`/drives/{driveId}/items/{itemId}/content` or `/me/drive/items/{id}/content`.

Known limits:

1. **No OCR.** Extraction reads embedded text only, so a scanned PDF yields nothing. Fall back to
   `download_file`.
2. **Per-format input caps**, checked against the file's metadata before any download: 100 MB PDF,
   50 MB DOCX, 25 MB XLSX, 25 MB text. These are not uniform because memory cost is not uniform —
   a spreadsheet expands into a full workbook object model at roughly 10–20x its file size, which
   is why XLSX is the tightest rather than PDF. `MS365_MAX_EXTRACT_BYTES` tightens all four at once
   in a memory-constrained container; it is a ceiling and cannot raise a format above its default.
   Over the cap, use `download_file` for a download URL instead.
3. **Output is bounded separately** by `max_chars` (50,000 default, 200,000 max), with an explicit
   truncation marker. This is unrelated to the input caps above — do not tune them as one number.
4. **Discovery is split.** `search_files` is rooted at your own OneDrive and `search_site_files`
   takes a `site_id`. Neither is a tenant-wide search.

## Meeting transcripts

`list_meeting_transcripts` and `get_meeting_transcript` read the transcript of a Teams meeting, so
"turn this meeting into action items" works against a recorded review.

These are the only tools here that need a permission the server does not request by default, because
that permission needs a tenant administrator.

### 1. Grant the permission

The delegated permissions are:

| Scope                              | Needed for                                                                       | Consent |
| ---------------------------------- | -------------------------------------------------------------------------------- | ------- |
| `OnlineMeetingTranscript.Read.All` | listing transcripts and reading their content                                    | Admin   |
| `OnlineMeetings.Read`              | resolving a meeting from a `join_web_url` (skip if you always pass `meeting_id`) | Admin   |

Both are admin-consent scopes on the Graph service principal. A non-admin user cannot consent past
them, which is why they are **excluded from the default scope set** — adding them there would break
sign-in for every deployment whose tenant has not granted them.

Grant them on the app registration under **API permissions → Microsoft Graph → Delegated
permissions**, then **Grant admin consent**.

If tenant-wide consent is more than you want, a per-user grant works too: create the OAuth2
permission grant with `consentType: "Principal"` and a `principalId`, and only that user's tokens
carry the scope. This is the narrower option when one team needs transcripts and the rest of the
tenant does not.

The delegated path needs no Teams application access policy. (App-only access does — it additionally
requires `New-CsApplicationAccessPolicy` / `Grant-CsApplicationAccessPolicy` and reaches every
meeting in the tenant. This server uses the delegated path, which stays bounded by what the
signed-in user can already see.)

### 2. Ask the server to request it

```bash
MS365_EXTRA_SCOPES="OnlineMeetingTranscript.Read.All,OnlineMeetings.Read"
```

Scopes listed here are appended to the default set, so a granted permission reaches the issued
access token without a code change or a release. Users must sign in again afterwards — an
already-issued token does not gain a scope retroactively.

This applies to **OAuth proxy mode**. The credential-based modes request
`https://graph.microsoft.com/.default`, meaning "everything this app registration has been consented
to", so there the app registration alone decides and `MS365_EXTRA_SCOPES` is not consulted.

### Using the tools

```
list_meeting_transcripts(meeting_id: "MSo1N2Y5...")            # transcript IDs
get_meeting_transcript(meeting_id: "MSo1N2Y5...", transcript_id: "MSMjMCMj...")
```

`join_web_url` works in place of `meeting_id` — take it from a calendar event's
`onlineMeeting.joinUrl` — at the cost of one extra lookup and the `OnlineMeetings.Read` grant.

Worth knowing:

1. **Attendees, not just organizers.** Anyone on the meeting's calendar invite can read the
   transcript with their own token.
2. **The meeting needs a calendar event.** Meetings created through the create-onlineMeeting API
   without one are unsupported, and live events are excluded entirely. Expired meetings drop off the
   API too.
3. **Speaker attribution degrades rather than fails.** A tenant setting can disable speaker-attributed
   transcripts; asking for the attributed format then returns `403 SpeakerAttributionNotAllowed`. The
   tool retries automatically for the unattributed format and notes in its output that the names are
   missing. Pass `include_speaker_names: false` to skip the attributed attempt entirely.
4. **A separate tenant setting can block transcripts outright.** `GraphAccessToTranscriptsDisabled`
   has no request-side workaround — a Teams admin has to re-enable Graph API access to transcripts
   (`Set-CsTeamsMeetingConfiguration`). The tool says so rather than retrying.
5. **Output is bounded, and long transcripts are read in parts.** Each call returns at most
   `max_chars` (50,000 default). An hour's meeting runs to about 60,000 chars, so a truncated
   result ends with a marker naming the next start, e.g.
   `[truncated: chars 0–50,000 of 183,402; call again with offset: 50000]`. Pass that `offset` to
   continue; the last part ends with `end of transcript`.
6. **No metering.** These Teams APIs stopped being metered on August 25, 2025; no billing
   configuration is required.
