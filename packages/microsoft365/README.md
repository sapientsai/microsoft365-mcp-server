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

- **73 Tools** across 12 Microsoft 365 domains + generic Graph API escape hatch
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

### Safety Layers

| Layer                    | Protection                                               | Default            |
| ------------------------ | -------------------------------------------------------- | ------------------ |
| **User assignment**      | Only assigned users can authenticate                     | Off (enable above) |
| **Platform governance**  | Per-tool allow/confirm/deny in Claude Desktop Enterprise | Platform-level     |
| **Tool filtering**       | Presets, read-only, org-mode gating                      | All tools          |
| **Tenant restriction**   | `MS365_TENANT_ID` locks to one org                       | `common`           |
| **M365 native recovery** | Recycle bins, version history                            | Built-in           |

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

Org mode is required for Teams, Chats, Meetings, Groups, Planner, and user listing. Without it, these tools are hidden to prevent 403 errors on personal accounts.

## Available Tools

### Mail (16 tools)

| Tool                     | Description                                                              |
| ------------------------ | ------------------------------------------------------------------------ |
| `list_messages`          | List inbox messages with optional filtering                              |
| `get_message`            | Get a specific message with full body                                    |
| `search_messages`        | Search messages by query                                                 |
| `send_message`           | Send a new email                                                         |
| `send_reply`             | Reply to the sender and send now (threaded, original quoted)             |
| `send_reply_all`         | Reply to all recipients and send now (threaded, original quoted)         |
| `send_forward`           | Forward a message to new recipients and send now (original quoted)       |
| `create_draft`           | Create a new email draft in the Drafts folder                            |
| `create_reply_draft`     | Create a reply draft — threaded, with the original quoted underneath     |
| `create_reply_all_draft` | Create a reply-all draft — threaded, with the original quoted underneath |
| `create_forward_draft`   | Create a forward draft — original quoted, recipients you specify         |
| `send_draft`             | Send an existing email draft                                             |
| `list_attachments`       | List a message's attachments, with a read_document path for file ones    |
| `list_mail_folders`      | List top-level mail folders with item, unread and subfolder counts       |
| `move_message`           | Move a message to a well-known folder, a folder name, or a folder ID     |
| `batch_move_messages`    | Move up to 50 messages to one folder in a single call                    |

> The `create_*_draft` tools produce a properly threaded draft (same conversation, full
> quoted history) for review, then send via `send_draft`. They remain available under
> `MS365_REQUIRE_DRAFT=true`; the `send_*` tools are hidden in that mode.

### Calendar (7 tools)

| Tool                        | Description                                                     |
| --------------------------- | --------------------------------------------------------------- |
| `list_events`               | List calendar events                                            |
| `list_calendar_view`        | List event instances in a date range (expands recurring series) |
| `find_meeting_availability` | Suggest meeting times where all participants are free           |
| `get_event`                 | Get event details                                               |
| `create_event`              | Create a new event                                              |
| `update_event`              | Update an existing event                                        |
| `delete_event`              | Delete an event                                                 |

### Contacts (4 tools)

| Tool              | Description          |
| ----------------- | -------------------- |
| `list_contacts`   | List contacts        |
| `get_contact`     | Get contact details  |
| `create_contact`  | Create a new contact |
| `search_contacts` | Search contacts      |

### Files / OneDrive (7 tools)

| Tool               | Description                                                                   |
| ------------------ | ----------------------------------------------------------------------------- |
| `list_drive_items` | List files and folders (supports `folder_id` or `folder_path` for navigation) |
| `get_drive_item`   | Get file/folder metadata                                                      |
| `search_files`     | Search OneDrive/SharePoint                                                    |
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
| `search_site_files` | Search files within a SharePoint site                                   |

SharePoint tools use **delegated permissions** — users see only the sites and files they have access to. Private channel sites are properly isolated; access requires channel membership.

### Chats (3 tools, org mode)

| Tool                 | Description                                            |
| -------------------- | ------------------------------------------------------ |
| `list_chats`         | List Teams chats (1:1, group, meeting)                 |
| `list_chat_messages` | List messages in a chat                                |
| `send_chat_message`  | Send a message in a chat. Use `48:notes` for self-chat |

### Teams (4 tools, org mode)

| Tool                    | Description                  |
| ----------------------- | ---------------------------- |
| `list_teams`            | List joined teams            |
| `list_channels`         | List channels in a team      |
| `list_channel_messages` | List recent channel messages |
| `send_channel_message`  | Send a message to a channel  |

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

### To Do (4 tools)

| Tool               | Description          |
| ------------------ | -------------------- |
| `list_todo_lists`  | List task lists      |
| `list_todo_tasks`  | List tasks in a list |
| `create_todo_task` | Create a new task    |
| `update_todo_task` | Update a task        |

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

| Variable                  | Description                                                                             | Default             |
| ------------------------- | --------------------------------------------------------------------------------------- | ------------------- |
| `MS365_AUTH_MODE`         | Auth mode: `interactive`, `certificate`, `client-secret`, `client-token`, `oauth-proxy` | `interactive`       |
| `MS365_TENANT_ID`         | Azure AD tenant ID                                                                      | `common`            |
| `MS365_CLIENT_ID`         | Azure AD application (client) ID                                                        | --                  |
| `MS365_CLIENT_SECRET`     | Client secret (for `client-secret` and `oauth-proxy` modes)                             | --                  |
| `MS365_CERT_PATH`         | Certificate path (for `certificate` mode)                                               | --                  |
| `MS365_CERT_PASSWORD`     | Certificate password (optional)                                                         | --                  |
| `MS365_ACCESS_TOKEN`      | Initial access token (for `client-token` mode)                                          | --                  |
| `MS365_OAUTH_BASE_URL`    | Base URL for OAuth proxy mode                                                           | --                  |
| `MS365_GRAPH_VERSION`     | Graph API version: `v1.0` or `beta`                                                     | `v1.0`              |
| `TRANSPORT_TYPE`          | Transport: `stdio` or `httpStream`                                                      | `stdio`             |
| `PORT`                    | HTTP server port                                                                        | `3000`              |
| `HOST`                    | HTTP server host                                                                        | `127.0.0.1`         |
| `MS365_PRESETS`           | Comma-separated presets: `personal`, `collaboration`, `productivity`, `rag`, `all`      | -- (all tools)      |
| `MS365_EXTRA_SCOPES`      | Comma-separated Graph scopes added to the requested set (OAuth proxy mode)              | --                  |
| `MS365_MAX_EXTRACT_BYTES` | Ceiling over `read_document`'s per-format input caps, in bytes. Never raises them.      | -- (per-format)     |
| `MS365_ENABLED_TOOLS`     | Regex pattern to filter tools                                                           | --                  |
| `MS365_READ_ONLY`         | Hide write tools                                                                        | `false`             |
| `MS365_ORG_MODE`          | Enable org-only tools (teams, chats, groups, planner)                                   | `false`             |
| `MS365_REQUIRE_DRAFT`     | Hide all `send_*` mail tools; force the `create_*_draft` + `send_draft` flow            | `false`             |
| `TOKEN_STORAGE_PATH`      | Directory for persistent OAuth token storage                                            | `/tmp/ms365-tokens` |
| `FASTMCP_HOST`            | Bind address for HTTP server (set `0.0.0.0` in containers)                              | `localhost`         |

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
5. **Output is bounded** by `max_chars` (50,000 default) with a truncation marker, like
   `read_document`.
6. **No metering.** These Teams APIs stopped being metered on August 25, 2025; no billing
   configuration is required.
