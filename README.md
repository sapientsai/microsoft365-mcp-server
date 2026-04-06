## microsoft365-mcp-server

[![Node.js CI](https://github.com/sapientsai/microsoft365-mcp-server/actions/workflows/node.js.yml/badge.svg)](https://github.com/sapientsai/microsoft365-mcp-server/actions/workflows/node.js.yml)
[![npm version](https://img.shields.io/npm/v/microsoft365-mcp-server.svg)](https://www.npmjs.com/package/microsoft365-mcp-server)

A Model Context Protocol (MCP) server for Microsoft 365 — manage email, calendar, contacts, files, Teams chats, channels, Planner, OneNote, To Do, users, and groups via Microsoft Graph API.

## Features

- **58 Tools** across 12 Microsoft 365 domains + generic Graph API escape hatch
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

| Permission                                     | Domain           |
| ---------------------------------------------- | ---------------- |
| `User.Read`                                    | User profile     |
| `Mail.Read`, `Mail.Send`                       | Email            |
| `Calendars.ReadWrite`                          | Calendar         |
| `Contacts.Read`                                | Contacts         |
| `Files.ReadWrite`                              | OneDrive         |
| `Sites.Read.All`, `Sites.ReadWrite.All`        | SharePoint sites |
| `Chat.ReadWrite`                               | Teams chats      |
| `ChatMessage.Read`, `ChatMessage.Send`         | Chat messages    |
| `Team.ReadBasic.All`                           | Teams            |
| `Channel.ReadBasic.All`, `ChannelMessage.Send` | Channels         |
| `Tasks.ReadWrite`                              | Planner & To Do  |
| `Notes.Read`                                   | OneNote          |

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
| `collaboration` | chats, teams, planner, groups                  |
| `productivity`  | mail, calendar, todo                           |
| `all`           | everything                                     |

```bash
MS365_PRESETS=personal                    # just personal tools
MS365_PRESETS=personal,collaboration      # personal + team tools
```

If not set, all tools are registered.

### Other Filters

```bash
MS365_ENABLED_TOOLS="mail|calendar"   # regex pattern — only matching tools registered
MS365_READ_ONLY=true                  # hide all write tools (send, create, update, delete)
MS365_ORG_MODE=true                   # enable org-only tools (teams, chats, groups, planner, list_users)
```

Org mode is required for Teams, Chats, Groups, Planner, and user listing. Without it, these tools are hidden to prevent 403 errors on personal accounts.

## Available Tools

### Mail (7 tools)

| Tool               | Description                                   |
| ------------------ | --------------------------------------------- |
| `list_messages`    | List inbox messages with optional filtering   |
| `get_message`      | Get a specific message with full body         |
| `send_message`     | Send a new email                              |
| `reply_to_message` | Reply to a message                            |
| `search_messages`  | Search messages by query                      |
| `create_draft`     | Create a new email draft in the Drafts folder |
| `send_draft`       | Send an existing email draft                  |

### Calendar (5 tools)

| Tool           | Description              |
| -------------- | ------------------------ |
| `list_events`  | List calendar events     |
| `get_event`    | Get event details        |
| `create_event` | Create a new event       |
| `update_event` | Update an existing event |
| `delete_event` | Delete an event          |

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

### OneNote (4 tools)

| Tool               | Description                 |
| ------------------ | --------------------------- |
| `list_notebooks`   | List notebooks              |
| `list_sections`    | List sections in a notebook |
| `list_pages`       | List pages in a section     |
| `get_page_content` | Get page content as HTML    |

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

| Variable               | Description                                                                             | Default             |
| ---------------------- | --------------------------------------------------------------------------------------- | ------------------- |
| `MS365_AUTH_MODE`      | Auth mode: `interactive`, `certificate`, `client-secret`, `client-token`, `oauth-proxy` | `interactive`       |
| `MS365_TENANT_ID`      | Azure AD tenant ID                                                                      | `common`            |
| `MS365_CLIENT_ID`      | Azure AD application (client) ID                                                        | --                  |
| `MS365_CLIENT_SECRET`  | Client secret (for `client-secret` and `oauth-proxy` modes)                             | --                  |
| `MS365_CERT_PATH`      | Certificate path (for `certificate` mode)                                               | --                  |
| `MS365_CERT_PASSWORD`  | Certificate password (optional)                                                         | --                  |
| `MS365_ACCESS_TOKEN`   | Initial access token (for `client-token` mode)                                          | --                  |
| `MS365_OAUTH_BASE_URL` | Base URL for OAuth proxy mode                                                           | --                  |
| `MS365_GRAPH_VERSION`  | Graph API version: `v1.0` or `beta`                                                     | `v1.0`              |
| `TRANSPORT_TYPE`       | Transport: `stdio` or `httpStream`                                                      | `stdio`             |
| `PORT`                 | HTTP server port                                                                        | `3000`              |
| `HOST`                 | HTTP server host                                                                        | `127.0.0.1`         |
| `MS365_PRESETS`        | Comma-separated presets: `personal`, `collaboration`, `productivity`, `all`             | -- (all tools)      |
| `MS365_ENABLED_TOOLS`  | Regex pattern to filter tools                                                           | --                  |
| `MS365_READ_ONLY`      | Hide write tools                                                                        | `false`             |
| `MS365_ORG_MODE`       | Enable org-only tools (teams, chats, groups, planner)                                   | `false`             |
| `TOKEN_STORAGE_PATH`   | Directory for persistent OAuth token storage                                            | `/tmp/ms365-tokens` |
| `FASTMCP_HOST`         | Bind address for HTTP server (set `0.0.0.0` in containers)                              | `localhost`         |

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
