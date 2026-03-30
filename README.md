## ms365-mcp-server

[![Node.js CI](https://github.com/sapientsai/ms365-mcp-server/actions/workflows/node.js.yml/badge.svg)](https://github.com/sapientsai/ms365-mcp-server/actions/workflows/node.js.yml)
[![npm version](https://img.shields.io/npm/v/ms365-mcp-server.svg)](https://www.npmjs.com/package/ms365-mcp-server)

A Model Context Protocol (MCP) server for Microsoft 365 — manage email, calendar, contacts, files, Teams, Planner, OneNote, To Do, users, and groups via Microsoft Graph API.

## Features

- **46 Tools** across 10 Microsoft 365 domains + generic Graph API escape hatch
- **4 Auth Modes**: Interactive (device code), certificate, client secret, client-provided token
- **Functional Programming**: [functype](https://github.com/jordanburke/functype) patterns — `Either`, `Option`, `Try`, `Brand` types
- **Type-Safe**: Branded IDs, Zod parameter schemas, strict TypeScript
- **Modern Build System**: [ts-builds](https://github.com/jordanburke/ts-builds) + [tsdown](https://tsdown.dev/)
- **Dual Transport**: stdio (default) and HTTP stream

## Quick Start

```bash
# Install globally
npm install -g ms365-mcp-server

# Or run directly
npx ms365-mcp-server
```

### Claude Desktop / VS Code Configuration

Add to your `claude_desktop_config.json` or MCP settings:

```json
{
  "mcpServers": {
    "ms-365": {
      "command": "npx",
      "args": ["-y", "ms365-mcp-server"],
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

### Azure AD App Registration

You need an Azure AD app registration with the appropriate Microsoft Graph permissions:

1. Go to [Azure Portal](https://portal.azure.com) > App registrations > New registration
2. Set redirect URI to `http://localhost:3000` (for interactive mode)
3. Add Microsoft Graph API permissions for the domains you need:
   - `Mail.Read`, `Mail.Send` — Email
   - `Calendars.ReadWrite` — Calendar
   - `Contacts.Read` — Contacts
   - `Files.Read` — OneDrive/SharePoint
   - `Team.ReadBasic.All` — Teams
   - `Tasks.ReadWrite` — Planner & To Do
   - `Notes.Read` — OneNote
   - `User.Read` — User profile

## Available Tools

### Mail (5 tools)

| Tool               | Description                                 |
| ------------------ | ------------------------------------------- |
| `list_messages`    | List inbox messages with optional filtering |
| `get_message`      | Get a specific message with full body       |
| `send_message`     | Send a new email                            |
| `reply_to_message` | Reply to a message                          |
| `search_messages`  | Search messages by query                    |

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

### Files / OneDrive (5 tools)

| Tool               | Description                        |
| ------------------ | ---------------------------------- |
| `list_drive_items` | List files and folders             |
| `get_drive_item`   | Get file/folder metadata           |
| `search_files`     | Search OneDrive/SharePoint         |
| `download_file`    | Get file metadata and download URL |
| `create_folder`    | Create a new folder                |

### Teams (4 tools)

| Tool                    | Description                  |
| ----------------------- | ---------------------------- |
| `list_teams`            | List joined teams            |
| `list_channels`         | List channels in a team      |
| `list_channel_messages` | List recent channel messages |
| `send_channel_message`  | Send a message to a channel  |

### Users & Groups (6 tools)

| Tool                 | Description                      |
| -------------------- | -------------------------------- |
| `get_me`             | Get authenticated user's profile |
| `list_users`         | List organization users          |
| `get_user`           | Get a specific user's profile    |
| `list_groups`        | List organization groups         |
| `get_group`          | Get group details                |
| `list_group_members` | List group members               |

### Planner (5 tools)

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

### Auth & Generic (3 tools)

| Tool               | Description                            |
| ------------------ | -------------------------------------- |
| `get_auth_status`  | Check authentication status and scopes |
| `set_access_token` | Update token (client-token mode)       |
| `graph_query`      | Execute arbitrary Graph API queries    |

## Environment Variables

| Variable              | Description                                                              | Default       |
| --------------------- | ------------------------------------------------------------------------ | ------------- |
| `MS365_AUTH_MODE`     | Auth mode: `interactive`, `certificate`, `client-secret`, `client-token` | `interactive` |
| `MS365_TENANT_ID`     | Azure AD tenant ID                                                       | `common`      |
| `MS365_CLIENT_ID`     | Azure AD application (client) ID                                         | —             |
| `MS365_CLIENT_SECRET` | Client secret (for `client-secret` mode)                                 | —             |
| `MS365_CERT_PATH`     | Certificate path (for `certificate` mode)                                | —             |
| `MS365_CERT_PASSWORD` | Certificate password (optional)                                          | —             |
| `MS365_ACCESS_TOKEN`  | Initial access token (for `client-token` mode)                           | —             |
| `MS365_GRAPH_VERSION` | Graph API version: `v1.0` or `beta`                                      | `v1.0`        |
| `TRANSPORT_TYPE`      | Transport: `stdio` or `httpStream`                                       | `stdio`       |
| `PORT`                | HTTP server port                                                         | `3000`        |
| `HOST`                | HTTP server host                                                         | `127.0.0.1`   |

## Development

```bash
pnpm install
pnpm validate        # format + lint + typecheck + test + build
pnpm dev             # development build with watch mode
pnpm inspect         # build and open MCP Inspector
```

## Architecture

Built on the same patterns as [dakboard-mcp-server](https://github.com/jordanburke/dakboard-mcp-server):

- **[FastMCP](https://github.com/punkpeye/fastmcp)** — MCP server framework with Zod schema validation
- **[functype](https://github.com/jordanburke/functype)** — Functional programming: `Either` for error handling, `Option` for nullable fields, `Brand` for type-safe IDs
- **[ts-builds](https://github.com/jordanburke/ts-builds)** — Standardized TypeScript build toolchain
- **[@azure/identity](https://github.com/Azure/azure-sdk-for-js)** — Azure AD authentication
- Raw `fetch` with `Either`-based error handling (no Microsoft Graph SDK dependency)

Inspired by [lokka](https://github.com/merill/lokka) but with discrete typed tools per domain instead of a single passthrough tool.

## License

MIT

---

**Sponsored by <a href="https://sapientsai.com/"><img src="https://sapientsai.com/images/logo.svg" alt="SapientsAI" width="20" style="vertical-align: middle;"> SapientsAI</a>** — Building agentic AI for businesses
