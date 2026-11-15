# Microsoft 365 MCP Server

[![npm version](https://img.shields.io/npm/v/microsoft365-mcp-server.svg)](https://www.npmjs.com/package/microsoft365-mcp-server)
[![Node.js CI](https://github.com/sapientsai/microsoft365-mcp-server/actions/workflows/node.js.yml/badge.svg)](https://github.com/sapientsai/microsoft365-mcp-server/actions/workflows/node.js.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

An [MCP](https://modelcontextprotocol.io) server that lets an AI assistant work in a user's Microsoft 365: mail,
calendar, contacts, OneDrive and SharePoint files, Teams chats and channels, Planner, OneNote and To Do, all through
Microsoft Graph.

It signs in **as the user** (delegated OAuth), so it reaches only what that user can already reach. Run it locally over
stdio, or host it as a remote server with per-user sign-in.

## Quick start

**1. Register an app in Microsoft Entra ID.** The server needs an application (client) ID. Create one under
[App registrations](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade), add the redirect
URI `http://localhost` on the _Mobile and desktop_ platform, and grant the delegated Graph permissions for the domains
you want. To keep `MS365_TENANT_ID` as `common`, choose an account type that covers every organization, plus personal
accounts if you need them. For one organization only, choose single tenant and set `MS365_TENANT_ID` to its tenant ID.
The [full checklist](packages/microsoft365/README.md#azure-ad-app-registration) lists every permission.

**2. Add the server to your MCP client** (Claude Desktop, Claude Code, VS Code, Cursor, …):

```json
{
  "mcpServers": {
    "microsoft365": {
      "command": "npx",
      "args": ["-y", "microsoft365-mcp-server"],
      "env": {
        "MS365_CLIENT_ID": "your-app-client-id",
        "MS365_TENANT_ID": "common"
      }
    }
  }
}
```

**3. Start your client.** When the client starts the server, a browser opens to sign in. On a headless machine, set
`MS365_USE_DEVICE_CODE=true` instead: the server prints a device code to its log, and the app registration needs
_Allow public client flows_ turned on. The sign-in is cached, so it survives restarts.

## What it can do

**99 tools.** 27 of them need a work or school tenant, so they stay hidden until you set
`MS365_ORG_MODE=true`.

| Area                  | Tools | What the assistant can do                                                      |
| --------------------- | ----: | ------------------------------------------------------------------------------ |
| Mail                  |    23 | Read, search, send, reply, forward, move, and draft messages; save attachments |
| Calendar              |     9 | List and manage events, find free time, accept or decline invitations          |
| Contacts              |     7 | List, search, create, update, delete, and export contacts                      |
| Files (OneDrive)      |     8 | Browse, search, download, and upload files; create folders                     |
| SharePoint ¹          |     5 | Browse sites, their drives and items; search a site's files                    |
| Documents             |     1 | Extract readable text from PDF, DOCX, XLSX, and text files (`read_document`)   |
| OneNote               |    10 | List notebooks, sections and pages; read, create, update, copy, delete pages   |
| To Do                 |     8 | Manage task lists and tasks, including repeating ones                          |
| Teams chats ¹         |     3 | List chats and messages; send a chat message                                   |
| Teams channels ¹      |     4 | List teams, channels and messages; post to a channel                           |
| Meeting transcripts ¹ |     2 | List and read Teams meeting transcripts                                        |
| Planner ¹             |     8 | List plans, buckets and tasks; create and update tasks                         |
| Users                 |     3 | Your own profile; look up other users ¹                                        |
| Groups ¹              |     3 | List groups and their members                                                  |
| Accounts              |     4 | Check sign-in status, switch between accounts, supply a token                  |
| Graph query           |     1 | Call any Microsoft Graph endpoint the user has access to (`graph_query`)       |

¹ Needs `MS365_ORG_MODE=true`. In the Users row, only the lookup of other users does; your own profile does not.

A few tools exist on only one transport: local file paths work over stdio, and upload links work over HTTP. Every tool
carries MCP annotations (read-only, destructive, idempotent), so a client can tell reads from writes. The
[tool reference](packages/microsoft365/README.md#available-tools) lists each one.

## Keeping it safe

You decide how much the assistant can touch:

- **`MS365_READ_ONLY=true`** hides every tool that writes.
- **`MS365_REQUIRE_DRAFT=true`** removes direct email sending. The assistant can only write email drafts, which you
  review in Outlook and send yourself, or approve with `send_draft`. Teams messages and calendar invites still send.
- **`MS365_PRESETS`** loads only some areas: `personal`, `collaboration`, `productivity`, `rag`, or `all`.
- **`MS365_ENABLED_TOOLS`** loads only the tools whose names match a regex.
- **`MS365_ALLOWED_MAILBOXES`** limits which other people's or shared mailboxes it may open.
- **`MS365_TENANT_ID`** locks sign-in to one organization.

See [Safety Layers](packages/microsoft365/README.md#safety-layers) for what Microsoft 365 can recover after a mistake,
and what it can't.

## Sign-in modes

| Mode          | `MS365_AUTH_MODE` | Use it when                                                     |
| ------------- | ----------------- | --------------------------------------------------------------- |
| Interactive   | `interactive`     | One person on their own machine (the default)                   |
| OAuth proxy   | `oauth-proxy`     | A hosted server where each user signs in through the MCP client |
| Client token  | `client-token`    | Your app already holds a Graph token and passes it in           |
| Client secret | `client-secret`   | Automation running as an app registration, with no user present |
| Certificate   | `certificate`     | The same, using a certificate instead of a secret               |

Each mode's settings are in [Authentication](packages/microsoft365/README.md#authentication). The main
environment variables are listed in [Environment Variables](packages/microsoft365/README.md#environment-variables).

## Hosting it as a remote server

`oauth-proxy` mode turns the server into an OAuth 2.1 authorization server in front of Entra ID. Each user signs in
from their own MCP client, and each request runs with that user's own Graph token. The repo ships a `Dockerfile`
and `docker-compose.yml`. Set `MS365_CLIENT_ID`, `MS365_CLIENT_SECRET` and `MS365_OAUTH_BASE_URL` first; the compose file
refuses to start without them:

```bash
docker compose up -d
```

Then point the client at `https://your-host/mcp`. [DEPLOYMENT.md](DEPLOYMENT.md) covers the app registration,
reverse proxy, Dokploy, and hardening.

## Also in this repo

This is a pnpm monorepo. The server above is [`packages/microsoft365`](packages/microsoft365), published to npm as
`microsoft365-mcp-server`. Alongside it:

- **[`packages/graph`](packages/graph)**: an app-only Microsoft Graph server for headless, tenant-wide document search
  and retrieval, with no user signed in. It ships as the Docker image `ghcr.io/sapientsai/microsoft-mcp-server` and
  replaces the archived [`sapientsai/microsoft-mcp-server`](https://github.com/sapientsai/microsoft-mcp-server).
- **[`packages/core`](packages/core)**: Graph plumbing both servers share, such as uploads, paging, errors, and auth
  types. Internal, not published.
- **[`packages/extract`](packages/extract)**: turns PDF, DOCX, XLSX, and text files into plain text. Internal, not
  published.

Pick by who signs in: `microsoft365` when a person is there to consent, `graph` when only a service credential is.

## Development

Requires Node.js 22.16 or later and pnpm.

```bash
pnpm install
pnpm validate   # format, lint, test, and build every package
pnpm --filter microsoft365-mcp-server inspect   # build and open the MCP Inspector
```

## License

[MIT](LICENSE). Some shared code is adapted from other MIT projects; see [NOTICE](NOTICE).

---

**Sponsored by <a href="https://sapientsai.com/"><img src="https://sapientsai.com/images/logo.svg" alt="SapientsAI" width="20" style="vertical-align: middle;"> SapientsAI</a>** — Building agentic AI for businesses
