# Deployment Guide

Deploy microsoft365-mcp-server as a remote MCP server with OAuth authentication. Users connect via Claude Desktop (or any MCP client) and authenticate with their own Microsoft 365 account.

## How It Works

1. You deploy the server with your Azure AD app credentials
2. Users add the server URL to their MCP client (Claude Desktop, VS Code, etc.)
3. On first use, the server redirects to Microsoft login
4. FastMCP's built-in AzureProvider handles OAuth (PKCE, consent, JWT, token refresh)
5. Each request carries the user's own token — they only access their own data

## Prerequisites

- An Azure AD (Entra ID) app registration ([setup guide](#azure-ad-app-setup))
- Docker (or Node.js 22+ for bare metal)
- A public URL with HTTPS (for the OAuth callback)

## Quick Start with Docker Compose

### 1. Create a `.env` file

```env
MS365_CLIENT_ID=your-azure-app-client-id
MS365_CLIENT_SECRET=your-azure-app-client-secret
MS365_TENANT_ID=common
MS365_OAUTH_BASE_URL=https://your-server.example.com
MS365_ORG_MODE=true
```

- `MS365_TENANT_ID=common` allows any Microsoft account (personal + org)
- Set to a specific tenant ID to restrict to one organization
- `MS365_ORG_MODE=true` enables Teams, Chats, Groups, and Planner tools

### 2. Run

```bash
docker compose up -d
```

The server starts on port 8080 with OAuth proxy mode.

### 3. Connect from Claude Desktop

Add to your Claude Desktop MCP settings:

```json
{
  "mcpServers": {
    "microsoft365": {
      "url": "https://your-server.example.com/mcp"
    }
  }
}
```

Claude Desktop will handle the OAuth redirect automatically.

## Docker Build Only

```bash
docker build -t microsoft365-mcp-server .
docker run -p 8080:8080 \
  -e MS365_AUTH_MODE=oauth-proxy \
  -e MS365_CLIENT_ID=your-client-id \
  -e MS365_CLIENT_SECRET=your-client-secret \
  -e MS365_TENANT_ID=common \
  -e MS365_OAUTH_BASE_URL=https://your-server.example.com \
  -e MS365_ORG_MODE=true \
  microsoft365-mcp-server
```

## Bare Metal (No Docker)

```bash
npm install -g microsoft365-mcp-server

MS365_AUTH_MODE=oauth-proxy \
MS365_CLIENT_ID=your-client-id \
MS365_CLIENT_SECRET=your-client-secret \
MS365_TENANT_ID=common \
MS365_OAUTH_BASE_URL=https://your-server.example.com \
TRANSPORT_TYPE=httpStream \
PORT=8080 \
MS365_ORG_MODE=true \
  microsoft365-mcp-server
```

## Azure AD App Setup

### 1. Create the app registration

```bash
az ad app create --display-name "Microsoft365-MCP-Server" \
  --sign-in-audience AzureADandPersonalMicrosoftAccount
```

Or via [Azure Portal](https://portal.azure.com) > App registrations > New registration:

- Name: `Microsoft365-MCP-Server`
- Supported account types: "Accounts in any organizational directory and personal Microsoft accounts"

### 2. Add redirect URIs

Two platforms needed:

**Mobile/Desktop** (for local development with interactive mode):

```bash
az ad app update --id <app-id> --public-client-redirect-uris "http://localhost"
```

**Web** (for OAuth proxy in production):

```bash
az ad app update --id <app-id> --web-redirect-uris "https://your-server.example.com/oauth/callback"
```

For local OAuth testing, also add `http://localhost:3000/oauth/callback`.

### 3. Add delegated permissions

```bash
# Get the app ID from step 1, then run:
az ad app update --id <app-id> --required-resource-accesses '[{
  "resourceAppId": "00000003-0000-0000-c000-000000000000",
  "resourceAccess": [
    {"id": "e1fe6dd8-ba31-4d61-89e7-88639da4683d", "type": "Scope"},
    {"id": "a154be20-db9c-4678-8ab7-66f6cc099a59", "type": "Scope"},
    {"id": "570282fd-fa5c-430d-a7fd-fc8dc98a9dca", "type": "Scope"},
    {"id": "e383f46e-2787-4529-855e-0e479a3ffac0", "type": "Scope"},
    {"id": "1ec239c2-d7c9-4623-a91a-a9775856bb36", "type": "Scope"},
    {"id": "ff74d97f-43af-4b68-9f2a-b77f2692e6e0", "type": "Scope"},
    {"id": "5c28f0bf-8a70-41f1-8ab2-9032436ddb65", "type": "Scope"},
    {"id": "9ff7295e-131b-4d94-90e1-69fde507ac11", "type": "Scope"},
    {"id": "cdcdac3a-fd45-410d-83ef-554db620e5c7", "type": "Scope"},
    {"id": "116b7235-7cc6-461e-b163-8e55691d839e", "type": "Scope"},
    {"id": "767156cb-16ae-4d10-8f8b-41b657c8c8c8", "type": "Scope"},
    {"id": "7b2449af-571a-4f28-956b-2dae53ee55e3", "type": "Scope"},
    {"id": "bb8f0e85-c1ed-4e2c-9b72-c95a166c5588", "type": "Scope"},
    {"id": "ebf0f66e-9fb1-49e4-a278-222f76911cf4", "type": "Scope"},
    {"id": "2219042f-cab5-40cc-b0d2-16b1540b4c5f", "type": "Scope"},
    {"id": "371361e4-b9e2-4a3f-8315-2a301a3b0a3d", "type": "Scope"},
    {"id": "5f8c59db-677d-491f-a6b8-5f174b11ec1d", "type": "Scope"}
  ]
}]'
```

These map to:

| GUID          | Permission              |
| ------------- | ----------------------- |
| `e1fe6dd8...` | User.Read               |
| `a154be20...` | User.Read.All           |
| `570282fd...` | Mail.Read               |
| `e383f46e...` | Mail.Send               |
| `1ec239c2...` | Calendars.ReadWrite     |
| `ff74d97f...` | Contacts.Read           |
| `5c28f0bf...` | Files.ReadWrite         |
| `9ff7295e...` | Chat.ReadWrite          |
| `cdcdac3a...` | ChatMessage.Read        |
| `116b7235...` | ChatMessage.Send        |
| `767156cb...` | ChannelMessage.Read.All |
| `7b2449af...` | Team.ReadBasic.All      |
| `bb8f0e85...` | Channel.ReadBasic.All   |
| `ebf0f66e...` | ChannelMessage.Send     |
| `2219042f...` | Tasks.ReadWrite         |
| `371361e4...` | Notes.Read              |
| `5f8c59db...` | Group.Read.All          |

### 4. Grant admin consent (org tenants only)

```bash
az ad app permission admin-consent --id <app-id>
```

Skip this for multi-tenant apps where each org grants their own consent.

### 5. Create a client secret

```bash
az ad app credential reset --id <app-id> --display-name "mcp-server" --years 1
```

Save the `password` field — this is your `MS365_CLIENT_SECRET`.

## Deployment Configurations

### Read-Only Mode (Safe for demos)

```env
MS365_READ_ONLY=true
```

Only list/get/search tools are registered. No send, create, update, or delete.

### Executive Preset (CEO use case)

```env
MS365_PRESETS=personal
MS365_ORG_MODE=false
```

Mail, calendar, contacts, To Do, files, OneNote. No Teams/chat exposure.

### Full Access with Org Tools

```env
MS365_ORG_MODE=true
```

All 58 tools enabled across all domains.

## Reverse Proxy

The server needs to be behind HTTPS for OAuth to work in production. Example nginx config:

```nginx
server {
    listen 443 ssl;
    server_name ms365-mcp.example.com;

    ssl_certificate /etc/ssl/certs/cert.pem;
    ssl_certificate_key /etc/ssl/private/key.pem;

    location / {
        proxy_pass http://localhost:8080;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection "upgrade";
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```

## Security Considerations

- **Client secret**: Store in environment variables or a secrets manager (Azure Key Vault, etc.). Never commit to git.
- **HTTPS required**: OAuth callbacks must use HTTPS in production. Use a reverse proxy or managed platform.
- **Tenant restriction**: Set `MS365_TENANT_ID` to your org's tenant ID to prevent other organizations from authenticating.
- **Read-only mode**: Use `MS365_READ_ONLY=true` for demos or untrusted environments.
- **Token storage**: FastMCP stores OAuth tokens in memory by default. Tokens are lost on server restart (users re-authenticate). For persistence, configure a `DiskStore` or custom `TokenStorage`.

## Dokploy Deployment

If you're using [Dokploy](https://dokploy.com/):

1. Create a new application from the GitHub repo
2. Set build type to Dockerfile
3. Add environment variables in the Environment tab:
   - `MS365_AUTH_MODE=oauth-proxy`
   - `MS365_CLIENT_ID=<your-client-id>`
   - `MS365_CLIENT_SECRET=<your-client-secret>`
   - `MS365_TENANT_ID=<your-tenant-id>`
   - `MS365_OAUTH_BASE_URL=https://<your-dokploy-domain>`
   - `MS365_ORG_MODE=true`
   - `PORT=8080`
   - `TRANSPORT_TYPE=httpStream`
4. Set the port to 8080
5. Configure the domain (e.g., `ms365-mcp.civala.ai`)
6. Deploy

Don't forget to add `https://<your-dokploy-domain>/oauth/callback` as a redirect URI on the Azure AD app.
