#!/usr/bin/env node

declare const __VERSION__: string

process.env.TRANSPORT_TYPE ??= "stdio"

const args = process.argv.slice(2)

if (args.includes("--version") || args.includes("-v")) {
  console.log(__VERSION__)
  process.exit(0)
}

if (args.includes("--help") || args.includes("-h")) {
  console.log(`
MS 365 MCP Server v${__VERSION__}

Usage: ms-365-mcp-server [options]

Options:
  -v, --version        Show version number
  -h, --help           Show help

Environment Variables:
  MS365_AUTH_MODE       Auth mode: interactive (default), certificate, client-secret, client-token
  MS365_TENANT_ID      Azure AD tenant ID (default: common)
  MS365_CLIENT_ID      Azure AD application (client) ID
  MS365_CLIENT_SECRET  Client secret (for client-secret mode)
  MS365_CERT_PATH      Certificate path (for certificate mode)
  MS365_CERT_PASSWORD  Certificate password (optional, for certificate mode)
  MS365_ACCESS_TOKEN   Initial access token (for client-token mode)
  MS365_GRAPH_VERSION  Graph API version: v1.0 or beta (default: v1.0)
  TRANSPORT_TYPE       Transport type: stdio (default) or httpStream
  PORT                 HTTP server port (default: 3000)
  HOST                 HTTP server host (default: 127.0.0.1)

For more information, visit: https://github.com/jordanburke/ms-365-mcp-server
`)
  process.exit(0)
}

async function main() {
  await import("./index.js")
}

main().catch(console.error)
