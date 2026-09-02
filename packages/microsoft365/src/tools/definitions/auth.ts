// Auth tool definitions.

import { z } from "zod"

import { getAuthStatusTool, listAccountsTool, setAccessTokenTool, switchAccountTool } from ".."
import type { ToolDefinition } from "../tool-definitions"
import { unwrapResult } from "./shared"

export const authTools: ReadonlyArray<ToolDefinition> = [
  {
    name: "get_auth_status",
    description: "Get current authentication status, mode, and scopes",
    parameters: z.object({}),
    execute: async () => unwrapResult(await getAuthStatusTool()),
    domain: "auth",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "list_accounts",
    description: "List all registered accounts and show which is the default",
    parameters: z.object({}),
    execute: async () => unwrapResult(await listAccountsTool()),
    domain: "auth",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "switch_account",
    description: "Switch the default account for subsequent tool calls",
    parameters: z.object({
      account_id: z.string().describe("Account ID to set as default"),
    }),
    execute: async (params) => unwrapResult(await switchAccountTool(params)),
    domain: "auth",
    readOnly: false,
  },
  {
    name: "set_access_token",
    description: "Set or update the access token (client-token auth mode only)",
    parameters: z.object({
      access_token: z.string().describe("The access token for Microsoft Graph"),
      expires_on: z.string().optional().describe("Token expiration time in ISO format"),
    }),
    // eslint-disable-next-line @typescript-eslint/require-await -- FastMCP requires async execute
    execute: async (params) => unwrapResult(setAccessTokenTool(params)),
    domain: "auth",
    readOnly: false,
  },
]
