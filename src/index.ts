import dotenv from "dotenv"
import type { UserError } from "fastmcp"
import { FastMCP } from "fastmcp"
// AzureSession shape: { accessToken: string; scopes: string[]; refreshToken?: string; upn?: string }
type OAuthSessionContext = { accessToken?: string }
import type { Either } from "functype/either"
import { z } from "zod"

import { initializeAuth } from "./auth"
import { createAzureAuthProvider } from "./auth/oauth-provider"
import { withToken } from "./auth/token-context"
import { initializeGraphClient } from "./client/graph-client"
import {
  createContact,
  createDraft,
  createEvent,
  createFolder,
  createPlannerTask,
  createTodoTask,
  deleteEvent,
  downloadFile,
  getAuthStatusTool,
  getContact,
  getDriveItem,
  getEvent,
  getGroup,
  getMe,
  getMessage,
  getPageContent,
  getPlannerTask,
  getSite,
  getUser,
  graphQuery,
  listAccountsTool,
  listChannelMessages,
  listChannels,
  listChatMessages,
  listChats,
  listContacts,
  listDriveItems,
  listEvents,
  listGroupMembers,
  listGroups,
  listMessages,
  listNotebooks,
  listPages,
  listPlannerTasks,
  listPlans,
  listSections,
  listSiteDrives,
  listSiteItems,
  listSites,
  listTeams,
  listTodoLists,
  listTodoTasks,
  listUsers,
  replyToMessage,
  searchContacts,
  searchFiles,
  searchMessages,
  searchSiteFiles,
  sendChannelMessage,
  sendChatMessage,
  sendDraft,
  sendMessage,
  setAccessTokenTool,
  switchAccountTool,
  updateEvent,
  updatePlannerTask,
  updateTodoTask,
  uploadFile,
} from "./tools"
import type { ToolDefinition } from "./tools/tool-definitions"
import { filterTools, type ToolFilterConfig } from "./tools/tool-registry"
import type { AuthConfig } from "./types"
import { auditToolCall, auditToolError, auditToolResult } from "./utils/audit"

dotenv.config({ quiet: true })

declare const __VERSION__: string
const VERSION = (typeof __VERSION__ !== "undefined" ? __VERSION__ : "0.0.0-dev") as `${number}.${number}.${number}`

const resolveAuthConfig = (): AuthConfig => {
  const mode = process.env.MS365_AUTH_MODE ?? "interactive"
  const tenantId = process.env.MS365_TENANT_ID ?? "common"
  const clientId = process.env.MS365_CLIENT_ID ?? ""

  switch (mode) {
    case "certificate":
      return {
        mode: "certificate",
        tenantId,
        clientId,
        certPath: process.env.MS365_CERT_PATH ?? "",
        certPassword: process.env.MS365_CERT_PASSWORD,
      }
    case "client-secret":
      return {
        mode: "client-secret",
        tenantId,
        clientId,
        clientSecret: process.env.MS365_CLIENT_SECRET ?? "",
      }
    case "client-token":
      return {
        mode: "client-token",
        accessToken: process.env.MS365_ACCESS_TOKEN,
      }
    case "oauth-proxy":
      return {
        mode: "oauth-proxy",
        tenantId,
        clientId,
        clientSecret: process.env.MS365_CLIENT_SECRET ?? "",
        baseUrl: process.env.MS365_OAUTH_BASE_URL ?? "http://localhost:3000",
      }
    default:
      return {
        mode: "interactive",
        tenantId,
        clientId,
        redirectUri: process.env.MS365_REDIRECT_URI,
      }
  }
}

const setupAuth = async () => {
  const config = resolveAuthConfig()
  const result = await initializeAuth(config)

  if (result.isLeft()) {
    const error = result.value as { message: string }
    if (config.mode === "client-token" && !config.accessToken) {
      console.error("[Setup] Client token mode: use set_access_token tool to provide a token")
    } else {
      console.error(`[Error] Authentication failed: ${error.message}`)
      process.exit(1)
    }
  } else {
    console.error(`[Setup] Authentication initialized (${config.mode} mode)`)
  }

  initializeGraphClient()
  console.error("[Setup] Graph client initialized")
}

/* eslint-disable functype/prefer-either -- boundary: converts Either to FastMCP's throw-based error signaling */
const unwrapResult = <T>(result: Either<UserError, T>): T =>
  result.fold(
    (e) => {
      throw e
    },
    (v) => v,
  )
/* eslint-enable functype/prefer-either */

const resolveFilterConfig = (): ToolFilterConfig => ({
  presets: process.env.MS365_PRESETS?.split(",").map((s) => s.trim()),
  enabledPattern: process.env.MS365_ENABLED_TOOLS,
  readOnly: process.env.MS365_READ_ONLY === "true",
  orgMode: process.env.MS365_ORG_MODE === "true",
})

const FETCH_ALL_PAGES_PARAM = z.boolean().optional().describe("Fetch all pages of results (max 50 pages)")

const toolDefinitions: ReadonlyArray<ToolDefinition> = [
  // === Auth Tools ===
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

  // === Mail Tools ===
  {
    name: "list_messages",
    description: "List email messages from your inbox",
    parameters: z.object({
      top: z.number().optional().describe("Number of messages to return (default: 25)"),
      filter: z.string().optional().describe("OData filter expression"),
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listMessages(params)),
    domain: "mail",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "get_message",
    description: "Get a specific email message with full body content",
    parameters: z.object({
      message_id: z.string().describe("The message ID"),
    }),
    execute: async (params) => unwrapResult(await getMessage(params)),
    domain: "mail",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "send_message",
    description: "Send a new email message",
    parameters: z.object({
      to: z.string().describe("Recipient email address"),
      subject: z.string().describe("Email subject"),
      body: z.string().describe("Email body content"),
      content_type: z.string().optional().describe("Body content type: Text or HTML (default: Text)"),
    }),
    execute: async (params) => unwrapResult(await sendMessage(params)),
    domain: "mail",
    readOnly: false,
    annotations: { destructiveHint: true },
  },
  {
    name: "reply_to_message",
    description: "Reply to an email message",
    parameters: z.object({
      message_id: z.string().describe("The message ID to reply to"),
      comment: z.string().describe("Reply content"),
    }),
    execute: async (params) => unwrapResult(await replyToMessage(params)),
    domain: "mail",
    readOnly: false,
    annotations: { destructiveHint: true },
  },
  {
    name: "search_messages",
    description: "Search email messages",
    parameters: z.object({
      query: z.string().describe("Search query string"),
      top: z.number().optional().describe("Number of results to return (default: 25)"),
    }),
    execute: async (params) => unwrapResult(await searchMessages(params)),
    domain: "mail",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "create_draft",
    description: "Create a new email draft in the Drafts folder",
    parameters: z.object({
      to: z.string().describe("Recipient email address"),
      subject: z.string().describe("Email subject"),
      body: z.string().describe("Email body content"),
      content_type: z.string().optional().describe("Body content type: Text or HTML (default: Text)"),
      cc: z.string().optional().describe("CC recipients (comma-separated email addresses)"),
      bcc: z.string().optional().describe("BCC recipients (comma-separated email addresses)"),
    }),
    execute: async (params) => unwrapResult(await createDraft(params)),
    domain: "mail",
    readOnly: false,
  },
  {
    name: "send_draft",
    description: "Send an existing email draft",
    parameters: z.object({
      message_id: z.string().describe("The draft message ID to send"),
    }),
    execute: async (params) => unwrapResult(await sendDraft(params)),
    domain: "mail",
    readOnly: false,
    annotations: { destructiveHint: true },
  },

  // === Calendar Tools ===
  {
    name: "list_events",
    description: "List calendar events",
    parameters: z.object({
      top: z.number().optional().describe("Number of events to return (default: 25)"),
      filter: z.string().optional().describe("OData filter expression"),
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listEvents(params)),
    domain: "calendar",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "get_event",
    description: "Get detailed information about a calendar event",
    parameters: z.object({
      event_id: z.string().describe("The event ID"),
    }),
    execute: async (params) => unwrapResult(await getEvent(params)),
    domain: "calendar",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "create_event",
    description: "Create a new calendar event",
    parameters: z.object({
      subject: z.string().describe("Event subject/title"),
      start: z.string().describe("Start date/time (ISO format)"),
      end: z.string().describe("End date/time (ISO format)"),
      time_zone: z.string().optional().describe("Time zone (default: UTC)"),
      location: z.string().optional().describe("Event location"),
      body: z.string().optional().describe("Event description"),
      content_type: z.string().optional().describe("Body content type: Text or HTML (default: Text)"),
      attendees: z.string().optional().describe("Comma-separated email addresses of attendees"),
      is_draft: z
        .boolean()
        .optional()
        .describe("Create as draft without sending invites to attendees (default: false)"),
    }),
    execute: async (params) => unwrapResult(await createEvent(params)),
    domain: "calendar",
    readOnly: false,
  },
  {
    name: "update_event",
    description: "Update an existing calendar event",
    parameters: z.object({
      event_id: z.string().describe("The event ID to update"),
      subject: z.string().optional().describe("New subject"),
      start: z.string().optional().describe("New start date/time (ISO format)"),
      end: z.string().optional().describe("New end date/time (ISO format)"),
      time_zone: z.string().optional().describe("Time zone (default: UTC)"),
      location: z.string().optional().describe("New location"),
      body: z.string().optional().describe("New description"),
      content_type: z.string().optional().describe("Body content type: Text or HTML (default: Text)"),
    }),
    execute: async (params) => unwrapResult(await updateEvent(params)),
    domain: "calendar",
    readOnly: false,
  },
  {
    name: "delete_event",
    description: "Delete a calendar event",
    parameters: z.object({
      event_id: z.string().describe("The event ID to delete"),
    }),
    execute: async (params) => unwrapResult(await deleteEvent(params)),
    domain: "calendar",
    readOnly: false,
    annotations: { destructiveHint: true },
  },

  // === Contacts Tools ===
  {
    name: "list_contacts",
    description: "List contacts",
    parameters: z.object({
      top: z.number().optional().describe("Number of contacts to return (default: 25)"),
      filter: z.string().optional().describe("OData filter expression"),
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listContacts(params)),
    domain: "contacts",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "get_contact",
    description: "Get detailed contact information",
    parameters: z.object({
      contact_id: z.string().describe("The contact ID"),
    }),
    execute: async (params) => unwrapResult(await getContact(params)),
    domain: "contacts",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "create_contact",
    description: "Create a new contact",
    parameters: z.object({
      given_name: z.string().describe("First name"),
      surname: z.string().optional().describe("Last name"),
      email: z.string().optional().describe("Email address"),
      mobile_phone: z.string().optional().describe("Mobile phone number"),
      company_name: z.string().optional().describe("Company name"),
      job_title: z.string().optional().describe("Job title"),
    }),
    execute: async (params) => unwrapResult(await createContact(params)),
    domain: "contacts",
    readOnly: false,
  },
  {
    name: "search_contacts",
    description: "Search contacts by name or email",
    parameters: z.object({
      query: z.string().describe("Search query"),
      top: z.number().optional().describe("Number of results (default: 25)"),
    }),
    execute: async (params) => unwrapResult(await searchContacts(params)),
    domain: "contacts",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },

  // === Files Tools ===
  {
    name: "list_drive_items",
    description: "List files and folders in OneDrive",
    parameters: z.object({
      folder_id: z.string().optional().describe("Folder ID (omit for root)"),
      folder_path: z.string().optional().describe("Folder path (e.g., 'Documents' or 'Documents/Subfolder')"),
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listDriveItems(params)),
    domain: "files",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "get_drive_item",
    description: "Get file or folder metadata",
    parameters: z.object({
      item_id: z.string().describe("Drive item ID"),
    }),
    execute: async (params) => unwrapResult(await getDriveItem(params)),
    domain: "files",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "search_files",
    description: "Search OneDrive/SharePoint files",
    parameters: z.object({
      query: z.string().describe("Search query"),
    }),
    execute: async (params) => unwrapResult(await searchFiles(params)),
    domain: "files",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "download_file",
    description:
      "Download a file. Returns content inline for text files under 100KB, otherwise returns metadata and download URL.",
    parameters: z.object({
      item_id: z.string().describe("Drive item ID"),
    }),
    execute: async (params) => unwrapResult(await downloadFile(params)),
    domain: "files",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "create_folder",
    description: "Create a new folder in OneDrive",
    parameters: z.object({
      parent_id: z.string().describe("Parent folder ID"),
      name: z.string().describe("Folder name"),
    }),
    execute: async (params) => unwrapResult(await createFolder(params)),
    domain: "files",
    readOnly: false,
  },
  {
    name: "upload_file",
    description:
      "Upload a file to OneDrive. For text files, pass content directly. For binary files, pass base64-encoded content with content_type 'application/octet-stream'. Max ~4MB via this tool.",
    parameters: z.object({
      path: z
        .string()
        .describe("Destination path in colon-path format (e.g., /me/drive/root:/Documents/file.txt:/content)"),
      content: z.string().describe("File content (text or base64-encoded for binary)"),
      content_type: z
        .string()
        .optional()
        .describe("MIME type (default: text/plain). Use application/octet-stream for base64 binary."),
    }),
    execute: async (params) => unwrapResult(await uploadFile(params)),
    domain: "files",
    readOnly: false,
  },

  // === SharePoint Tools ===
  {
    name: "list_sites",
    description: "List SharePoint sites. Without a query, returns followed sites. With a query, searches all sites.",
    parameters: z.object({
      query: z.string().optional().describe("Search query to find sites (omit to list followed sites)"),
    }),
    execute: async (params) => unwrapResult(await listSites(params)),
    domain: "files",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "get_site",
    description: "Get SharePoint site details",
    parameters: z.object({
      site_id: z.string().describe("Site ID (e.g., 'contoso.sharepoint.com,siteId,webId')"),
    }),
    execute: async (params) => unwrapResult(await getSite(params)),
    domain: "files",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "list_site_drives",
    description: "List document libraries (drives) in a SharePoint site",
    parameters: z.object({
      site_id: z.string().describe("Site ID"),
    }),
    execute: async (params) => unwrapResult(await listSiteDrives(params)),
    domain: "files",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "list_site_items",
    description: "List files and folders in a SharePoint site's document library",
    parameters: z.object({
      site_id: z.string().describe("Site ID"),
      drive_id: z.string().optional().describe("Drive ID (omit for default document library)"),
      folder_id: z.string().optional().describe("Folder ID (omit for root)"),
      folder_path: z.string().optional().describe("Folder path (e.g., 'General/Reports')"),
    }),
    execute: async (params) => unwrapResult(await listSiteItems(params)),
    domain: "files",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "search_site_files",
    description: "Search files within a SharePoint site",
    parameters: z.object({
      site_id: z.string().describe("Site ID"),
      query: z.string().describe("Search query"),
      drive_id: z.string().optional().describe("Drive ID (omit to search default document library)"),
    }),
    execute: async (params) => unwrapResult(await searchSiteFiles(params)),
    domain: "files",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },

  // === Chat Tools ===
  {
    name: "list_chats",
    description:
      "List your Teams chats (1:1, group, and meeting chats). Note: the self-chat (notes to self) is not listed here — use chat_id '48:notes' to send to it directly.",
    parameters: z.object({
      top: z.number().optional().describe("Number of chats to return (default: 25)"),
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listChats(params)),
    domain: "chats" as const,
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "list_chat_messages",
    description: "List messages in a Teams chat",
    parameters: z.object({
      chat_id: z.string().describe("Chat ID"),
      top: z.number().optional().describe("Number of messages (default: 25)"),
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listChatMessages(params)),
    domain: "chats" as const,
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "send_chat_message",
    description:
      "Send a message in a Teams chat. Use chat_id '48:notes' to send to the user's self-chat (notes to self).",
    parameters: z.object({
      chat_id: z.string().describe("Chat ID. Use '48:notes' for the user's self-chat."),
      content: z.string().describe("Message content"),
    }),
    execute: async (params) => unwrapResult(await sendChatMessage(params)),
    domain: "chats" as const,
    readOnly: false,
    annotations: { destructiveHint: true },
  },

  // === Teams Tools ===
  {
    name: "list_teams",
    description: "List teams you are a member of",
    parameters: z.object({
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listTeams(params)),
    domain: "teams",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "list_channels",
    description: "List channels in a team",
    parameters: z.object({
      team_id: z.string().describe("Team ID"),
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listChannels(params)),
    domain: "teams",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "list_channel_messages",
    description: "List recent messages in a channel",
    parameters: z.object({
      team_id: z.string().describe("Team ID"),
      channel_id: z.string().describe("Channel ID"),
      top: z.number().optional().describe("Number of messages (default: 25)"),
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listChannelMessages(params)),
    domain: "teams",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "send_channel_message",
    description: "Send a message to a Teams channel",
    parameters: z.object({
      team_id: z.string().describe("Team ID"),
      channel_id: z.string().describe("Channel ID"),
      content: z.string().describe("Message content"),
    }),
    execute: async (params) => unwrapResult(await sendChannelMessage(params)),
    domain: "teams",
    readOnly: false,
    annotations: { destructiveHint: true },
  },

  // === Users & Groups Tools ===
  {
    name: "get_me",
    description: "Get the authenticated user's profile",
    parameters: z.object({}),
    execute: async () => unwrapResult(await getMe()),
    domain: "users",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "list_users",
    description: "List users in the organization",
    parameters: z.object({
      top: z.number().optional().describe("Number of users (default: 25)"),
      filter: z.string().optional().describe("OData filter expression"),
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listUsers(params)),
    domain: "users",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "get_user",
    description: "Get a specific user's profile",
    parameters: z.object({
      user_id: z.string().describe("User ID or UPN"),
    }),
    execute: async (params) => unwrapResult(await getUser(params)),
    domain: "users",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "list_groups",
    description: "List groups in the organization",
    parameters: z.object({
      top: z.number().optional().describe("Number of groups (default: 25)"),
      filter: z.string().optional().describe("OData filter expression"),
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listGroups(params)),
    domain: "groups",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "get_group",
    description: "Get detailed group information",
    parameters: z.object({
      group_id: z.string().describe("Group ID"),
    }),
    execute: async (params) => unwrapResult(await getGroup(params)),
    domain: "groups",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "list_group_members",
    description: "List members of a group",
    parameters: z.object({
      group_id: z.string().describe("Group ID"),
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listGroupMembers(params)),
    domain: "groups",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },

  // === Planner Tools ===
  {
    name: "list_plans",
    description: "List Planner plans",
    parameters: z.object({
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listPlans(params)),
    domain: "planner",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "list_planner_tasks",
    description: "List tasks in a Planner plan",
    parameters: z.object({
      plan_id: z.string().describe("Plan ID"),
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listPlannerTasks(params)),
    domain: "planner",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "get_planner_task",
    description: "Get detailed Planner task information",
    parameters: z.object({
      task_id: z.string().describe("Task ID"),
    }),
    execute: async (params) => unwrapResult(await getPlannerTask(params)),
    domain: "planner",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "create_planner_task",
    description: "Create a new Planner task",
    parameters: z.object({
      plan_id: z.string().describe("Plan ID"),
      title: z.string().describe("Task title"),
      bucket_id: z.string().optional().describe("Bucket ID"),
      due_date: z.string().optional().describe("Due date (ISO format)"),
      assignments: z.string().optional().describe("Comma-separated user IDs to assign"),
    }),
    execute: async (params) => unwrapResult(await createPlannerTask(params)),
    domain: "planner",
    readOnly: false,
  },
  {
    name: "update_planner_task",
    description: "Update a Planner task",
    parameters: z.object({
      task_id: z.string().describe("Task ID"),
      etag: z.string().describe("Task ETag for concurrency control"),
      title: z.string().optional().describe("New title"),
      percent_complete: z.number().optional().describe("Completion percentage (0-100)"),
      due_date: z.string().optional().describe("New due date (ISO format)"),
      priority: z.number().optional().describe("Priority (0-10)"),
    }),
    execute: async (params) => unwrapResult(await updatePlannerTask(params)),
    domain: "planner",
    readOnly: false,
  },

  // === OneNote Tools ===
  {
    name: "list_notebooks",
    description: "List OneNote notebooks",
    parameters: z.object({
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listNotebooks(params)),
    domain: "onenote",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "list_sections",
    description: "List sections in a OneNote notebook",
    parameters: z.object({
      notebook_id: z.string().describe("Notebook ID"),
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listSections(params)),
    domain: "onenote",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "list_pages",
    description: "List pages in a OneNote section",
    parameters: z.object({
      section_id: z.string().describe("Section ID"),
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listPages(params)),
    domain: "onenote",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "get_page_content",
    description: "Get OneNote page content as HTML",
    parameters: z.object({
      page_id: z.string().describe("Page ID"),
    }),
    execute: async (params) => unwrapResult(await getPageContent(params)),
    domain: "onenote",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },

  // === To Do Tools ===
  {
    name: "list_todo_lists",
    description: "List Microsoft To Do task lists",
    parameters: z.object({
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listTodoLists(params)),
    domain: "todo",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "list_todo_tasks",
    description: "List tasks in a To Do list",
    parameters: z.object({
      list_id: z.string().describe("To Do list ID"),
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listTodoTasks(params)),
    domain: "todo",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "create_todo_task",
    description: "Create a new To Do task",
    parameters: z.object({
      list_id: z.string().describe("To Do list ID"),
      title: z.string().describe("Task title"),
      body: z.string().optional().describe("Task body/notes"),
      due_date: z.string().optional().describe("Due date (ISO format)"),
      importance: z.string().optional().describe("Importance: low, normal, or high"),
    }),
    execute: async (params) => unwrapResult(await createTodoTask(params)),
    domain: "todo",
    readOnly: false,
  },
  {
    name: "update_todo_task",
    description: "Update a To Do task",
    parameters: z.object({
      list_id: z.string().describe("To Do list ID"),
      task_id: z.string().describe("Task ID"),
      title: z.string().optional().describe("New title"),
      status: z.string().optional().describe("Status: notStarted, inProgress, completed, waitingOnOthers, deferred"),
      due_date: z.string().optional().describe("New due date (ISO format)"),
      importance: z.string().optional().describe("Importance: low, normal, or high"),
      body: z.string().optional().describe("New body/notes"),
    }),
    execute: async (params) => unwrapResult(await updateTodoTask(params)),
    domain: "todo",
    readOnly: false,
  },

  // === Graph Query (Escape Hatch) ===
  {
    name: "graph_query",
    description: "Execute an arbitrary Microsoft Graph API query. Use this for operations not covered by other tools.",
    parameters: z.object({
      method: z.string().describe("HTTP method: GET, POST, PUT, PATCH, or DELETE"),
      path: z.string().describe("Graph API path (e.g., /me/memberOf)"),
      body: z.string().optional().describe("JSON request body as a string"),
      version: z.string().optional().describe("API version: v1.0 or beta (default: v1.0)"),
    }),
    execute: async (params) => unwrapResult(await graphQuery(params)),
    domain: "query",
    readOnly: false,
    annotations: { destructiveHint: true, openWorldHint: true },
  },
]

type ExecuteContext = { session?: OAuthSessionContext }

const wrapExecute = (tool: ToolDefinition, oauthMode: boolean): never => {
  const baseFn = tool.execute as (p: Record<string, unknown>) => Promise<string>

  // Layer 1: OAuth token injection (wraps the base function)
  const withOAuth = oauthMode
    ? (params: Record<string, unknown>, context: ExecuteContext) =>
        withToken(context.session?.accessToken, () => baseFn(params))
    : (params: Record<string, unknown>) => baseFn(params)

  // Layer 2: Audit logging (wraps the OAuth-aware function)
  const withAudit = async (params: Record<string, unknown>, context: ExecuteContext) => {
    auditToolCall(tool.name, params)
    const start = Date.now()

    try {
      const result = await withOAuth(params, context)
      auditToolResult(tool.name, true, Date.now() - start)
      return result
    } catch (error) {
      auditToolError(tool.name, error instanceof Error ? error.message : String(error))
      auditToolResult(tool.name, false, Date.now() - start)
      throw error
    }
  }

  return withAudit as never
}

const registerTools = (server: FastMCP, allowedTools: Set<string>, oauthMode: boolean) => {
  let registered = 0
  let skipped = 0

  for (const tool of toolDefinitions) {
    if (!allowedTools.has(tool.name)) {
      skipped++
      continue
    }

    server.addTool({
      name: tool.name,
      description: tool.description,
      parameters: tool.parameters,
      execute: wrapExecute(tool, oauthMode),
      annotations: tool.annotations,
    })
    registered++
  }

  console.error(`[Setup] Tools registered: ${registered}, skipped: ${skipped}`)
}

const buildInstructions = (allowedTools: Set<string>): string => {
  const domains = new Set(toolDefinitions.filter((t) => allowedTools.has(t.name)).map((t) => t.domain))
  const domainDescriptions: Record<string, string> = {
    auth: "Authentication: Check auth status and manage tokens",
    mail: "Mail: List, read, send, reply, search, and draft email messages",
    calendar: "Calendar: List, view, create, update, and delete events",
    contacts: "Contacts: List, view, create, and search contacts",
    files: "Files: List, view, search, and download OneDrive files; create folders",
    chats: "Chats: List Teams chats and messages; send chat messages",
    teams: "Teams: List teams, channels, and messages; send channel messages",
    users: "Users: View profiles and list users",
    groups: "Groups: List groups and group members",
    planner: "Planner: List plans and tasks; create and update tasks",
    onenote: "OneNote: List notebooks, sections, pages; read page content",
    todo: "To Do: List task lists and tasks; create and update tasks",
    query: "Graph Query: Execute arbitrary Microsoft Graph API queries",
  }

  const capabilities = [...domains]
    .map((d) => domainDescriptions[d])
    .filter(Boolean)
    .map((desc) => `- ${desc}`)
    .join("\n")

  return `A Microsoft 365 MCP server via Microsoft Graph API.\n\nAvailable capabilities:\n${capabilities}`
}

// === Server Startup ===
const main = async () => {
  const authConfig = resolveAuthConfig()
  const oauthMode = authConfig.mode === "oauth-proxy"

  const filterConfig = resolveFilterConfig()
  const allowedTools = filterTools(filterConfig)

  if (oauthMode) {
    // OAuth proxy mode: FastMCP handles auth via AzureProvider
    const provider = createAzureAuthProvider({
      baseUrl: (authConfig as { baseUrl: string }).baseUrl,
      clientId: (authConfig as { clientId: string }).clientId,
      clientSecret: (authConfig as { clientSecret: string }).clientSecret,
      tenantId: (authConfig as { tenantId: string }).tenantId,
    })

    const server = new FastMCP({
      name: "microsoft365-mcp-server",
      version: VERSION,
      instructions: buildInstructions(allowedTools),
      auth: provider,
    } as never)

    // Initialize graph client without credential-based auth (tokens come from session)
    initializeGraphClient()

    registerTools(server, allowedTools, true)

    const port = parseInt(process.env.PORT ?? "3000", 10)
    await server.start({ transportType: "httpStream", httpStream: { port } })
    console.error(`[Server] MS 365 MCP Server v${VERSION} (OAuth proxy) running on port ${port}`)
  } else {
    // Standard mode: credential-based auth
    await setupAuth()

    const server = new FastMCP({
      name: "microsoft365-mcp-server",
      version: VERSION,
      instructions: buildInstructions(allowedTools),
    })

    registerTools(server, allowedTools, false)

    const transportType = process.env.TRANSPORT_TYPE ?? "stdio"

    if (transportType === "httpStream") {
      const port = parseInt(process.env.PORT ?? "3000", 10)
      const host = process.env.HOST ?? "127.0.0.1"
      await server.start({ transportType: "httpStream", httpStream: { port } })
      console.error(`[Server] MS 365 MCP Server v${VERSION} running on ${host}:${port}`)
    } else {
      await server.start({ transportType: "stdio" })
      console.error(`[Server] MS 365 MCP Server v${VERSION} running on stdio`)
    }
  }
}

main().catch((error) => {
  console.error("[Fatal]", error)
  process.exit(1)
})
