import dotenv from "dotenv"
import type { UserError } from "fastmcp"
import { FastMCP } from "fastmcp"
import type { Either } from "functype/either"
import { z } from "zod"

import { initializeAuth } from "./auth"
import { initializeGraphClient } from "./client/graph-client"
import {
  createContact,
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
  getUser,
  graphQuery,
  listChannelMessages,
  listChannels,
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
  listTeams,
  listTodoLists,
  listTodoTasks,
  listUsers,
  replyToMessage,
  searchContacts,
  searchFiles,
  searchMessages,
  sendChannelMessage,
  sendMessage,
  setAccessTokenTool,
  updateEvent,
  updatePlannerTask,
  updateTodoTask,
} from "./tools"
import type { AuthConfig } from "./types"

dotenv.config()

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

const server = new FastMCP({
  name: "ms-365-mcp-server",
  version: VERSION,
  instructions: `A Microsoft 365 MCP server for managing email, calendar, contacts, files, Teams, Planner, OneNote, To Do, users, and groups via Microsoft Graph API.

Available capabilities:
- Mail: List, read, send, reply, and search email messages
- Calendar: List, view, create, update, and delete events
- Contacts: List, view, create, and search contacts
- Files: List, view, search, and download OneDrive files; create folders
- Teams: List teams, channels, and messages; send channel messages
- Users & Groups: View profiles, list users, groups, and group members
- Planner: List plans and tasks; create and update tasks
- OneNote: List notebooks, sections, pages; read page content
- To Do: List task lists and tasks; create and update tasks
- Graph Query: Execute arbitrary Microsoft Graph API queries`,
})

// === Auth Tools ===
server.addTool({
  name: "get_auth_status",
  description: "Get current authentication status, mode, and scopes",
  parameters: z.object({}),
  execute: async () => unwrapResult(await getAuthStatusTool()),
})

server.addTool({
  name: "set_access_token",
  description: "Set or update the access token (client-token auth mode only)",
  parameters: z.object({
    access_token: z.string().describe("The access token for Microsoft Graph"),
    expires_on: z.string().optional().describe("Token expiration time in ISO format"),
  }),
  // eslint-disable-next-line @typescript-eslint/require-await -- FastMCP requires async execute
  execute: async (params) => unwrapResult(setAccessTokenTool(params)),
})

// === Mail Tools ===
server.addTool({
  name: "list_messages",
  description: "List email messages from your inbox",
  parameters: z.object({
    top: z.number().optional().describe("Number of messages to return (default: 25)"),
    filter: z.string().optional().describe("OData filter expression"),
  }),
  execute: async (params) => unwrapResult(await listMessages(params)),
})

server.addTool({
  name: "get_message",
  description: "Get a specific email message with full body content",
  parameters: z.object({
    message_id: z.string().describe("The message ID"),
  }),
  execute: async (params) => unwrapResult(await getMessage(params)),
})

server.addTool({
  name: "send_message",
  description: "Send a new email message",
  parameters: z.object({
    to: z.string().describe("Recipient email address"),
    subject: z.string().describe("Email subject"),
    body: z.string().describe("Email body content"),
    content_type: z.string().optional().describe("Body content type: Text or HTML (default: Text)"),
  }),
  execute: async (params) => unwrapResult(await sendMessage(params)),
})

server.addTool({
  name: "reply_to_message",
  description: "Reply to an email message",
  parameters: z.object({
    message_id: z.string().describe("The message ID to reply to"),
    comment: z.string().describe("Reply content"),
  }),
  execute: async (params) => unwrapResult(await replyToMessage(params)),
})

server.addTool({
  name: "search_messages",
  description: "Search email messages",
  parameters: z.object({
    query: z.string().describe("Search query string"),
    top: z.number().optional().describe("Number of results to return (default: 25)"),
  }),
  execute: async (params) => unwrapResult(await searchMessages(params)),
})

// === Calendar Tools ===
server.addTool({
  name: "list_events",
  description: "List calendar events",
  parameters: z.object({
    top: z.number().optional().describe("Number of events to return (default: 25)"),
    filter: z.string().optional().describe("OData filter expression"),
  }),
  execute: async (params) => unwrapResult(await listEvents(params)),
})

server.addTool({
  name: "get_event",
  description: "Get detailed information about a calendar event",
  parameters: z.object({
    event_id: z.string().describe("The event ID"),
  }),
  execute: async (params) => unwrapResult(await getEvent(params)),
})

server.addTool({
  name: "create_event",
  description: "Create a new calendar event",
  parameters: z.object({
    subject: z.string().describe("Event subject/title"),
    start: z.string().describe("Start date/time (ISO format)"),
    end: z.string().describe("End date/time (ISO format)"),
    time_zone: z.string().optional().describe("Time zone (default: UTC)"),
    location: z.string().optional().describe("Event location"),
    body: z.string().optional().describe("Event description"),
    attendees: z.string().optional().describe("Comma-separated email addresses of attendees"),
  }),
  execute: async (params) => unwrapResult(await createEvent(params)),
})

server.addTool({
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
  }),
  execute: async (params) => unwrapResult(await updateEvent(params)),
})

server.addTool({
  name: "delete_event",
  description: "Delete a calendar event",
  parameters: z.object({
    event_id: z.string().describe("The event ID to delete"),
  }),
  execute: async (params) => unwrapResult(await deleteEvent(params)),
})

// === Contacts Tools ===
server.addTool({
  name: "list_contacts",
  description: "List contacts",
  parameters: z.object({
    top: z.number().optional().describe("Number of contacts to return (default: 25)"),
    filter: z.string().optional().describe("OData filter expression"),
  }),
  execute: async (params) => unwrapResult(await listContacts(params)),
})

server.addTool({
  name: "get_contact",
  description: "Get detailed contact information",
  parameters: z.object({
    contact_id: z.string().describe("The contact ID"),
  }),
  execute: async (params) => unwrapResult(await getContact(params)),
})

server.addTool({
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
})

server.addTool({
  name: "search_contacts",
  description: "Search contacts by name or email",
  parameters: z.object({
    query: z.string().describe("Search query"),
    top: z.number().optional().describe("Number of results (default: 25)"),
  }),
  execute: async (params) => unwrapResult(await searchContacts(params)),
})

// === Files Tools ===
server.addTool({
  name: "list_drive_items",
  description: "List files and folders in OneDrive",
  parameters: z.object({
    folder_id: z.string().optional().describe("Folder ID (omit for root)"),
  }),
  execute: async (params) => unwrapResult(await listDriveItems(params)),
})

server.addTool({
  name: "get_drive_item",
  description: "Get file or folder metadata",
  parameters: z.object({
    item_id: z.string().describe("Drive item ID"),
  }),
  execute: async (params) => unwrapResult(await getDriveItem(params)),
})

server.addTool({
  name: "search_files",
  description: "Search OneDrive/SharePoint files",
  parameters: z.object({
    query: z.string().describe("Search query"),
  }),
  execute: async (params) => unwrapResult(await searchFiles(params)),
})

server.addTool({
  name: "download_file",
  description: "Get file metadata and download URL",
  parameters: z.object({
    item_id: z.string().describe("Drive item ID"),
  }),
  execute: async (params) => unwrapResult(await downloadFile(params)),
})

server.addTool({
  name: "create_folder",
  description: "Create a new folder in OneDrive",
  parameters: z.object({
    parent_id: z.string().describe("Parent folder ID"),
    name: z.string().describe("Folder name"),
  }),
  execute: async (params) => unwrapResult(await createFolder(params)),
})

// === Teams Tools ===
server.addTool({
  name: "list_teams",
  description: "List teams you are a member of",
  parameters: z.object({}),
  execute: async () => unwrapResult(await listTeams()),
})

server.addTool({
  name: "list_channels",
  description: "List channels in a team",
  parameters: z.object({
    team_id: z.string().describe("Team ID"),
  }),
  execute: async (params) => unwrapResult(await listChannels(params)),
})

server.addTool({
  name: "list_channel_messages",
  description: "List recent messages in a channel",
  parameters: z.object({
    team_id: z.string().describe("Team ID"),
    channel_id: z.string().describe("Channel ID"),
    top: z.number().optional().describe("Number of messages (default: 25)"),
  }),
  execute: async (params) => unwrapResult(await listChannelMessages(params)),
})

server.addTool({
  name: "send_channel_message",
  description: "Send a message to a Teams channel",
  parameters: z.object({
    team_id: z.string().describe("Team ID"),
    channel_id: z.string().describe("Channel ID"),
    content: z.string().describe("Message content"),
  }),
  execute: async (params) => unwrapResult(await sendChannelMessage(params)),
})

// === Users & Groups Tools ===
server.addTool({
  name: "get_me",
  description: "Get the authenticated user's profile",
  parameters: z.object({}),
  execute: async () => unwrapResult(await getMe()),
})

server.addTool({
  name: "list_users",
  description: "List users in the organization",
  parameters: z.object({
    top: z.number().optional().describe("Number of users (default: 25)"),
    filter: z.string().optional().describe("OData filter expression"),
  }),
  execute: async (params) => unwrapResult(await listUsers(params)),
})

server.addTool({
  name: "get_user",
  description: "Get a specific user's profile",
  parameters: z.object({
    user_id: z.string().describe("User ID or UPN"),
  }),
  execute: async (params) => unwrapResult(await getUser(params)),
})

server.addTool({
  name: "list_groups",
  description: "List groups in the organization",
  parameters: z.object({
    top: z.number().optional().describe("Number of groups (default: 25)"),
    filter: z.string().optional().describe("OData filter expression"),
  }),
  execute: async (params) => unwrapResult(await listGroups(params)),
})

server.addTool({
  name: "get_group",
  description: "Get detailed group information",
  parameters: z.object({
    group_id: z.string().describe("Group ID"),
  }),
  execute: async (params) => unwrapResult(await getGroup(params)),
})

server.addTool({
  name: "list_group_members",
  description: "List members of a group",
  parameters: z.object({
    group_id: z.string().describe("Group ID"),
  }),
  execute: async (params) => unwrapResult(await listGroupMembers(params)),
})

// === Planner Tools ===
server.addTool({
  name: "list_plans",
  description: "List Planner plans",
  parameters: z.object({}),
  execute: async () => unwrapResult(await listPlans()),
})

server.addTool({
  name: "list_planner_tasks",
  description: "List tasks in a Planner plan",
  parameters: z.object({
    plan_id: z.string().describe("Plan ID"),
  }),
  execute: async (params) => unwrapResult(await listPlannerTasks(params)),
})

server.addTool({
  name: "get_planner_task",
  description: "Get detailed Planner task information",
  parameters: z.object({
    task_id: z.string().describe("Task ID"),
  }),
  execute: async (params) => unwrapResult(await getPlannerTask(params)),
})

server.addTool({
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
})

server.addTool({
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
})

// === OneNote Tools ===
server.addTool({
  name: "list_notebooks",
  description: "List OneNote notebooks",
  parameters: z.object({}),
  execute: async () => unwrapResult(await listNotebooks()),
})

server.addTool({
  name: "list_sections",
  description: "List sections in a OneNote notebook",
  parameters: z.object({
    notebook_id: z.string().describe("Notebook ID"),
  }),
  execute: async (params) => unwrapResult(await listSections(params)),
})

server.addTool({
  name: "list_pages",
  description: "List pages in a OneNote section",
  parameters: z.object({
    section_id: z.string().describe("Section ID"),
  }),
  execute: async (params) => unwrapResult(await listPages(params)),
})

server.addTool({
  name: "get_page_content",
  description: "Get OneNote page content as HTML",
  parameters: z.object({
    page_id: z.string().describe("Page ID"),
  }),
  execute: async (params) => unwrapResult(await getPageContent(params)),
})

// === To Do Tools ===
server.addTool({
  name: "list_todo_lists",
  description: "List Microsoft To Do task lists",
  parameters: z.object({}),
  execute: async () => unwrapResult(await listTodoLists()),
})

server.addTool({
  name: "list_todo_tasks",
  description: "List tasks in a To Do list",
  parameters: z.object({
    list_id: z.string().describe("To Do list ID"),
  }),
  execute: async (params) => unwrapResult(await listTodoTasks(params)),
})

server.addTool({
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
})

server.addTool({
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
})

// === Graph Query (Escape Hatch) ===
server.addTool({
  name: "graph_query",
  description: "Execute an arbitrary Microsoft Graph API query. Use this for operations not covered by other tools.",
  parameters: z.object({
    method: z.string().describe("HTTP method: GET, POST, PUT, PATCH, or DELETE"),
    path: z.string().describe("Graph API path (e.g., /me/memberOf)"),
    body: z.string().optional().describe("JSON request body as a string"),
    version: z.string().optional().describe("API version: v1.0 or beta (default: v1.0)"),
  }),
  execute: async (params) => unwrapResult(await graphQuery(params)),
})

// === Server Startup ===
const main = async () => {
  await setupAuth()

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

main().catch((error) => {
  console.error("[Fatal]", error)
  process.exit(1)
})
