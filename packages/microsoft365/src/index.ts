import type { UserError } from "fastmcp"
import { FastMCP } from "fastmcp"
// AzureSession shape: { accessToken: string; scopes: string[]; refreshToken?: string; upn?: string }
type OAuthSessionContext = { accessToken?: string }
import {
  attachFile,
  decodeBase64Upload,
  describeFetchError,
  filenameFromPath,
  loadEnvFileIfPresent,
  MAX_ATTACHMENT_SIZE,
  MAX_UPLOAD_SIZE,
  requireDraftMessage,
  resolveUploadContentType,
  sessionUpload,
  SIMPLE_UPLOAD_LIMIT,
  simpleUpload,
} from "@sapientsai/ms-graph-core"
import { type Either, Left, Right } from "functype/either"
import { z } from "zod"

import { getAccessToken, initializeAuth } from "./auth"
import { createAzureAuthProvider } from "./auth/oauth-provider"
import { GRAPH_API_BASE } from "./auth/scopes"
import { withToken } from "./auth/token-context"
import { getGraphClient, initializeGraphClient } from "./client/graph-client"
import { defaultVersion } from "./client/graph-version"
import { resolveMailbox } from "./mail/mailbox"
import {
  batchMoveMessages,
  copyOnenotePage,
  createContact,
  createDraft,
  createEvent,
  createFolder,
  createForwardDraft,
  createOnenoteNotebook,
  createOnenotePage,
  createOnenoteSection,
  createPlannerBucket,
  createPlannerTask,
  createReplyAllDraft,
  createReplyDraft,
  createTodoList,
  createTodoTask,
  deleteContact,
  deleteDraft,
  deleteEvent,
  deleteOnenotePage,
  deleteTodoTask,
  downloadFile,
  exportContacts,
  findMeetingAvailability,
  getAttachmentUploadConfig,
  getAuthStatusTool,
  getContact,
  getDriveItem,
  getEvent,
  getGroup,
  getMe,
  getMeetingTranscript,
  getMessage,
  getOnenotePageContent,
  getPlannerTask,
  getSite,
  getUploadConfig,
  getUser,
  graphQuery,
  listAccountsTool,
  listAttachments,
  listCalendarView,
  listChannelMessages,
  listChannels,
  listChatMessages,
  listChats,
  listContacts,
  listDriveItems,
  listEvents,
  listGroupMembers,
  listGroups,
  listMailFolders,
  listMeetingTranscripts,
  listMessages,
  listOnenoteNotebooks,
  listOnenotePages,
  listOnenoteSections,
  listPlannerBuckets,
  listPlannerTasks,
  listPlans,
  listSiteDrives,
  listSiteItems,
  listSites,
  listTeams,
  listTodoLists,
  listTodoTasks,
  listUsers,
  moveMessage,
  moveMessagesMatching,
  readDocument,
  respondToEvent,
  saveAttachment,
  scanMessages,
  searchContacts,
  searchFiles,
  searchMessages,
  searchSiteFiles,
  sendChannelMessage,
  sendChatMessage,
  sendDraft,
  sendForward,
  sendMessage,
  sendReply,
  sendReplyAll,
  setAccessTokenTool,
  setTodoTaskState,
  summarizeSenders,
  switchAccountTool,
  updateContact,
  updateDraft,
  updateEvent,
  updateOnenotePageContent,
  updatePlannerTask,
  updatePlannerTaskDetails,
  updateTodoTask,
  uploadFile,
  uploadFileFromPath,
} from "./tools"
import { SET_TODO_TASK_STATE_PARAMS, TODO_EXTENSION_PARAM } from "./tools/todo-schemas"
import type { ToolDefinition } from "./tools/tool-definitions"
import {
  annotationsFor,
  DOMAIN_DESCRIPTIONS,
  filterTools,
  requireDraftEnabled,
  type ToolFilterConfig,
} from "./tools/tool-registry"
import type { AuthConfig } from "./types"
import { resolveUploadAccessToken } from "./upload/upload-auth"
import { auditToolCall, auditToolError, auditToolResult } from "./utils/audit"
import { DAYS_OF_WEEK, RECURRENCE_PATTERN_TYPES, RECURRENCE_RANGE_TYPES, WEEK_INDEXES } from "./utils/recurrence"

loadEnvFileIfPresent()

declare const __VERSION__: string
const VERSION = (typeof __VERSION__ !== "undefined" ? __VERSION__ : "0.0.0-dev") as `${number}.${number}.${number}`

// The Dockerfiles have always set GIT_HASH as a container env, and nothing read it — so a
// running container could name its version but not its build. Two deploys of one tag are
// indistinguishable by version alone, and "which build is actually running" is precisely the
// question asked when a deploy looks wrong. Answering it from the image digest works but costs
// a round trip through the registry.
const gitHash = process.env.GIT_HASH
const BUILD = gitHash !== undefined && gitHash.length > 0 ? ` (${gitHash.slice(0, 7)})` : ""

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
        useDeviceCode: process.env.MS365_USE_DEVICE_CODE === "true",
        disableSignIn: process.env.MS365_DISABLE_SIGN_IN === "true",
      }
  }
}

const setupAuth = async () => {
  const config = resolveAuthConfig()
  const result = await initializeAuth(config)

  result.fold(
    (error) => {
      if (config.mode === "client-token" && !config.accessToken) {
        console.error("[Setup] Client token mode: use set_access_token tool to provide a token")
      } else {
        console.error(`[Error] Authentication failed: ${(error as { message: string }).message}`)
        process.exit(1)
      }
    },
    () => console.error(`[Setup] Authentication initialized (${config.mode} mode)`),
  )

  initializeGraphClient({ getAccessToken })
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

const resolveFilterConfig = (transport: "stdio" | "httpStream"): ToolFilterConfig => ({
  presets: process.env.MS365_PRESETS?.split(",").map((s) => s.trim()),
  enabledPattern: process.env.MS365_ENABLED_TOOLS,
  readOnly: process.env.MS365_READ_ONLY === "true",
  orgMode: process.env.MS365_ORG_MODE === "true",
  requireDraft: requireDraftEnabled(),
  transport,
})

const FETCH_ALL_PAGES_PARAM = z.boolean().optional().describe("Fetch all pages of results (max 50 pages)")

// Mail tools accept an optional mailbox so one server can serve a household: the
// signed-in user's own mail by default, a delegated or shared mailbox when named.
// Which addresses are permitted is deployment configuration (MS365_ALLOWED_MAILBOXES),
// not something the caller can widen, so the description points at the error rather
// than listing addresses that would go stale.
const MAILBOX_PARAM = z
  .string()
  .optional()
  .describe(
    "Email address of another mailbox to act on (delegated or shared), e.g. 'someone@example.com'. " +
      "Omit for your own mailbox; required when this server signs in as an application. " +
      "Only addresses this server is configured to allow will be accepted.",
  )

// The flat shape here mirrors Graph's patternedRecurrence without making the caller
// nest two objects. Which fields are required depends on the pattern — the builder
// enforces that and names what is missing, so the description stays short.
const RECURRENCE_PARAM = z
  .object({
    pattern: z.enum(RECURRENCE_PATTERN_TYPES).describe("How the task repeats"),
    interval: z
      .number()
      .int()
      .positive()
      .optional()
      .describe("Units between occurrences, in the pattern's unit. Default 1. Quarterly is absoluteMonthly with 3."),
    days_of_week: z
      .array(z.enum(DAYS_OF_WEEK))
      .optional()
      .describe("Required for weekly, relativeMonthly and relativeYearly"),
    day_of_month: z
      .number()
      .int()
      .min(1)
      .max(31)
      .optional()
      .describe("Required for absoluteMonthly and absoluteYearly"),
    month: z.number().int().min(1).max(12).optional().describe("Required for absoluteYearly and relativeYearly"),
    index: z
      .enum(WEEK_INDEXES)
      .optional()
      .describe("Which week of the month, for the relative patterns. Default first"),
    first_day_of_week: z.enum(DAYS_OF_WEEK).optional().describe("For weekly patterns. Default monday"),
    range_type: z.enum(RECURRENCE_RANGE_TYPES).optional().describe("noEnd (default), endDate, or numbered"),
    start_date: z
      .string()
      .optional()
      .describe("YYYY-MM-DD. Defaults to the task's due date. On update, must match due_date if both are given"),
    end_date: z.string().optional().describe("YYYY-MM-DD. Required when range_type is endDate"),
    number_of_occurrences: z.number().int().positive().optional().describe("Required when range_type is numbered"),
    recurrence_time_zone: z.string().optional().describe("Time zone for the range, e.g. 'W. Australia Standard Time'"),
  })
  .optional()

// A To Do category is the displayName of one of the person's Outlook categories.
const TODO_CATEGORIES_PARAM = z
  .array(z.string())
  .optional()
  .describe("Category names, matching the person's Outlook categories")

const toolDefinitions: ReadonlyArray<ToolDefinition> = [
  // === Auth Tools ===
  {
    name: "get_auth_status",
    description: "Get current authentication status, mode, and scopes",
    parameters: z.object({}),
    execute: async () => unwrapResult(await getAuthStatusTool()),
    domain: "auth",
    readOnly: true,
  },
  {
    name: "list_accounts",
    description: "List all registered accounts and show which is the default",
    parameters: z.object({}),
    execute: async () => unwrapResult(await listAccountsTool()),
    domain: "auth",
    readOnly: true,
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
    description:
      "List email messages, newest first. Covers every folder (Inbox, Sent Items, Archive, ...) unless " +
      "folder is set. Each line shows the sender's name and address, the received time, the Graph " +
      "message ID and the internetMessageId.",
    parameters: z.object({
      top: z.number().optional().describe("Number of messages to return (default: 25)"),
      filter: z
        .string()
        .optional()
        .describe(
          'OData filter expression, e.g. "importance eq \'high\'" or "isRead eq false". Results are sorted ' +
            "newest first, and Graph rejects a filter that does not lead with receivedDateTime, so the tool " +
            "puts an always-true receivedDateTime condition first when yours does not start with one. To " +
            "restrict to one folder, use folder instead of filtering on parentFolderId.",
        ),
      folder: z
        .string()
        .optional()
        .describe(
          "Only this folder: a well-known name (inbox, sentitems, drafts, deleteditems, archive, junkemail), " +
            "a top-level folder's display name, or a folder ID from list_mail_folders (needed for a subfolder). " +
            "Default: every folder.",
        ),
      include_preview: z
        .boolean()
        .optional()
        .describe("Add each message's body preview (Graph's bodyPreview, up to 255 characters) on an indented line"),
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
      mailbox: MAILBOX_PARAM,
    }),
    execute: async (params) => unwrapResult(await listMessages(params)),
    domain: "mail",
    readOnly: true,
  },
  {
    name: "scan_messages",
    description:
      "Scan message headers compactly as ref|received|from|subject|flags. Scope with folder and narrow with filter or search. Pass refs to get_message, list_attachments, move_message or batch_move_messages. Both filter and search scans page using Graph nextLink cursors: when a result says INCOMPLETE, call scan_messages with the returned cursor alone until no cursor remains. A short or empty page can still have a continuation. Search returns at most 1,000 results, so use filter and exhaust its cursors when mailbox coverage matters. A header is not a message: open messages and attachments the subject cannot rule out.",
    parameters: z.object({
      folder: z
        .string()
        .optional()
        .describe(
          "Folder to scan: well-known name (archive, inbox, sentitems), display name, or ID. Default: all mail",
        ),
      filter: z.string().optional().describe("OData filter, e.g. receivedDateTime ge 2026-01-01T00:00:00Z"),
      search: z.string().optional().describe("Full-text search term. Cannot be combined with date ordering"),
      top: z.number().int().min(1).max(1000).optional().describe("Rows per page (default 100, max 1000)"),
      cursor: z
        .string()
        .optional()
        .describe(
          "Continuation cursor from a prior scan. Pass alone, plus mailbox if the scan named one; valid for this session",
        ),
      mailbox: MAILBOX_PARAM,
    }),
    execute: async (params) => unwrapResult(await scanMessages(params)),
    domain: "mail",
    readOnly: true,
  },
  {
    name: "get_message",
    description:
      "Get a specific email message with full body content. Pass body_format:'text' for marketing or newsletter mail — Graph converts server-side, avoiding tens of thousands of characters of HTML and CSS.",
    parameters: z.object({
      message_id: z.string().describe("The message ID, or a short ref returned by scan_messages"),
      body_format: z
        .enum(["text", "html"])
        .optional()
        .describe("Body format to request. 'text' strips HTML/CSS server-side. Default: the message's own format"),
      mailbox: MAILBOX_PARAM,
    }),
    execute: async (params) => unwrapResult(await getMessage(params)),
    domain: "mail",
    readOnly: true,
  },
  {
    name: "list_mail_folders",
    description:
      "List mail folders with item and unread counts, for resolving move destinations. Top-level folders only — a folder reporting subfolders has children this does not list.",
    parameters: z.object({
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
      mailbox: MAILBOX_PARAM,
    }),
    execute: async (params) => unwrapResult(await listMailFolders(params)),
    domain: "mail",
    readOnly: true,
  },
  {
    name: "list_attachments",
    description:
      "List a message's attachments with name, content type and size. Returns a read_document path for file attachments so their text can be extracted (PDF, Office, etc); cloud links and embedded Outlook items are marked as not readable that way.",
    parameters: z.object({
      message_id: z.string().describe("The message ID, or a short ref from scan_messages, whose attachments to list"),
      mailbox: MAILBOX_PARAM,
    }),
    execute: async (params) => unwrapResult(await listAttachments(params)),
    domain: "mail",
    readOnly: true,
  },
  {
    name: "move_message",
    description:
      "Move a message to another mail folder. Destination accepts a well-known name (archive, deleteditems, inbox, junkemail), a top-level folder display name, or a folder ID. A well-known name always wins over a custom folder of the same name, and the confirmation says which was used. Subfolders cannot be resolved by name — pass their ID. Moving to deleteditems is recoverable; use list_mail_folders to see what exists.",
    parameters: z.object({
      message_id: z.string().describe("The message ID, or a short ref from scan_messages, to move"),
      destination: z
        .string()
        .describe("Destination folder: well-known name (e.g. archive), display name, or folder ID"),
      mailbox: MAILBOX_PARAM,
    }),
    execute: async (params) => unwrapResult(await moveMessage(params)),
    domain: "mail",
    readOnly: false,
  },
  {
    name: "batch_move_messages",
    description:
      "Move several messages to the same folder in one call. Resolves the destination once and returns a single summary instead of one result per message. Reports any failures individually.",
    parameters: z.object({
      message_ids: z.array(z.string()).describe("Message IDs or scan_messages refs to move (max 50)"),
      destination: z
        .string()
        .describe("Destination folder: well-known name (e.g. archive), display name, or folder ID"),
      mailbox: MAILBOX_PARAM,
    }),
    execute: async (params) => unwrapResult(await batchMoveMessages(params)),
    domain: "mail",
    readOnly: false,
  },
  {
    name: "save_attachment",
    description:
      "Save a mail attachment to a local file and return its path. Use for anything read_document cannot " +
      "extract text from — scanned PDFs, photographed documents, images — and for handing a file to another " +
      "tool. Read the saved file directly: PDFs and images are viewable without text extraction. Omit " +
      "attachment_id when the message has exactly one attachment; otherwise get IDs from list_attachments.",
    parameters: z.object({
      message_id: z.string().describe("The message ID, or a short ref returned by scan_messages"),
      attachment_id: z
        .string()
        .optional()
        .describe("Which attachment to save. Optional when the message has exactly one."),
      out_dir: z
        .string()
        .optional()
        .describe(
          "Directory to write into (default: the system temp directory). An existing file is never replaced; a numbered name is used instead.",
        ),
      mailbox: MAILBOX_PARAM,
    }),
    execute: async (params) => unwrapResult(await saveAttachment(params)),
    domain: "mail",
    readOnly: false,
  },
  {
    name: "summarize_senders",
    description:
      "Count a folder's messages by sender, server-side, without listing them. Returns one row per sender " +
      "(count, unread, first and last date, address, name, latest subject), most frequent first — the " +
      "starting point for cleaning a large inbox: newsletters, promotions and notifications show up as a " +
      "few hundred senders repeating.\n\n" +
      "Reads every matching message in the folder (paged internally, up to ~50,000) so the counts are " +
      "complete; only the summary is returned. Narrow with filter (OData, e.g. a receivedDateTime range) " +
      "when only part of a folder matters. Use group_by: domain to see families of senders, then " +
      "group_by: address to get the addresses move_messages_matching needs.",
    parameters: z.object({
      folder: z
        .string()
        .optional()
        .describe("Folder to summarise: well-known name (inbox, archive), display name, or ID. Default: inbox"),
      filter: z
        .string()
        .optional()
        .describe('Optional OData filter to narrow the set, e.g. "receivedDateTime lt 2025-01-01T00:00:00Z"'),
      group_by: z
        .enum(["address", "domain"])
        .optional()
        .describe("Group by full sender address (default) or by the domain after the @"),
      top: z.number().optional().describe("How many senders to show (default 100, max 1000)"),
      mailbox: MAILBOX_PARAM,
    }),
    execute: async (params) => unwrapResult(await summarizeSenders(params)),
    domain: "mail",
    readOnly: true,
  },
  {
    name: "move_messages_matching",
    description:
      "Move every message in a folder that matches a sender list and/or an OData filter — the bulk step of " +
      "an inbox clean-up. Pass senders (addresses from summarize_senders) and/or filter; a whole folder " +
      "with no condition is refused. Destination deleteditems is how to bulk-delete: it is reversible from " +
      "Deleted Items, and a hard delete is deliberately not offered.\n\n" +
      "SAFE BY DEFAULT: dry_run is true unless set to false. A dry run fetches the matching set and reports " +
      "the count, date range, sender breakdown and newest subjects, moving nothing — read it before " +
      "re-running with dry_run: false. A live run refuses to move more than limit messages (default 1000) so " +
      "a mis-scoped filter cannot empty a folder; raise limit explicitly when the dry-run count is intended.\n\n" +
      "Moves go through Graph JSON batching (20 per round-trip, throttling retried), so thousands of " +
      "messages take seconds, not minutes, and no message IDs pass through the caller.",
    parameters: z.object({
      folder: z.string().describe("Folder to sweep: well-known name (inbox), display name, or ID. Required."),
      destination: z
        .string()
        .describe("Where matches go: well-known name (deleteditems, archive, junkemail), display name, or ID"),
      senders: z
        .array(z.string())
        .optional()
        .describe("Sender addresses to match, exact and case-insensitive (max 20 per call)"),
      filter: z
        .string()
        .optional()
        .describe('OData filter, e.g. "receivedDateTime lt 2025-01-01T00:00:00Z". Combined with senders using AND.'),
      dry_run: z
        .boolean()
        .optional()
        .describe("Default true: report what would move and move nothing. Pass false to move."),
      limit: z
        .number()
        .optional()
        .describe("Refuse a live run that would move more than this many (default 1000, max 20000)"),
      mailbox: MAILBOX_PARAM,
    }),
    execute: async (params) => unwrapResult(await moveMessagesMatching(params)),
    domain: "mail",
    readOnly: false,
  },
  {
    name: "send_message",
    description: "Send a new email message",
    parameters: z.object({
      to: z.string().describe("Recipient email address(es), comma-separated for multiple"),
      subject: z.string().describe("Email subject"),
      body: z.string().describe("Email body content"),
      content_type: z.string().optional().describe("Body content type: Text or HTML (default: Text)"),
      mailbox: MAILBOX_PARAM,
    }),
    execute: async (params) => unwrapResult(await sendMessage(params)),
    domain: "mail",
    readOnly: false,
  },
  {
    name: "send_reply",
    description:
      "Reply to the sender of an email and send immediately. Threads into the conversation and quotes the original. Use create_reply_draft to review before sending.",
    parameters: z.object({
      message_id: z.string().describe("The message ID to reply to"),
      comment: z.string().describe("Reply content (added above the quoted original)"),
      mailbox: MAILBOX_PARAM,
    }),
    execute: async (params) => unwrapResult(await sendReply(params)),
    domain: "mail",
    readOnly: false,
  },
  {
    name: "send_reply_all",
    description:
      "Reply to all recipients of an email and send immediately. Threads into the conversation and quotes the original. Use create_reply_all_draft to review before sending.",
    parameters: z.object({
      message_id: z.string().describe("The message ID to reply to"),
      comment: z.string().describe("Reply content (added above the quoted original)"),
      mailbox: MAILBOX_PARAM,
    }),
    execute: async (params) => unwrapResult(await sendReplyAll(params)),
    domain: "mail",
    readOnly: false,
  },
  {
    name: "send_forward",
    description:
      "Forward an email to new recipients and send immediately. Quotes the original. Use create_forward_draft to review before sending.",
    parameters: z.object({
      message_id: z.string().describe("The message ID to forward"),
      to: z.string().describe("Recipient email address(es), comma-separated for multiple"),
      comment: z.string().optional().describe("Optional note added above the quoted original"),
      mailbox: MAILBOX_PARAM,
    }),
    execute: async (params) => unwrapResult(await sendForward(params)),
    domain: "mail",
    readOnly: false,
  },
  {
    name: "create_reply_draft",
    description:
      "Create a reply draft (to the sender) in the Drafts folder. Threads into the conversation with the original quoted underneath. Review, then send with send_draft.",
    parameters: z.object({
      message_id: z.string().describe("The message ID to reply to"),
      comment: z.string().describe("Reply content (added above the quoted original)"),
      mailbox: MAILBOX_PARAM,
    }),
    execute: async (params) => unwrapResult(await createReplyDraft(params)),
    domain: "mail",
    readOnly: false,
  },
  {
    name: "create_reply_all_draft",
    description:
      "Create a reply-all draft (to all recipients) in the Drafts folder. Threads into the conversation with the original quoted underneath. Review, then send with send_draft.",
    parameters: z.object({
      message_id: z.string().describe("The message ID to reply to"),
      comment: z.string().describe("Reply content (added above the quoted original)"),
      mailbox: MAILBOX_PARAM,
    }),
    execute: async (params) => unwrapResult(await createReplyAllDraft(params)),
    domain: "mail",
    readOnly: false,
  },
  {
    name: "create_forward_draft",
    description:
      "Create a forward draft in the Drafts folder with the original quoted underneath. Review, then send with send_draft.",
    parameters: z.object({
      message_id: z.string().describe("The message ID to forward"),
      to: z.string().describe("Recipient email address(es), comma-separated for multiple"),
      comment: z.string().optional().describe("Optional note added above the quoted original"),
      mailbox: MAILBOX_PARAM,
    }),
    execute: async (params) => unwrapResult(await createForwardDraft(params)),
    domain: "mail",
    readOnly: false,
  },
  {
    name: "search_messages",
    description: "Search email messages",
    parameters: z.object({
      query: z.string().describe("Search query string"),
      top: z.number().optional().describe("Number of results to return (default: 25)"),
      mailbox: MAILBOX_PARAM,
    }),
    execute: async (params) => unwrapResult(await searchMessages(params)),
    domain: "mail",
    readOnly: true,
  },
  {
    name: "create_draft",
    description: "Create a new email draft in the Drafts folder",
    parameters: z.object({
      to: z.string().describe("Recipient email address(es), comma-separated for multiple"),
      subject: z.string().describe("Email subject"),
      body: z.string().describe("Email body content"),
      content_type: z.string().optional().describe("Body content type: Text or HTML (default: Text)"),
      cc: z.string().optional().describe("CC recipients (comma-separated email addresses)"),
      bcc: z.string().optional().describe("BCC recipients (comma-separated email addresses)"),
      mailbox: MAILBOX_PARAM,
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
      mailbox: MAILBOX_PARAM,
    }),
    execute: async (params) => unwrapResult(await sendDraft(params)),
    domain: "mail",
    readOnly: false,
  },
  {
    name: "update_draft",
    description:
      "Edit an unsent draft: its subject, body or recipients. Only fields you pass change. Refuses any " +
      "message that is not a draft. Send it afterward with send_draft.",
    parameters: z.object({
      message_id: z.string().describe("The draft's message ID"),
      subject: z.string().optional().describe("New subject"),
      body: z.string().optional().describe("New body; replaces the whole body, including any quoted history"),
      content_type: z.string().optional().describe("Body content type: Text or HTML (default: Text)"),
      to: z
        .string()
        .optional()
        .describe("Comma-separated email addresses; replaces the current To list. Empty string clears it"),
      cc: z
        .string()
        .optional()
        .describe("Comma-separated email addresses; replaces the current CC list. Empty string clears it"),
      bcc: z
        .string()
        .optional()
        .describe("Comma-separated email addresses; replaces the current BCC list. Empty string clears it"),
      mailbox: MAILBOX_PARAM,
    }),
    execute: async (params) => unwrapResult(await updateDraft(params)),
    domain: "mail",
    readOnly: false,
  },
  {
    name: "delete_draft",
    description: "Delete an unsent draft. Refuses any message that is not a draft.",
    parameters: z.object({
      message_id: z.string().describe("The draft's message ID"),
      mailbox: MAILBOX_PARAM,
    }),
    execute: async (params) => unwrapResult(await deleteDraft(params)),
    domain: "mail",
    readOnly: false,
  },
  {
    name: "get_attachment_upload_config",
    description:
      "Get an authenticated upload URL + curl command for attaching a file (PDF, image, any binary) to an existing draft. " +
      "Pipe base64 file contents to the returned URL via POST; this server attaches the file itself (inline under 3 MB, " +
      "Outlook upload session up to 150 MB), so the bytes never pass through the conversation and no sandbox access to " +
      "Outlook is needed. Refuses any message that is not a draft. Create the draft first, then send it with send_draft.",
    parameters: z.object({
      message_id: z.string().describe("The draft's message ID, or a ref from scan_messages"),
      name: z.string().describe("File name the attachment will show, e.g. expense-report.pdf"),
      localFile: z
        .string()
        .optional()
        .describe("Local file path to include in the curl example. If omitted, a placeholder is used."),
      contentType: z.string().optional().describe("MIME type override. Auto-detected from the name if omitted."),
      mailbox: MAILBOX_PARAM,
    }),
    execute: async (params) => unwrapResult(await getAttachmentUploadConfig(params)),
    domain: "mail",
    readOnly: false,
  },

  // === Calendar Tools ===
  {
    name: "list_events",
    description:
      "List calendar event resources (/me/events). Returns series masters for recurring meetings, NOT individual occurrences. " +
      "For 'what's on my calendar between X and Y' use list_calendar_view instead — it expands recurrences into instances.",
    parameters: z.object({
      top: z.number().optional().describe("Number of events to return (default: 25)"),
      filter: z.string().optional().describe("OData filter expression"),
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listEvents(params)),
    domain: "calendar",
    readOnly: true,
  },
  {
    name: "list_calendar_view",
    description:
      "List event instances on the calendar between start_date_time and end_date_time. Expands recurring series into " +
      "individual occurrences (unlike list_events which returns series masters). Use this for 'what's on my calendar this week'.",
    parameters: z.object({
      start_date_time: z.string().describe("Window start (ISO 8601, e.g. 2026-05-22T00:00:00Z)"),
      end_date_time: z.string().describe("Window end (ISO 8601, e.g. 2026-05-29T00:00:00Z)"),
      top: z.number().optional().describe("Max events to return (default: 50)"),
    }),
    execute: async (params) => unwrapResult(await listCalendarView(params)),
    domain: "calendar",
    readOnly: true,
  },
  {
    name: "find_meeting_availability",
    description:
      "Find meeting time suggestions where all participants are free, ranked by confidence. Searches between " +
      "after_date_time and before_date_time for a slot of duration_minutes. Read-only: suggests times only — " +
      "use create_event to book one.",
    parameters: z.object({
      participants: z
        .array(z.string())
        .min(1)
        .describe("Attendee email addresses. The signed-in user is automatically the organizer."),
      after_date_time: z.string().describe("Search window start (ISO 8601, e.g. 2026-06-04T00:00:00Z)"),
      before_date_time: z.string().describe("Search window end (ISO 8601, e.g. 2026-06-06T00:00:00Z)"),
      duration_minutes: z
        .number()
        .int()
        .min(15)
        .max(480)
        .optional()
        .describe("Meeting length in minutes (default: 30)"),
      max_candidates: z.number().int().min(1).max(50).optional().describe("Max slots to return (default: 3)"),
      is_organizer_optional: z
        .boolean()
        .optional()
        .describe("Whether the organizer's attendance is optional (default: false)"),
    }),
    execute: async (params) => unwrapResult(await findMeetingAvailability(params)),
    domain: "calendar",
    readOnly: true,
  },
  {
    name: "get_event",
    description:
      "Get detailed information about a calendar event. The body comes back as plain text by default; a " +
      "Teams invite's text still includes its join link and any meeting ID and passcode.",
    parameters: z.object({
      event_id: z.string().describe("The event ID"),
      body_format: z
        .enum(["text", "html"])
        .optional()
        .describe("Body format: 'text' (default) has Graph strip the HTML server-side; 'html' returns the raw markup"),
    }),
    execute: async (params) => unwrapResult(await getEvent(params)),
    domain: "calendar",
    readOnly: true,
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
        .describe(
          "Save to your calendar without inviting attendees (any attendees param is ignored). " +
            "Add attendees and send the meeting later from Outlook. Default: false.",
        ),
      online_meeting: z
        .boolean()
        .optional()
        .describe("Add a Teams meeting to the event; joinUrl is returned in the event details (default: false)"),
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
      attendees: z
        .string()
        .optional()
        .describe("Comma-separated email addresses; replaces the current attendee list when provided"),
      online_meeting: z
        .boolean()
        .optional()
        .describe(
          "Add a Teams meeting to the event. One-way per Graph: once enabled it cannot be turned off or re-provisioned. Default: false.",
        ),
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
  },
  {
    name: "respond_to_event",
    description:
      "Accept, decline or tentatively accept a meeting invitation. By default the organizer gets the " +
      "response by email, with the comment if one is given.",
    parameters: z.object({
      event_id: z.string().describe("The event ID of the invitation"),
      response: z.enum(["accept", "decline", "tentative"]).describe("How to respond"),
      comment: z.string().optional().describe("Note to include in the response"),
      send_response: z
        .boolean()
        .optional()
        .describe("Send the response to the organizer (default: true). False updates only your calendar."),
    }),
    execute: async (params) => unwrapResult(await respondToEvent(params)),
    domain: "calendar",
    readOnly: false,
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
    name: "update_contact",
    description:
      "Update fields on an existing contact. Only the fields passed are changed. List fields " +
      "(email_addresses, business_phones, home_phones, categories) REPLACE the whole list — read the " +
      "contact first and pass the complete new list; an empty list clears it.",
    parameters: z.object({
      contact_id: z.string().describe("The contact ID"),
      given_name: z.string().optional().describe("First name"),
      surname: z.string().optional().describe("Last name"),
      nickname: z.string().optional().describe("Nickname"),
      company_name: z.string().optional().describe("Company name"),
      job_title: z.string().optional().describe("Job title"),
      mobile_phone: z.string().optional().describe("Mobile phone number"),
      email_addresses: z.array(z.string()).optional().describe("Complete replacement list of email addresses"),
      business_phones: z.array(z.string()).optional().describe("Complete replacement list of business phones"),
      home_phones: z.array(z.string()).optional().describe("Complete replacement list of home phones"),
      categories: z.array(z.string()).optional().describe("Complete replacement list of Outlook categories"),
      personal_notes: z.string().optional().describe("Replaces the contact's notes"),
      birthday: z
        .string()
        .optional()
        .describe(
          "Birthday as YYYY-MM-DD, or an ISO 8601 timestamp (e.g. copied from another contact's birthday as-is)",
        ),
    }),
    execute: async (params) => unwrapResult(await updateContact(params)),
    domain: "contacts",
    readOnly: false,
  },
  {
    name: "delete_contact",
    description:
      "Delete a contact. Intended for duplicates: first copy anything the duplicate holds that the kept " +
      "contact lacks (emails, phones, notes, birthday) onto the kept contact with update_contact, then delete. " +
      "Returns the deleted contact's details so they are not lost from the conversation.",
    parameters: z.object({
      contact_id: z.string().describe("The contact ID"),
    }),
    execute: async (params) => unwrapResult(await deleteContact(params)),
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
  },
  {
    name: "export_contacts",
    description:
      "Export every contact in the signed-in mailbox to a local JSON file and return its path. Covers the " +
      "default Contacts folder and every contact folder beneath it, with all fields (emails, phones, " +
      "addresses, categories, birthday, notes, folder path). Use this to reconcile or bulk-process " +
      "contacts; use list_contacts or search_contacts to look one up.",
    parameters: z.object({
      out_dir: z.string().optional().describe("Directory to write into (default: the system temp directory)"),
    }),
    execute: async (params) => unwrapResult(await exportContacts(params)),
    domain: "contacts",
    readOnly: false,
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
  },
  {
    name: "search_files",
    description:
      "Search OneDrive/SharePoint files. Search results show folders without item counts or sizes, which search does not report; use list_drive_items or get_drive_item for those.",
    parameters: z.object({
      query: z.string().describe("Search query"),
      top: z
        .number()
        .int()
        .positive()
        .optional()
        .describe("Maximum results to return (default: 25). Use fetch_all_pages when every match is needed."),
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await searchFiles(params)),
    domain: "files",
    readOnly: true,
  },
  {
    name: "download_file",
    description:
      "Get a file's metadata and download URL. Returns content inline for text files under 100KB. For a " +
      "SharePoint file, pass drive_id as well — without it the item is looked up in your own OneDrive and " +
      "will not be found. For readable text from PDF/DOCX/XLSX, use read_document instead; come here when " +
      "extraction fails (scanned PDFs, unsupported types, files over the extraction cap) or you need raw bytes.",
    parameters: z.object({
      item_id: z.string().describe("Drive item ID"),
      drive_id: z
        .string()
        .optional()
        .describe("Drive ID holding the item. Required for SharePoint; omit for your own OneDrive."),
    }),
    execute: async (params) => unwrapResult(await downloadFile(params)),
    domain: "files",
    readOnly: true,
  },
  {
    name: "read_document",
    description:
      "Download a file from SharePoint or OneDrive and return its readable text content. Supports DOCX, PDF, XLSX, " +
      "and text-based files. Use instead of download_file when you need document contents. Pair with " +
      "search_site_files (SharePoint) or search_files (OneDrive) to get IDs, then pass " +
      "/drives/{driveId}/items/{itemId}/content or /me/drive/items/{id}/content. Text extraction only, no OCR: " +
      "scanned PDFs return no text — for those, and for images, use save_attachment and read the file.",
    parameters: z.object({
      path: z.string().describe("Graph path to the file content endpoint, ending in /content"),
      api_version: z.enum(["v1.0", "beta"]).optional().describe("Graph API version"),
      format: z.string().optional().describe("Optional conversion format (e.g. 'pdf'), for supported types only"),
      max_chars: z
        .number()
        .int()
        .min(1000)
        .max(200000)
        .optional()
        .describe("Max characters to return (1000-200000, default 50000); content beyond is truncated"),
    }),
    execute: async (params) => unwrapResult(await readDocument(params)),
    domain: "rag",
    readOnly: true,
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
      "Upload TEXT content (plain text, markdown, CSV, JSON, HTML, XML) to OneDrive inline via this tool call. For binary files (docx, pdf, images, etc.), use get_upload_config (HTTP/SSE) or upload_file_from_path (stdio/local) — never base64-encode binary into this tool's content param. Max ~4 MB text.",
    parameters: z.object({
      path: z
        .string()
        .describe("Destination path in colon-path format (e.g., /me/drive/root:/Documents/file.txt:/content)"),
      content: z.string().describe("Text content to upload (UTF-8)"),
      content_type: z
        .string()
        .optional()
        .describe(
          "MIME type of the content, default text/plain. Must be a text type (text/*, application/json, application/xml, application/javascript, *+json, *+xml); binary types are rejected. It only gates this check: OneDrive decides the stored file's MIME type itself and may not keep this value.",
        ),
      conflict_behavior: z
        .enum(["rename", "replace", "fail"])
        .optional()
        .describe('Conflict behavior: "rename" (default), "replace" overwrites, "fail" returns 409 on collision'),
    }),
    execute: async (params) => unwrapResult(await uploadFile(params)),
    domain: "files",
    readOnly: false,
  },
  {
    name: "get_upload_config",
    description:
      "Get an authenticated upload URL + curl command for uploading files to OneDrive. Primary path for binary files or anything >1 MB. Pipe base64 file contents to the returned URL via POST; the server decodes and streams to Microsoft Graph (chunked session upload for >4 MB, up to 250 MB). Intermediate folders are auto-created. Note: response includes operational caveats in notes — read before executing the curl, especially on failure.",
    parameters: z.object({
      path: z
        .string()
        .describe(
          "Graph API destination path ending with :/content (e.g., /me/drive/root:/Documents/file.docx:/content)",
        ),
      localFile: z
        .string()
        .optional()
        .describe("Local file path to include in the curl example. If omitted, a placeholder is used."),
      contentType: z.string().optional().describe("MIME type override. Auto-detected from file extension if omitted."),
      conflictBehavior: z
        .enum(["rename", "replace", "fail"])
        .optional()
        .describe('Conflict behavior: "rename" (default), "replace" overwrites, "fail" returns an error'),
    }),
    execute: async (params) => unwrapResult(await getUploadConfig(params)),
    domain: "files",
    readOnly: false,
  },
  {
    name: "upload_file_from_path",
    description:
      "Upload a local file to OneDrive by reading it from disk on the server. The file must exist on this machine's filesystem. If you generated the file in a cloud container (e.g., claude.ai), first use Desktop Commander's write_file to save it to the user's local filesystem, then call this tool with that local path. Supports files up to 250 MB (chunked session upload above 4 MB). Intermediate folders are auto-created.",
    parameters: z.object({
      local_path: z.string().describe("Absolute path to the local file to upload"),
      path: z
        .string()
        .describe(
          "Graph API destination path ending with :/content (e.g., /me/drive/root:/Documents/file.docx:/content)",
        ),
      content_type: z.string().optional().describe("MIME type override. Auto-detected from file extension if omitted."),
      conflict_behavior: z
        .enum(["rename", "replace", "fail"])
        .optional()
        .describe('Conflict behavior: "rename" (default), "replace" overwrites, "fail" returns an error'),
    }),
    execute: async (params) => unwrapResult(await uploadFileFromPath(params)),
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
  },
  {
    name: "search_site_files",
    description:
      "Search files within a SharePoint site. Search results show folders without item counts or sizes, which search does not report; use list_drive_items or get_drive_item for those.",
    parameters: z.object({
      site_id: z.string().describe("Site ID"),
      query: z.string().describe("Search query"),
      drive_id: z.string().optional().describe("Drive ID (omit to search default document library)"),
      top: z
        .number()
        .int()
        .positive()
        .optional()
        .describe("Maximum results to return (default: 25). Use fetch_all_pages when every match is needed."),
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await searchSiteFiles(params)),
    domain: "files",
    readOnly: true,
  },

  // === Chat Tools ===
  {
    name: "list_chats",
    description:
      "List your Teams chats (1:1, group, and meeting chats), most recent message first. Each line names " +
      "the chat by its topic or its members and shows when the last message was sent and by whom. " +
      "Note: the self-chat (notes to self) is not listed here — use chat_id '48:notes' to send to it directly.",
    parameters: z.object({
      since: z
        .string()
        .optional()
        .describe(
          "Only chats whose last message is after this ISO 8601 time with a time zone, e.g. " +
            "2026-10-01T00:00:00Z. Reads pages until it reaches an older chat (at most 20 pages) and returns " +
            "every match unless top is set. Chats with no messages are skipped.",
        ),
      top: z
        .number()
        .int()
        .positive()
        .max(50)
        .optional()
        .describe(
          "Number of chats to return (default: 25, max 50; with since, every match unless set). Graph may " +
            "return fewer per page when member names are included; use fetch_all_pages for more.",
        ),
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listChats(params)),
    domain: "chats" as const,
    readOnly: true,
  },
  {
    name: "list_chat_messages",
    description:
      "List messages in a Teams chat, newest first. Each message is a line " +
      "'- **<sender>** (<created>) [You] [App] [Urgent|High importance] [Mentions you] (ID: <id>)' " +
      "(flags only when they apply), followed by its text on a '  > ' line. System events (joins, " +
      "renames) and deleted messages are left out.",
    parameters: z.object({
      chat_id: z.string().describe("Chat ID"),
      top: z.number().int().positive().max(50).optional().describe("Number of messages (default: 25, max 50)"),
      since: z
        .string()
        .optional()
        .describe(
          "Only messages created or changed after this ISO 8601 time with a time zone, e.g. " +
            "2026-10-01T00:00:00Z. An old message edited after it is included; its line still shows when it " +
            "was created. Returns at most top messages (default 25); if more match, the output ends with " +
            "'Note: more messages match; raise top or move since later to see them.'",
        ),
      max_chars: z
        .number()
        .int()
        .positive()
        .optional()
        .describe("Cut each message's text at this length (default: 300)"),
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listChatMessages(params)),
    domain: "chats" as const,
    readOnly: true,
  },
  {
    name: "send_chat_message",
    description:
      "Send a Teams chat message, either in an existing chat (chat_id) or to people by email (to). With " +
      "to, one person reuses your one-on-one chat with them, or starts it; two or more start a new group " +
      "chat every call, so reuse the returned chat_id for follow-ups. Use chat_id '48:notes' to send to " +
      "the user's self-chat (notes to self).",
    parameters: z.object({
      chat_id: z
        .string()
        .optional()
        .describe("Chat ID of an existing chat. Use '48:notes' for the user's self-chat. Pass this or to."),
      to: z
        .string()
        .optional()
        .describe(
          "Comma-separated work email addresses or sign-in names of people in your organization (guests " +
            "included) to message. Pass this or chat_id.",
        ),
      topic: z.string().optional().describe("Title for a new group chat (two or more people in to)"),
      content: z.string().describe("Message content"),
    }),
    execute: async (params) => unwrapResult(await sendChatMessage(params)),
    domain: "chats" as const,
    readOnly: false,
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
  },
  {
    name: "send_channel_message",
    description: "Send a message to a Teams channel, or reply in a thread with reply_to_message_id",
    parameters: z.object({
      team_id: z.string().describe("Team ID"),
      channel_id: z.string().describe("Channel ID"),
      content: z.string().describe("Message content"),
      reply_to_message_id: z
        .string()
        .optional()
        .describe(
          "ID of a top-level channel message to reply to, from list_channel_messages. Omit to start a new thread.",
        ),
    }),
    execute: async (params) => unwrapResult(await sendChannelMessage(params)),
    domain: "teams",
    readOnly: false,
  },

  // === Meeting Transcript Tools ===
  //
  // Gated by admin-consent scopes that are deliberately not in the defaults, so on a deployment
  // that has not opted in via MS365_EXTRA_SCOPES these are visible but return a 403 explaining
  // what to grant. See the README's "Meeting transcripts" section.
  {
    name: "list_meeting_transcripts",
    description:
      "List the transcripts of a Teams meeting. Identify the meeting by meeting_id, or by " +
      "join_web_url taken from a calendar event's onlineMeeting.joinUrl. Returns transcript IDs " +
      "for get_meeting_transcript.",
    parameters: z.object({
      meeting_id: z.string().optional().describe("Online meeting ID. Preferred — needs no lookup."),
      join_web_url: z
        .string()
        .optional()
        .describe("Teams join URL (an event's onlineMeeting.joinUrl). Requires the OnlineMeetings.Read scope."),
    }),
    execute: async (params) => unwrapResult(await listMeetingTranscripts(params)),
    domain: "meetings",
    readOnly: true,
  },
  {
    name: "get_meeting_transcript",
    description:
      "Get the text of one Teams meeting transcript, as timestamped WebVTT utterances. Works for " +
      "attendees, not just the organizer. Falls back to a speaker-less transcript automatically " +
      "when the tenant has speaker attribution disabled.",
    parameters: z.object({
      transcript_id: z.string().describe("Transcript ID from list_meeting_transcripts"),
      meeting_id: z.string().optional().describe("Online meeting ID. Preferred — needs no lookup."),
      join_web_url: z
        .string()
        .optional()
        .describe("Teams join URL, as an alternative to meeting_id. Requires the OnlineMeetings.Read scope."),
      include_speaker_names: z
        .boolean()
        .optional()
        .describe("Ask for speaker-attributed text (default: true). Ignored if the tenant disallows it."),
      max_chars: z
        .number()
        .int()
        .positive()
        .optional()
        .describe("Return at most this many characters per call (default: 50000)"),
      offset: z
        .number()
        .int()
        .nonnegative()
        .optional()
        .describe(
          "Character position to start from (default: 0). When the output ends with a truncation marker, " +
            "call again with the offset it names to get the next part.",
        ),
    }),
    execute: async (params) => unwrapResult(await getMeetingTranscript(params)),
    domain: "meetings",
    readOnly: true,
  },

  // === Users & Groups Tools ===
  {
    name: "get_me",
    description: "Get the authenticated user's profile",
    parameters: z.object({}),
    execute: async () => unwrapResult(await getMe()),
    domain: "users",
    readOnly: true,
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
  },

  // === Planner Tools ===
  {
    name: "list_plans",
    description: "List all Planner plans visible to you, aggregated across your group memberships",
    parameters: z.object({}),
    execute: async () => unwrapResult(await listPlans()),
    domain: "planner",
    readOnly: true,
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
  },
  {
    name: "list_planner_buckets",
    description: "List the buckets (columns) in a Planner plan. Use a bucket ID with create_planner_task.",
    parameters: z.object({
      plan_id: z.string().describe("Plan ID"),
    }),
    execute: async (params) => unwrapResult(await listPlannerBuckets(params)),
    domain: "planner",
    readOnly: true,
  },
  {
    name: "create_planner_bucket",
    description: "Create a new bucket (column) in a Planner plan",
    parameters: z.object({
      plan_id: z.string().describe("Plan ID"),
      name: z.string().describe("Bucket name"),
    }),
    execute: async (params) => unwrapResult(await createPlannerBucket(params)),
    domain: "planner",
    readOnly: false,
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
    description:
      "Update a Planner task (title, percent_complete — 100 closes it — due date, priority). The ETag " +
      "is auto-fetched if omitted; pass one only to enforce optimistic concurrency.",
    parameters: z.object({
      task_id: z.string().describe("Task ID"),
      etag: z.string().optional().describe("Task ETag — omit to auto-fetch; pass to enforce concurrency"),
      title: z.string().optional().describe("New title"),
      percent_complete: z.number().optional().describe("Completion percentage (0-100; 100 marks complete)"),
      due_date: z.string().optional().describe("New due date (ISO format)"),
      priority: z.number().optional().describe("Priority (0-10)"),
    }),
    execute: async (params) => unwrapResult(await updatePlannerTask(params)),
    domain: "planner",
    readOnly: false,
  },
  {
    name: "update_planner_task_details",
    description:
      "Update a Planner task's details: description, checklist items, and references (links). The ETag " +
      "is fetched automatically. Add, update, or remove entries — updates merge (omitted fields keep " +
      "their current value); update/remove of a missing item is reported as skipped, not a silent no-op. " +
      "Checklist items are addressed by their GUID (from get_planner_task); references by their URL.",
    parameters: z.object({
      task_id: z.string().describe("Task ID"),
      description: z.string().optional().describe("Task description / notes"),
      preview_type: z
        .enum(["automatic", "noPreview", "checklist", "description", "reference"])
        .optional()
        .describe("What shows on the task card face"),
      add_checklist: z
        .array(z.object({ title: z.string(), isChecked: z.boolean().optional() }))
        .optional()
        .describe("Checklist items to add"),
      update_checklist: z
        .array(z.object({ id: z.string(), title: z.string().optional(), isChecked: z.boolean().optional() }))
        .optional()
        .describe("Checklist items to edit, addressed by GUID (e.g. toggle isChecked)"),
      remove_checklist: z.array(z.string()).optional().describe("Checklist item GUIDs to remove"),
      add_references: z
        .array(z.object({ url: z.string(), alias: z.string().optional() }))
        .optional()
        .describe("Reference links to add"),
      update_references: z
        .array(z.object({ url: z.string(), alias: z.string().optional() }))
        .optional()
        .describe("Reference links to edit, addressed by URL (e.g. rename the alias)"),
      remove_references: z.array(z.string()).optional().describe("Reference URLs to remove"),
    }),
    execute: async (params) => unwrapResult(await updatePlannerTaskDetails(params)),
    domain: "planner",
    readOnly: false,
  },

  // === OneNote Tools ===
  {
    name: "list_onenote_notebooks",
    description: "List OneNote notebooks",
    parameters: z.object({
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listOnenoteNotebooks(params)),
    domain: "onenote",
    readOnly: true,
  },
  {
    name: "list_onenote_sections",
    description: "List sections in a OneNote notebook",
    parameters: z.object({
      notebook_id: z.string().describe("Notebook ID"),
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listOnenoteSections(params)),
    domain: "onenote",
    readOnly: true,
  },
  {
    name: "list_onenote_pages",
    description: "List pages in a OneNote section",
    parameters: z.object({
      section_id: z.string().describe("Section ID"),
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listOnenotePages(params)),
    domain: "onenote",
    readOnly: true,
  },
  {
    name: "get_onenote_page_content",
    description: "Get OneNote page content as HTML",
    parameters: z.object({
      page_id: z.string().describe("Page ID"),
    }),
    execute: async (params) => unwrapResult(await getOnenotePageContent(params)),
    domain: "onenote",
    readOnly: true,
  },
  {
    name: "create_onenote_page",
    description:
      "Create a OneNote page in a section. Content is HTML — OneNote supports a constrained subset and silently drops unsupported CSS/tags.",
    parameters: z.object({
      section_id: z.string().describe("Section ID to create the page in"),
      title: z.string().describe("Page title"),
      content: z.string().describe("Page body as HTML (e.g. <p>Hello</p>)"),
    }),
    execute: async (params) => unwrapResult(await createOnenotePage(params)),
    domain: "onenote",
    readOnly: false,
  },
  {
    name: "update_onenote_page_content",
    description:
      "Append to or modify a OneNote page's content without rewriting it. Targets an element (default 'body') and applies an action. Content is HTML.",
    parameters: z.object({
      page_id: z.string().describe("Page ID to update"),
      content: z.string().describe("HTML content to apply"),
      action: z.enum(["append", "insert", "prepend", "replace"]).optional().describe("Update action (default: append)"),
      target: z.string().optional().describe("Target element data-id or known name (default: body)"),
      position: z.enum(["before", "after"]).optional().describe("Position relative to target, for insert"),
    }),
    execute: async (params) => unwrapResult(await updateOnenotePageContent(params)),
    domain: "onenote",
    readOnly: false,
  },
  {
    name: "create_onenote_section",
    description: "Create a new section in a OneNote notebook",
    parameters: z.object({
      notebook_id: z.string().describe("Notebook ID to create the section in"),
      display_name: z.string().describe("Section name"),
    }),
    execute: async (params) => unwrapResult(await createOnenoteSection(params)),
    domain: "onenote",
    readOnly: false,
  },
  {
    name: "create_onenote_notebook",
    description: "Create a new OneNote notebook",
    parameters: z.object({
      display_name: z.string().describe("Notebook name"),
    }),
    execute: async (params) => unwrapResult(await createOnenoteNotebook(params)),
    domain: "onenote",
    readOnly: false,
  },
  {
    name: "copy_onenote_page",
    description: "Copy a OneNote page to another section. Runs asynchronously in OneNote.",
    parameters: z.object({
      page_id: z.string().describe("Page ID to copy"),
      section_id: z.string().describe("Destination section ID"),
    }),
    execute: async (params) => unwrapResult(await copyOnenotePage(params)),
    domain: "onenote",
    readOnly: false,
  },
  {
    name: "delete_onenote_page",
    description: "Delete a OneNote page. This cannot be undone via the API.",
    parameters: z.object({
      page_id: z.string().describe("Page ID to delete"),
    }),
    execute: async (params) => unwrapResult(await deleteOnenotePage(params)),
    domain: "onenote",
    readOnly: false,
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
  },
  {
    name: "create_todo_list",
    description:
      "Create a new To Do list. Graph cannot share a list — the new list is private until someone shares it in the To Do app",
    parameters: z.object({
      display_name: z.string().describe("Name of the list. Must not match an existing list"),
    }),
    execute: async (params) => unwrapResult(await createTodoList(params)),
    domain: "todo",
    readOnly: false,
  },
  {
    name: "list_todo_tasks",
    description:
      "List tasks in a To Do list, or what changed in it since your last call. For changes, call once with " +
      "delta: true, which returns every task and a 'Delta token: …' line; keep the token and pass it as " +
      "delta_token next time to get only the tasks added, changed or removed since (removed ones as " +
      "'Removed: <id>' lines), and a new token. Completing a repeating task makes Graph keep the series and " +
      "split off the finished occurrence as a new, already-completed task with a new id, so an unseen " +
      "completed task is a done occurrence, not new work. Your own writes come back as changes too",
    parameters: z.object({
      list_id: z.string().describe("To Do list ID. A delta token works only with the list it came from"),
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
      extension_name: z
        .string()
        .optional()
        .describe(
          "Print this extension's fields at the end of each task line, as (Ext: {...}). With a change feed, " +
            "only calls with a delta_token can include them",
        ),
      delta: z
        .boolean()
        .optional()
        .describe("Start a change feed: return every task and a first delta token. Always reads every page"),
      delta_token: z
        .string()
        .optional()
        .describe(
          "The token from the last 'Delta token: …' line. Returns only what changed since. If the token is " +
            "refused, start again with delta: true",
        ),
    }),
    execute: async (params) => unwrapResult(await listTodoTasks(params)),
    domain: "todo",
    readOnly: true,
  },
  {
    name: "create_todo_task",
    description:
      "Create a new To Do task, optionally recurring. With linked_resource, the task is tagged with its source, " +
      "and a create whose external_id is already on a task in the list returns that task instead of a second " +
      "copy, so a retry is safe. That check reads the whole list first",
    parameters: z.object({
      list_id: z.string().describe("To Do list ID"),
      title: z.string().describe("Task title"),
      body: z.string().optional().describe("Task body/notes"),
      due_date: z.string().optional().describe("Due date (ISO format). Required when recurrence is set"),
      importance: z.string().optional().describe("Importance: low, normal, or high"),
      recurrence: RECURRENCE_PARAM.describe("Repeat pattern. Omit for a one-off task"),
      categories: TODO_CATEGORIES_PARAM,
      extension: TODO_EXTENSION_PARAM.describe("Your own data on the task, which the person can't see or edit"),
      linked_resource: z
        .object({
          external_id: z
            .string()
            .describe(
              "Your key for the source, e.g. 'civala:mail:<message id>'. At most 200 characters, with no line " +
                "break, |, (, ) or comma. One task per external_id per list",
            ),
          application_name: z.string().describe("The source app, e.g. 'Outlook'"),
          web_url: z.string().optional().describe("An http(s) link back to the source, shown in To Do"),
          display_name: z.string().optional().describe("A label for the link. Defaults to application_name"),
        })
        .optional()
        .describe("Tag the task with where it came from. Stored as the task's linkedResources"),
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
      due_date: z
        .string()
        .optional()
        .describe(
          "New due date (ISO format). On a repeating task this moves the series to start on that date; " +
            "To Do puts it on the first date the pattern allows from there. Must match recurrence.start_date if both are given",
        ),
      importance: z.string().optional().describe("Importance: low, normal, or high"),
      body: z.string().optional().describe("New body/notes"),
      recurrence: RECURRENCE_PARAM.describe("Replace the repeat pattern"),
      clear_recurrence: z.boolean().optional().describe("Remove the repeat pattern, making the task one-off"),
      categories: TODO_CATEGORIES_PARAM.describe(
        "Replace the task's categories with these names. Pass [] to clear them",
      ),
    }),
    execute: async (params) => unwrapResult(await updateTodoTask(params)),
    domain: "todo",
    readOnly: false,
  },
  {
    name: "set_todo_task_state",
    description:
      "Set only a To Do task's status, importance and your own extension fields, for unattended use. It can't " +
      "change the title, notes, dates or repeat pattern, or delete. The same call twice changes nothing more, so " +
      "a failed call is safe to repeat. Completing a repeating task is refused; use update_todo_task. Complete " +
      "only tasks you created, and change importance only on tasks carrying your extension, so the person's own " +
      "choices are never overwritten",
    parameters: SET_TODO_TASK_STATE_PARAMS,
    execute: async (params) => unwrapResult(await setTodoTaskState(params)),
    domain: "todo",
    readOnly: false,
  },

  {
    name: "delete_todo_task",
    description:
      "Delete a To Do task. There is no API to restore it, so treat it as permanent. Deleting a repeating task ends the whole series and needs force: true.",
    parameters: z.object({
      list_id: z.string().describe("To Do list ID"),
      task_id: z.string().describe("Task ID"),
      force: z
        .boolean()
        .optional()
        .describe(
          "Required to delete a repeating task, which ends the whole series rather than one occurrence. " +
            "To stop it repeating but keep the task, use clear_recurrence on update_todo_task instead.",
        ),
    }),
    execute: async (params) => unwrapResult(await deleteTodoTask(params)),
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
      version: z.enum(["v1.0", "beta"]).optional().describe("API version: v1.0 or beta (default: v1.0)"),
      headers: z
        .record(z.string(), z.string())
        .optional()
        .describe('Extra request headers, e.g. { "If-Match": "<etag>" } for concurrency-controlled writes'),
    }),
    execute: async (params) => unwrapResult(await graphQuery(params)),
    domain: "query",
    readOnly: false,
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
  const toRegister = toolDefinitions.filter((tool) => allowedTools.has(tool.name))

  toRegister.forEach((tool) => {
    server.addTool({
      name: tool.name,
      description: tool.description,
      parameters: tool.parameters,
      execute: wrapExecute(tool, oauthMode),
      annotations: annotationsFor(tool.name),
    })
  })

  console.error(
    `[Setup] Tools registered: ${toRegister.length}, skipped: ${toolDefinitions.length - toRegister.length}`,
  )
}

const buildUploadWorkflow = (allowedTools: Set<string>): string => {
  const hasFromPath = allowedTools.has("upload_file_from_path")
  const hasUploadConfig = allowedTools.has("get_upload_config")

  const bullets: string[] = allowedTools.has("upload_file")
    ? ["- Text content → upload_file (inline text, any transport)"]
    : []

  if (hasFromPath) {
    bullets.push(
      "- Binary files from the server's local disk → upload_file_from_path (requires absolute path on this machine)",
      "- Binary files generated in a cloud container (e.g., claude.ai) → first save to the user's local filesystem using Desktop Commander's write_file, then call upload_file_from_path with that local path",
    )
  }

  if (hasUploadConfig) {
    bullets.push(
      "- Binary files from HTTP/SSE deployments → get_upload_config returns an authenticated URL + curl command; execute the curl in a shell to upload without routing bytes through the LLM",
    )
  }

  return bullets.length > 0 ? `\n\nUpload workflows:\n${bullets.join("\n")}` : ""
}

const buildInstructions = (allowedTools: Set<string>): string => {
  const domains = new Set(toolDefinitions.filter((t) => allowedTools.has(t.name)).map((t) => t.domain))

  // No filter(Boolean) and no cast. ToolDefinition.domain is already ToolDomain, and
  // DOMAIN_DESCRIPTIONS is Record<ToolDomain, string>, so the index cannot miss: every domain is
  // guaranteed a line at compile time. Filtering could only hide a gap that no longer exists,
  // which is exactly how `rag` went unadvertised; a cast would imply a type gap that is not there.
  const capabilities = [...domains].map((d) => `- ${DOMAIN_DESCRIPTIONS[d]}`).join("\n")

  const uploadSection = domains.has("files") ? buildUploadWorkflow(allowedTools) : ""
  const attachmentSection = allowedTools.has("get_attachment_upload_config")
    ? "\n\nMail attachments:\n- Attach a file to a draft → get_attachment_upload_config returns an authenticated URL + curl command; execute the curl in a shell. The server attaches the file (up to 150 MB), so the sandbox only needs to reach this server, never Outlook."
    : ""

  return `A Microsoft 365 MCP server via Microsoft Graph API.\n\nAvailable capabilities:\n${capabilities}${uploadSection}${attachmentSection}`
}

type UploadRequestContext = {
  header: (name: string) => string | undefined
  query: (name: string) => string | undefined
  arrayBuffer: () => Promise<ArrayBuffer>
}

type RelayResponse = { status: number; body: unknown }

const relayAuth = async (req: UploadRequestContext, oauthMode: boolean): Promise<Either<RelayResponse, string>> => {
  const authHeader = req.header("authorization") ?? req.header("Authorization")
  const bearer = authHeader?.replace(/^Bearer\s+/i, "")
  const auth = await resolveUploadAccessToken(oauthMode, bearer)
  if (auth.error || !auth.token) {
    return Left({ status: auth.status ?? 401, body: { error: auth.error ?? "Unauthorized" } })
  }
  return Right(auth.token)
}

// Base64 is 4/3 the size of the bytes it carries; the slack covers any line breaks left in.
const maxRequestBytes = (encoding: string | undefined, maxBytes: number): number =>
  encoding === "base64" ? Math.ceil(((maxBytes * 4) / 3) * 1.02) + 1024 : maxBytes

// Refuses an oversized request on its declared length, before the body is read into memory.
const relayBody = async (
  req: UploadRequestContext,
  encoding: string | undefined,
  maxBytes: number,
): Promise<Either<RelayResponse, Buffer>> => {
  const declared = Number(req.header("content-length") ?? req.header("Content-Length"))
  if (Number.isFinite(declared) && declared > maxRequestBytes(encoding, maxBytes)) {
    return Left({ status: 413, body: { error: `File too large (max ${maxBytes} bytes)` } })
  }
  const raw = await req.arrayBuffer().then(
    (bytes) => Right<string, Buffer>(Buffer.from(bytes)),
    (error: unknown) =>
      Left<string, Buffer>(`Failed to read request body: ${error instanceof Error ? error.message : String(error)}`),
  )
  if (raw.isLeft()) return Left({ status: 400, body: { error: raw.value as string } })
  const rawBuffer = raw.value as Buffer
  if (rawBuffer.length === 0) return Left({ status: 400, body: { error: "Empty request body" } })

  const buffer = encoding === "base64" ? decodeBase64Upload(rawBuffer) : rawBuffer
  if (buffer.length === 0) return Left({ status: 400, body: { error: "Invalid base64 content" } })
  if (buffer.length > maxBytes) {
    return Left({ status: 413, body: { error: `File too large (max ${maxBytes} bytes)` } })
  }
  return Right(buffer)
}

// Only the versions the Graph version resolver knows reach the request path.
const relayApiBase = (req: UploadRequestContext): Either<RelayResponse, string> => {
  const version = req.query("apiVersion") ?? defaultVersion()
  return version === "v1.0" || version === "beta"
    ? Right(`${GRAPH_API_BASE}/${version}`)
    : Left({ status: 400, body: { error: 'apiVersion must be "v1.0" or "beta"' } })
}

// The upload URL names the mailbox in its query string, where a caller could edit it, so the
// relay applies the same mailbox rules the tool did, as the user whose token it carries.
const relayMailboxPrefix = async (req: UploadRequestContext, token: string): Promise<Either<RelayResponse, string>> => {
  const client = getGraphClient()
  if (client.isNone()) return Left({ status: 503, body: { error: "MS 365 client not initialized" } })
  const scope = await withToken(token, () => resolveMailbox(req.query("mailbox"), client.orThrow()))
  return scope.fold<Either<RelayResponse, string>>(
    (error) => Left({ status: 403, body: { error: error.message } }),
    (resolved) => Right(resolved.prefix),
  )
}

const handleAttachmentUpload = async (req: UploadRequestContext, oauthMode: boolean): Promise<RelayResponse> => {
  const messageId = req.query("messageId")
  const name = req.query("name")?.trim()
  if (!messageId || !name) {
    return { status: 400, body: { error: "messageId and name query parameters are required" } }
  }
  const apiBase = relayApiBase(req)
  if (apiBase.isLeft()) return apiBase.value as RelayResponse

  const auth = await relayAuth(req, oauthMode)
  if (auth.isLeft()) return auth.value as RelayResponse
  const token = auth.value as string

  const mailbox = await relayMailboxPrefix(req, token)
  if (mailbox.isLeft()) return mailbox.value as RelayResponse
  const prefix = mailbox.value as string

  // In oauth-proxy mode any bearer passes relayAuth; this one GET is what proves it, so it runs
  // before the body (up to 150 MB) is read. It also enforces the draft rule.
  const draft = await requireDraftMessage(apiBase.value as string, messageId, token, prefix)
  if (draft.isLeft()) {
    const err = draft.value as { message: string; status?: number }
    return { status: err.status ?? 502, body: { error: err.message } }
  }

  const body = await relayBody(req, req.query("encoding"), MAX_ATTACHMENT_SIZE)
  if (body.isLeft()) return body.value as RelayResponse
  const buffer = body.value as Buffer

  const contentType = resolveUploadContentType(req.query("contentType"), name)
  console.error(`[AttachmentUpload] bytes=${buffer.length} contentType=${contentType}`)

  const result = await attachFile(apiBase.value as string, messageId, token, buffer, name, contentType, prefix)
  if (result.isLeft()) {
    const err = result.value as { message: string; status?: number }
    return { status: err.status ?? 502, body: { error: err.message } }
  }
  return { status: 200, body: result.value }
}

const handleUpload = async (req: UploadRequestContext, oauthMode: boolean): Promise<RelayResponse> => {
  const path = req.query("path")
  if (!path) return { status: 400, body: { error: "path query parameter is required" } }
  if (!/:\/content$/i.test(path)) {
    return { status: 400, body: { error: 'path must end with ":/content"' } }
  }

  const apiBase = relayApiBase(req)
  if (apiBase.isLeft()) return apiBase.value as RelayResponse
  const conflictBehavior = req.query("conflictBehavior") ?? "rename"
  const explicitContentType = req.query("contentType")
  const encoding = req.query("encoding")

  const auth = await relayAuth(req, oauthMode)
  if (auth.isLeft()) return auth.value as RelayResponse
  const token = auth.value as string

  const body = await relayBody(req, encoding, MAX_UPLOAD_SIZE)
  if (body.isLeft()) return body.value as RelayResponse
  const buffer = body.value as Buffer

  const filename = filenameFromPath(path)
  const contentType = resolveUploadContentType(explicitContentType, filename)
  const base = apiBase.value as string

  console.error(
    `[Upload] path=${path} bytes=${buffer.length} contentType=${contentType} mode=${buffer.length <= SIMPLE_UPLOAD_LIMIT ? "simple" : "session"}`,
  )

  const result =
    buffer.length <= SIMPLE_UPLOAD_LIMIT
      ? await simpleUpload(base, path, token, buffer, contentType, conflictBehavior)
      : await sessionUpload(base, path, token, buffer, conflictBehavior)

  if (result.isLeft()) {
    const err = result.value as { message: string; status?: number }
    return { status: err.status ?? 500, body: { error: err.message } }
  }

  return { status: 200, body: result.value }
}

const mountRelayRoute = (
  server: FastMCP,
  path: string,
  handle: (req: UploadRequestContext, oauthMode: boolean) => Promise<RelayResponse>,
  oauthMode: boolean,
): void => {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any -- Hono app surface
  const app: any = (server as unknown as { getApp?: () => unknown }).getApp?.()
  if (!app) {
    console.error(`[Upload] FastMCP.getApp() unavailable; ${path} endpoint not mounted`)
    return
  }

  /* eslint-disable @typescript-eslint/no-explicit-any */
  const handler = async (c: any) => {
    try {
      const result = await handle(c.req as UploadRequestContext, oauthMode)
      return c.json(result.body, result.status)
    } catch (err) {
      const { message } = describeFetchError(err)
      console.error(`[Upload] ${path} unhandled error:`, message)
      return c.json({ error: message }, 500)
    }
  }

  app.post(path, handler)
  app.put(path, handler)
  /* eslint-enable @typescript-eslint/no-explicit-any */
  console.error(`[Setup] ${path} endpoint mounted (POST, PUT)`)
}

// Each relay route is mounted only when its tool survives the filters, so read-only mode, presets
// and MS365_ENABLED_TOOLS close the HTTP route as well as hiding the tool.
const mountRelayRoutes = (server: FastMCP, allowedTools: Set<string>, oauthMode: boolean): void => {
  if (allowedTools.has("get_upload_config")) mountRelayRoute(server, "/upload", handleUpload, oauthMode)
  if (allowedTools.has("get_attachment_upload_config")) {
    mountRelayRoute(server, "/upload/attachment", handleAttachmentUpload, oauthMode)
  }
}

// === Server Startup ===
// Stateless by default. This server consumes no server->client features (no roots,
// sampling, elicitation or progress anywhere in the package), and its OAuth tokens are
// keyed by bearer + TOKEN_STORAGE_PATH rather than by MCP session, so per-session
// transport state buys nothing and dies with the container on every redeploy.
// Opt out with MCP_STATELESS=false.
const statelessTransport = (): boolean => process.env.MCP_STATELESS?.trim().toLowerCase() !== "false"

// The only keepalive that works stateless: with no standing server-to-client stream, an
// in-flight tool call's own stream is the sole route, and a long silent call (a large
// read_document) can otherwise be closed as idle by a proxy.
const STREAM_KEEPALIVE = { enabled: true } as const

const main = async () => {
  const authConfig = resolveAuthConfig()
  const oauthMode = authConfig.mode === "oauth-proxy"

  const transport: "stdio" | "httpStream" = oauthMode
    ? "httpStream"
    : process.env.TRANSPORT_TYPE === "httpStream"
      ? "httpStream"
      : "stdio"

  const filterConfig = resolveFilterConfig(transport)
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
      streamKeepalive: STREAM_KEEPALIVE,
      name: "microsoft365-mcp-server",
      version: VERSION,
      instructions: buildInstructions(allowedTools),
      auth: provider,
      health: { enabled: true, path: "/ping", message: "ok" },
    } as never)

    // Initialize graph client without credential-based auth (tokens come from session)
    initializeGraphClient({ getAccessToken })

    registerTools(server, allowedTools, true)
    mountRelayRoutes(server, allowedTools, true)

    const port = parseInt(process.env.PORT ?? "3000", 10)
    const host = process.env.HOST ?? process.env.FASTMCP_HOST ?? "127.0.0.1"
    await server.start({ transportType: "httpStream", httpStream: { port, host, stateless: statelessTransport() } })
    console.error(`[Server] MS 365 MCP Server v${VERSION}${BUILD} (OAuth proxy) running on ${host}:${port}`)
  } else {
    // Standard mode: credential-based auth
    await setupAuth()

    const server = new FastMCP({
      streamKeepalive: STREAM_KEEPALIVE,
      name: "microsoft365-mcp-server",
      version: VERSION,
      instructions: buildInstructions(allowedTools),
      health: { enabled: true, path: "/ping", message: "ok" },
    })

    registerTools(server, allowedTools, false)

    const transportType = process.env.TRANSPORT_TYPE ?? "stdio"

    if (transportType === "httpStream") {
      mountRelayRoutes(server, allowedTools, false)
      const port = parseInt(process.env.PORT ?? "3000", 10)
      const host = process.env.HOST ?? process.env.FASTMCP_HOST ?? "127.0.0.1"
      await server.start({ transportType: "httpStream", httpStream: { port, host, stateless: statelessTransport() } })
      console.error(`[Server] MS 365 MCP Server v${VERSION}${BUILD} running on ${host}:${port}`)
    } else {
      await server.start({ transportType: "stdio" })
      console.error(`[Server] MS 365 MCP Server v${VERSION}${BUILD} running on stdio`)
    }
  }
}

main().catch((error) => {
  console.error("[Fatal]", error)
  process.exit(1)
})
