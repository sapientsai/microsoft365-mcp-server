export type ToolDomain =
  | "auth"
  | "mail"
  | "calendar"
  | "contacts"
  | "files"
  | "chats"
  | "teams"
  | "meetings"
  | "users"
  | "groups"
  | "planner"
  | "onenote"
  | "todo"
  | "query"
  | "rag"

export type TransportMode = "stdio" | "httpStream"

export type ToolMeta = {
  readonly name: string
  readonly domain: ToolDomain
  readonly readOnly: boolean
  readonly orgOnly: boolean
  readonly transportOnly?: TransportMode
  /** Shown to people in tool lists and permission prompts, so brand names keep their casing. */
  readonly title: string
  /**
   * A write that only creates something new in the user's own space and notifies nobody, such as
   * a draft. It is advertised with destructiveHint: false, which may let a client skip asking.
   * Absent on every other write: updates, deletes, sends, moves, and anything that reaches other
   * people, such as invites and assignments.
   */
  readonly additive?: true
  /** Reaches beyond this server's own tool set; advertised as openWorldHint: true. */
  readonly openWorld?: true
}

/**
 * One line per domain, forming the capability summary the model reads in the server's
 * instructions.
 *
 * `Record<ToolDomain, string>` is the point: a domain without a description is a typecheck
 * error, not a missing line. This map used to live in index.ts as a loose
 * `Record<string, string>` consumed with `.filter(Boolean)`, so when the `rag` domain was added
 * its description was simply never written — and the deployed server advertised every capability
 * except reading documents, the one it exists for. Nothing failed and nothing warned; the line
 * was just absent (issue #57).
 *
 * It lives here rather than beside the consumer because the thing that drifted was the
 * relationship between a domain and its description. Keeping both in this file means adding a
 * domain to the union above cannot compile until its description exists.
 */
export const DOMAIN_DESCRIPTIONS: Record<ToolDomain, string> = {
  auth: "Authentication: Check auth status and manage tokens",
  mail: "Mail: List, read, send, reply, search, and draft email messages; edit and delete drafts",
  calendar: "Calendar: List, view, create, update, and delete events; accept or decline invitations",
  contacts: "Contacts: List, view, create, update, delete, search, and export contacts",
  files:
    "Files: List, view, search, download OneDrive files; create folders; upload files (see Upload workflows below)",
  chats: "Chats: List Teams chats and messages; send chat messages, starting a chat by email if needed",
  teams: "Teams: List teams, channels, and messages; send channel messages and thread replies",
  meetings: "Meetings: List Teams meeting transcripts and read their text",
  users: "Users: View profiles and list users",
  groups: "Groups: List groups and group members",
  planner: "Planner: List plans and tasks; create and update tasks",
  onenote: "OneNote: List notebooks, sections, pages; read page content",
  todo: "To Do: List task lists and tasks; create, update, and delete tasks, including repeating ones",
  query: "Graph Query: Execute arbitrary Microsoft Graph API queries",
  rag: "Documents: Extract readable text from PDF, DOCX, XLSX and text files — in OneDrive, SharePoint, or attached to a message",
}

export const PRESETS: Record<string, ReadonlyArray<ToolDomain>> = {
  personal: ["mail", "calendar", "contacts", "todo", "files", "onenote"],
  collaboration: ["chats", "teams", "meetings", "planner", "groups"],
  productivity: ["mail", "calendar", "todo"],
  rag: ["rag", "files", "query"],
  all: [
    "auth",
    "mail",
    "calendar",
    "contacts",
    "files",
    "chats",
    "teams",
    "meetings",
    "users",
    "groups",
    "planner",
    "onenote",
    "todo",
    "query",
    "rag",
  ],
}

export const TOOL_METADATA: ReadonlyArray<ToolMeta> = [
  // Auth
  { name: "get_auth_status", title: "Get Sign-in Status", domain: "auth", readOnly: true, orgOnly: false },
  { name: "list_accounts", title: "List Signed-in Accounts", domain: "auth", readOnly: true, orgOnly: false },
  { name: "switch_account", title: "Switch Account", domain: "auth", readOnly: false, orgOnly: false },
  { name: "set_access_token", title: "Set Access Token", domain: "auth", readOnly: false, orgOnly: false },
  // Mail
  { name: "list_messages", title: "List Email Messages", domain: "mail", readOnly: true, orgOnly: false },
  { name: "scan_messages", title: "Scan Email Headers", domain: "mail", readOnly: true, orgOnly: false },
  { name: "get_message", title: "Read Email Message", domain: "mail", readOnly: true, orgOnly: false },
  { name: "list_mail_folders", title: "List Mail Folders", domain: "mail", readOnly: true, orgOnly: false },
  { name: "move_message", title: "Move Email Message", domain: "mail", readOnly: false, orgOnly: false },
  { name: "list_attachments", title: "List Email Attachments", domain: "mail", readOnly: true, orgOnly: false },
  {
    name: "batch_move_messages",
    title: "Move Email Messages in Bulk",
    domain: "mail",
    readOnly: false,
    orgOnly: false,
  },
  { name: "send_message", title: "Send Email", domain: "mail", readOnly: false, orgOnly: false },
  { name: "send_reply", title: "Send Email Reply", domain: "mail", readOnly: false, orgOnly: false },
  { name: "send_reply_all", title: "Send Email Reply to All", domain: "mail", readOnly: false, orgOnly: false },
  { name: "send_forward", title: "Forward Email", domain: "mail", readOnly: false, orgOnly: false },
  { name: "search_messages", title: "Search Email", domain: "mail", readOnly: true, orgOnly: false },
  {
    name: "create_draft",
    title: "Create Email Draft",
    additive: true,
    domain: "mail",
    readOnly: false,
    orgOnly: false,
  },
  {
    name: "create_reply_draft",
    title: "Create Reply Draft",
    additive: true,
    domain: "mail",
    readOnly: false,
    orgOnly: false,
  },
  {
    name: "create_reply_all_draft",
    title: "Create Reply-All Draft",
    additive: true,
    domain: "mail",
    readOnly: false,
    orgOnly: false,
  },
  {
    name: "create_forward_draft",
    title: "Create Forward Draft",
    additive: true,
    domain: "mail",
    readOnly: false,
    orgOnly: false,
  },
  { name: "send_draft", title: "Send Email Draft", domain: "mail", readOnly: false, orgOnly: false },
  { name: "update_draft", title: "Edit Email Draft", domain: "mail", readOnly: false, orgOnly: false },
  { name: "delete_draft", title: "Delete Email Draft", domain: "mail", readOnly: false, orgOnly: false },
  // Changes an existing draft, so it is not additive. httpStream only: the relay is an HTTP route.
  {
    name: "get_attachment_upload_config",
    title: "Get Draft Attachment Upload Link",
    domain: "mail",
    readOnly: false,
    orgOnly: false,
    transportOnly: "httpStream",
  },
  // Calendar
  { name: "list_events", title: "List Calendar Events", domain: "calendar", readOnly: true, orgOnly: false },
  {
    name: "list_calendar_view",
    title: "List Calendar Events in a Date Range",
    domain: "calendar",
    readOnly: true,
    orgOnly: false,
  },
  {
    name: "find_meeting_availability",
    title: "Find Meeting Times",
    domain: "calendar",
    readOnly: true,
    orgOnly: false,
  },
  { name: "get_event", title: "Read Calendar Event", domain: "calendar", readOnly: true, orgOnly: false },
  { name: "create_event", title: "Create Calendar Event", domain: "calendar", readOnly: false, orgOnly: false },
  { name: "update_event", title: "Update Calendar Event", domain: "calendar", readOnly: false, orgOnly: false },
  { name: "delete_event", title: "Delete Calendar Event", domain: "calendar", readOnly: false, orgOnly: false },
  {
    name: "respond_to_event",
    title: "Respond to Meeting Invitation",
    domain: "calendar",
    readOnly: false,
    orgOnly: false,
  },
  // Contacts
  { name: "list_contacts", title: "List Contacts", domain: "contacts", readOnly: true, orgOnly: false },
  { name: "get_contact", title: "Read Contact", domain: "contacts", readOnly: true, orgOnly: false },
  {
    name: "create_contact",
    title: "Create Contact",
    additive: true,
    domain: "contacts",
    readOnly: false,
    orgOnly: false,
  },
  { name: "search_contacts", title: "Search Contacts", domain: "contacts", readOnly: true, orgOnly: false },
  { name: "update_contact", title: "Update Contact", domain: "contacts", readOnly: false, orgOnly: false },
  { name: "delete_contact", title: "Delete Contact", domain: "contacts", readOnly: false, orgOnly: false },
  {
    name: "export_contacts",
    title: "Export Contacts to Local File",
    additive: true,
    domain: "contacts",
    readOnly: false,
    orgOnly: false,
    transportOnly: "stdio",
  },
  // Files
  { name: "list_drive_items", title: "List OneDrive Files", domain: "files", readOnly: true, orgOnly: false },
  { name: "get_drive_item", title: "Get OneDrive File Details", domain: "files", readOnly: true, orgOnly: false },
  { name: "search_files", title: "Search OneDrive Files", domain: "files", readOnly: true, orgOnly: false },
  { name: "download_file", title: "Download OneDrive File", domain: "files", readOnly: true, orgOnly: false },
  {
    name: "create_folder",
    title: "Create OneDrive Folder",
    additive: true,
    domain: "files",
    readOnly: false,
    orgOnly: false,
  },
  { name: "upload_file", title: "Upload Text File to OneDrive", domain: "files", readOnly: false, orgOnly: false },
  {
    name: "get_upload_config",
    title: "Get File Upload Link",
    additive: true,
    domain: "files",
    readOnly: false,
    orgOnly: false,
    transportOnly: "httpStream",
  },
  {
    name: "upload_file_from_path",
    title: "Upload Local File to OneDrive",
    domain: "files",
    readOnly: false,
    orgOnly: false,
    transportOnly: "stdio",
  },
  // SharePoint
  { name: "list_sites", title: "List SharePoint Sites", domain: "files", readOnly: true, orgOnly: true },
  { name: "get_site", title: "Get SharePoint Site", domain: "files", readOnly: true, orgOnly: true },
  {
    name: "list_site_drives",
    title: "List SharePoint Document Libraries",
    domain: "files",
    readOnly: true,
    orgOnly: true,
  },
  { name: "list_site_items", title: "List SharePoint Files", domain: "files", readOnly: true, orgOnly: true },
  { name: "search_site_files", title: "Search SharePoint Files", domain: "files", readOnly: true, orgOnly: true },
  // Chats
  { name: "list_chats", title: "List Teams Chats", domain: "chats", readOnly: true, orgOnly: true },
  { name: "list_chat_messages", title: "Read Teams Chat Messages", domain: "chats", readOnly: true, orgOnly: true },
  { name: "send_chat_message", title: "Send Teams Chat Message", domain: "chats", readOnly: false, orgOnly: true },
  // Teams
  { name: "list_teams", title: "List Teams", domain: "teams", readOnly: true, orgOnly: true },
  { name: "list_channels", title: "List Teams Channels", domain: "teams", readOnly: true, orgOnly: true },
  {
    name: "list_channel_messages",
    title: "Read Teams Channel Messages",
    domain: "teams",
    readOnly: true,
    orgOnly: true,
  },
  {
    name: "send_channel_message",
    title: "Post Teams Channel Message",
    domain: "teams",
    readOnly: false,
    orgOnly: true,
  },
  // Meetings (transcripts). Needs admin-consent scopes that are NOT in the defaults — see
  // MS365_EXTRA_SCOPES in the README. Both tools are reachable but will 403 without the grant.
  {
    name: "list_meeting_transcripts",
    title: "List Meeting Transcripts",
    domain: "meetings",
    readOnly: true,
    orgOnly: true,
  },
  {
    name: "get_meeting_transcript",
    title: "Read Meeting Transcript",
    domain: "meetings",
    readOnly: true,
    orgOnly: true,
  },
  // Users
  { name: "get_me", title: "Get My Profile", domain: "users", readOnly: true, orgOnly: false },
  { name: "list_users", title: "List People in Organization", domain: "users", readOnly: true, orgOnly: true },
  { name: "get_user", title: "Get Person's Profile", domain: "users", readOnly: true, orgOnly: true },
  // Groups
  { name: "list_groups", title: "List Groups", domain: "groups", readOnly: true, orgOnly: true },
  { name: "get_group", title: "Get Group", domain: "groups", readOnly: true, orgOnly: true },
  { name: "list_group_members", title: "List Group Members", domain: "groups", readOnly: true, orgOnly: true },
  // Planner
  { name: "list_plans", title: "List Planner Plans", domain: "planner", readOnly: true, orgOnly: true },
  { name: "list_planner_buckets", title: "List Planner Buckets", domain: "planner", readOnly: true, orgOnly: true },
  { name: "create_planner_bucket", title: "Create Planner Bucket", domain: "planner", readOnly: false, orgOnly: true },
  { name: "list_planner_tasks", title: "List Planner Tasks", domain: "planner", readOnly: true, orgOnly: true },
  { name: "get_planner_task", title: "Read Planner Task", domain: "planner", readOnly: true, orgOnly: true },
  { name: "create_planner_task", title: "Create Planner Task", domain: "planner", readOnly: false, orgOnly: true },
  { name: "update_planner_task", title: "Update Planner Task", domain: "planner", readOnly: false, orgOnly: true },
  {
    name: "update_planner_task_details",
    title: "Update Planner Task Details",
    domain: "planner",
    readOnly: false,
    orgOnly: true,
  },
  // OneNote
  {
    name: "list_onenote_notebooks",
    title: "List OneNote Notebooks",
    domain: "onenote",
    readOnly: true,
    orgOnly: false,
  },
  { name: "list_onenote_sections", title: "List OneNote Sections", domain: "onenote", readOnly: true, orgOnly: false },
  { name: "list_onenote_pages", title: "List OneNote Pages", domain: "onenote", readOnly: true, orgOnly: false },
  { name: "get_onenote_page_content", title: "Read OneNote Page", domain: "onenote", readOnly: true, orgOnly: false },
  {
    name: "create_onenote_page",
    title: "Create OneNote Page",
    additive: true,
    domain: "onenote",
    readOnly: false,
    orgOnly: false,
  },
  {
    name: "update_onenote_page_content",
    title: "Update OneNote Page",
    domain: "onenote",
    readOnly: false,
    orgOnly: false,
  },
  {
    name: "create_onenote_section",
    title: "Create OneNote Section",
    additive: true,
    domain: "onenote",
    readOnly: false,
    orgOnly: false,
  },
  {
    name: "create_onenote_notebook",
    title: "Create OneNote Notebook",
    additive: true,
    domain: "onenote",
    readOnly: false,
    orgOnly: false,
  },
  {
    name: "copy_onenote_page",
    title: "Copy OneNote Page",
    additive: true,
    domain: "onenote",
    readOnly: false,
    orgOnly: false,
  },
  { name: "delete_onenote_page", title: "Delete OneNote Page", domain: "onenote", readOnly: false, orgOnly: false },
  // To Do
  { name: "list_todo_lists", title: "List To Do Lists", domain: "todo", readOnly: true, orgOnly: false },
  { name: "list_todo_tasks", title: "List To Do Tasks", domain: "todo", readOnly: true, orgOnly: false },
  {
    name: "create_todo_task",
    title: "Create To Do Task",
    additive: true,
    domain: "todo",
    readOnly: false,
    orgOnly: false,
  },
  { name: "update_todo_task", title: "Update To Do Task", domain: "todo", readOnly: false, orgOnly: false },
  { name: "delete_todo_task", title: "Delete To Do Task", domain: "todo", readOnly: false, orgOnly: false },
  // Query
  {
    name: "graph_query",
    title: "Raw Microsoft Graph Request",
    openWorld: true,
    domain: "query",
    readOnly: false,
    orgOnly: false,
  },
  // RAG
  { name: "read_document", title: "Read Document Text", domain: "rag", readOnly: true, orgOnly: false },
]

export type ToolFilterConfig = {
  readonly presets?: ReadonlyArray<string>
  readonly enabledPattern?: string
  readonly readOnly?: boolean
  readonly orgMode?: boolean
  readonly transport?: TransportMode
  readonly requireDraft?: boolean
}

export type ToolAnnotations = {
  readonly title: string
  readonly readOnlyHint: boolean
  readonly destructiveHint?: boolean
  readonly openWorldHint?: true
}

/**
 * The MCP annotations a tool is advertised with. Claude reads them to decide when to ask: a
 * read-only tool can run without confirmation, and a destructive one always prompts. They are
 * built here from TOOL_METADATA, the one place each tool's title and kind is declared, so no
 * definition can claim a hint its metadata contradicts.
 */
export const toolAnnotations = (meta: ToolMeta): ToolAnnotations => ({
  title: meta.title,
  readOnlyHint: meta.readOnly,
  ...(meta.readOnly ? {} : { destructiveHint: meta.additive !== true }),
  ...(meta.openWorld ? { openWorldHint: true } : {}),
})

const ANNOTATIONS_BY_NAME: ReadonlyMap<string, ToolAnnotations> = new Map(
  TOOL_METADATA.map((meta) => [meta.name, toolAnnotations(meta)]),
)

export const annotationsFor = (name: string): ToolAnnotations | undefined => ANNOTATIONS_BY_NAME.get(name)

// Read per call, not captured at import, so tests can set it. graph_query reads it too: hiding the
// send_* tools alone would leave a raw POST /me/sendMail open.
export const requireDraftEnabled = (): boolean => process.env.MS365_REQUIRE_DRAFT === "true"

const DRAFT_BYPASS_TOOLS: ReadonlySet<string> = new Set([
  "send_message",
  "send_reply",
  "send_reply_all",
  "send_forward",
])

export const filterTools = (config: ToolFilterConfig): Set<string> => {
  // Native Set<string> is the returned contract (callers use .has()); the accumulation is functional.
  const allowedDomains =
    config.presets && config.presets.length > 0
      ? new Set<ToolDomain>([...config.presets.flatMap((preset) => PRESETS[preset] ?? []), "auth"])
      : undefined
  const enabledRegex = config.enabledPattern ? new RegExp(config.enabledPattern, "i") : undefined

  const included = (meta: ToolMeta): boolean => {
    if (allowedDomains && !allowedDomains.has(meta.domain)) return false // preset filter
    if (config.readOnly && !meta.readOnly) return false // read-only filter: skip write tools
    if (meta.orgOnly && !config.orgMode) return false // org-mode filter
    if (enabledRegex && !enabledRegex.test(meta.name)) return false // regex filter
    if (meta.transportOnly && config.transport && meta.transportOnly !== config.transport) return false // transport
    if (config.requireDraft && DRAFT_BYPASS_TOOLS.has(meta.name)) return false // require-draft filter
    return true
  }

  return new Set(TOOL_METADATA.filter(included).map((meta) => meta.name))
}
