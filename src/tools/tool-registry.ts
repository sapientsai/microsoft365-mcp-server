export type ToolDomain =
  | "auth"
  | "mail"
  | "calendar"
  | "contacts"
  | "files"
  | "chats"
  | "teams"
  | "users"
  | "groups"
  | "planner"
  | "onenote"
  | "todo"
  | "query"

export type TransportMode = "stdio" | "httpStream"

export type ToolMeta = {
  readonly name: string
  readonly domain: ToolDomain
  readonly readOnly: boolean
  readonly orgOnly: boolean
  readonly transportOnly?: TransportMode
}

export const PRESETS: Record<string, ReadonlyArray<ToolDomain>> = {
  personal: ["mail", "calendar", "contacts", "todo", "files", "onenote"],
  collaboration: ["chats", "teams", "planner", "groups"],
  productivity: ["mail", "calendar", "todo"],
  all: [
    "auth",
    "mail",
    "calendar",
    "contacts",
    "files",
    "chats",
    "teams",
    "users",
    "groups",
    "planner",
    "onenote",
    "todo",
    "query",
  ],
}

export const TOOL_METADATA: ReadonlyArray<ToolMeta> = [
  // Auth
  { name: "get_auth_status", domain: "auth", readOnly: true, orgOnly: false },
  { name: "list_accounts", domain: "auth", readOnly: true, orgOnly: false },
  { name: "switch_account", domain: "auth", readOnly: false, orgOnly: false },
  { name: "set_access_token", domain: "auth", readOnly: false, orgOnly: false },
  // Mail
  { name: "list_messages", domain: "mail", readOnly: true, orgOnly: false },
  { name: "get_message", domain: "mail", readOnly: true, orgOnly: false },
  { name: "send_message", domain: "mail", readOnly: false, orgOnly: false },
  { name: "reply_to_message", domain: "mail", readOnly: false, orgOnly: false },
  { name: "search_messages", domain: "mail", readOnly: true, orgOnly: false },
  { name: "create_draft", domain: "mail", readOnly: false, orgOnly: false },
  { name: "send_draft", domain: "mail", readOnly: false, orgOnly: false },
  // Calendar
  { name: "list_events", domain: "calendar", readOnly: true, orgOnly: false },
  { name: "get_event", domain: "calendar", readOnly: true, orgOnly: false },
  { name: "create_event", domain: "calendar", readOnly: false, orgOnly: false },
  { name: "update_event", domain: "calendar", readOnly: false, orgOnly: false },
  { name: "delete_event", domain: "calendar", readOnly: false, orgOnly: false },
  // Contacts
  { name: "list_contacts", domain: "contacts", readOnly: true, orgOnly: false },
  { name: "get_contact", domain: "contacts", readOnly: true, orgOnly: false },
  { name: "create_contact", domain: "contacts", readOnly: false, orgOnly: false },
  { name: "search_contacts", domain: "contacts", readOnly: true, orgOnly: false },
  // Files
  { name: "list_drive_items", domain: "files", readOnly: true, orgOnly: false },
  { name: "get_drive_item", domain: "files", readOnly: true, orgOnly: false },
  { name: "search_files", domain: "files", readOnly: true, orgOnly: false },
  { name: "download_file", domain: "files", readOnly: true, orgOnly: false },
  { name: "create_folder", domain: "files", readOnly: false, orgOnly: false },
  { name: "upload_file", domain: "files", readOnly: false, orgOnly: false },
  { name: "get_upload_config", domain: "files", readOnly: false, orgOnly: false, transportOnly: "httpStream" },
  { name: "upload_file_from_path", domain: "files", readOnly: false, orgOnly: false, transportOnly: "stdio" },
  // SharePoint
  { name: "list_sites", domain: "files", readOnly: true, orgOnly: true },
  { name: "get_site", domain: "files", readOnly: true, orgOnly: true },
  { name: "list_site_drives", domain: "files", readOnly: true, orgOnly: true },
  { name: "list_site_items", domain: "files", readOnly: true, orgOnly: true },
  { name: "search_site_files", domain: "files", readOnly: true, orgOnly: true },
  // Chats
  { name: "list_chats", domain: "chats", readOnly: true, orgOnly: true },
  { name: "list_chat_messages", domain: "chats", readOnly: true, orgOnly: true },
  { name: "send_chat_message", domain: "chats", readOnly: false, orgOnly: true },
  // Teams
  { name: "list_teams", domain: "teams", readOnly: true, orgOnly: true },
  { name: "list_channels", domain: "teams", readOnly: true, orgOnly: true },
  { name: "list_channel_messages", domain: "teams", readOnly: true, orgOnly: true },
  { name: "send_channel_message", domain: "teams", readOnly: false, orgOnly: true },
  // Users
  { name: "get_me", domain: "users", readOnly: true, orgOnly: false },
  { name: "list_users", domain: "users", readOnly: true, orgOnly: true },
  { name: "get_user", domain: "users", readOnly: true, orgOnly: true },
  // Groups
  { name: "list_groups", domain: "groups", readOnly: true, orgOnly: true },
  { name: "get_group", domain: "groups", readOnly: true, orgOnly: true },
  { name: "list_group_members", domain: "groups", readOnly: true, orgOnly: true },
  // Planner
  { name: "list_plans", domain: "planner", readOnly: true, orgOnly: true },
  { name: "list_planner_tasks", domain: "planner", readOnly: true, orgOnly: true },
  { name: "get_planner_task", domain: "planner", readOnly: true, orgOnly: true },
  { name: "create_planner_task", domain: "planner", readOnly: false, orgOnly: true },
  { name: "update_planner_task", domain: "planner", readOnly: false, orgOnly: true },
  // OneNote
  { name: "list_notebooks", domain: "onenote", readOnly: true, orgOnly: false },
  { name: "list_sections", domain: "onenote", readOnly: true, orgOnly: false },
  { name: "list_pages", domain: "onenote", readOnly: true, orgOnly: false },
  { name: "get_page_content", domain: "onenote", readOnly: true, orgOnly: false },
  // To Do
  { name: "list_todo_lists", domain: "todo", readOnly: true, orgOnly: false },
  { name: "list_todo_tasks", domain: "todo", readOnly: true, orgOnly: false },
  { name: "create_todo_task", domain: "todo", readOnly: false, orgOnly: false },
  { name: "update_todo_task", domain: "todo", readOnly: false, orgOnly: false },
  // Query
  { name: "graph_query", domain: "query", readOnly: false, orgOnly: false },
]

export type ToolFilterConfig = {
  readonly presets?: ReadonlyArray<string>
  readonly enabledPattern?: string
  readonly readOnly?: boolean
  readonly orgMode?: boolean
  readonly transport?: TransportMode
}

export const filterTools = (config: ToolFilterConfig): Set<string> => {
  const allowed = new Set<string>()

  const allowedDomains = new Set<ToolDomain>()
  if (config.presets && config.presets.length > 0) {
    for (const preset of config.presets) {
      const domains = PRESETS[preset]
      if (domains) {
        for (const d of domains) allowedDomains.add(d)
      }
    }
    // Always include auth tools
    allowedDomains.add("auth")
  }

  let enabledRegex: RegExp | undefined
  if (config.enabledPattern) {
    enabledRegex = new RegExp(config.enabledPattern, "i")
  }

  for (const meta of TOOL_METADATA) {
    // Preset filter: skip if presets are set and domain not included
    if (allowedDomains.size > 0 && !allowedDomains.has(meta.domain)) continue

    // Read-only filter: skip write tools
    if (config.readOnly && !meta.readOnly) continue

    // Org mode filter: skip org-only tools unless org mode is enabled
    if (meta.orgOnly && !config.orgMode) continue

    // Regex filter: skip if doesn't match pattern
    if (enabledRegex && !enabledRegex.test(meta.name)) continue

    // Transport filter: skip if tool is restricted to a different transport
    if (meta.transportOnly && config.transport && meta.transportOnly !== config.transport) continue

    allowed.add(meta.name)
  }

  return allowed
}
