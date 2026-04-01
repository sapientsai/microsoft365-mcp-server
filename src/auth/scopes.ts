export const GRAPH_SCOPES = {
  // Mail
  MAIL_READ: "Mail.Read",
  MAIL_READWRITE: "Mail.ReadWrite",
  MAIL_SEND: "Mail.Send",

  // Calendar
  CALENDARS_READ: "Calendars.Read",
  CALENDARS_READWRITE: "Calendars.ReadWrite",

  // Contacts
  CONTACTS_READ: "Contacts.Read",
  CONTACTS_READWRITE: "Contacts.ReadWrite",

  // Files (OneDrive/SharePoint)
  FILES_READ: "Files.Read",
  FILES_READWRITE: "Files.ReadWrite",
  FILES_READ_ALL: "Files.Read.All",

  // Teams
  TEAM_READ_BASIC_ALL: "Team.ReadBasic.All",
  CHANNEL_MESSAGE_SEND: "ChannelMessage.Send",
  CHANNEL_READ_BASIC_ALL: "Channel.ReadBasic.All",

  // Users
  USER_READ: "User.Read",
  USER_READ_ALL: "User.Read.All",

  // Groups
  GROUP_READ_ALL: "Group.Read.All",

  // Planner / Tasks
  TASKS_READ: "Tasks.Read",
  TASKS_READWRITE: "Tasks.ReadWrite",

  // OneNote
  NOTES_READ: "Notes.Read",
  NOTES_READWRITE: "Notes.ReadWrite",

  // Chats
  CHAT_READWRITE: "Chat.ReadWrite",
  CHAT_MESSAGE_READ: "ChatMessage.Read",
  CHAT_MESSAGE_SEND: "ChatMessage.Send",
  CHANNEL_MESSAGE_READ_ALL: "ChannelMessage.Read.All",

  // To Do
  // Uses Tasks.Read / Tasks.ReadWrite (same as Planner)
} as const

export const DEFAULT_INTERACTIVE_SCOPES: ReadonlyArray<string> = [
  GRAPH_SCOPES.USER_READ,
  GRAPH_SCOPES.USER_READ_ALL,
  GRAPH_SCOPES.MAIL_READ,
  GRAPH_SCOPES.MAIL_SEND,
  GRAPH_SCOPES.CALENDARS_READWRITE,
  GRAPH_SCOPES.CONTACTS_READ,
  GRAPH_SCOPES.FILES_READWRITE,
  GRAPH_SCOPES.TEAM_READ_BASIC_ALL,
  GRAPH_SCOPES.CHANNEL_READ_BASIC_ALL,
  GRAPH_SCOPES.CHANNEL_MESSAGE_SEND,
  GRAPH_SCOPES.TASKS_READWRITE,
  GRAPH_SCOPES.NOTES_READ,
  GRAPH_SCOPES.GROUP_READ_ALL,
  GRAPH_SCOPES.CHAT_READWRITE,
  GRAPH_SCOPES.CHAT_MESSAGE_READ,
  GRAPH_SCOPES.CHAT_MESSAGE_SEND,
  GRAPH_SCOPES.CHANNEL_MESSAGE_READ_ALL,
]

export const GRAPH_API_BASE = "https://graph.microsoft.com"
export const GRAPH_DEFAULT_SCOPE = "https://graph.microsoft.com/.default"
