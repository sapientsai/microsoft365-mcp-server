export const GRAPH_SCOPES = {
  // Mail
  MAIL_READ: "Mail.Read",
  MAIL_READWRITE: "Mail.ReadWrite",
  MAIL_SEND: "Mail.Send",

  // Calendar
  CALENDARS_READ: "Calendars.Read",
  CALENDARS_READWRITE: "Calendars.ReadWrite",
  // Required by findMeetingTimes to read attendees' free/busy. Distinct from the non-Shared
  // scopes above — Calendars.ReadWrite does NOT grant it.
  CALENDARS_READ_SHARED: "Calendars.Read.Shared",

  // Contacts
  CONTACTS_READ: "Contacts.Read",
  CONTACTS_READWRITE: "Contacts.ReadWrite",

  // Files (OneDrive/SharePoint)
  FILES_READ: "Files.Read",
  FILES_READWRITE: "Files.ReadWrite",
  FILES_READ_ALL: "Files.Read.All",

  // SharePoint Sites
  SITES_READ_ALL: "Sites.Read.All",
  SITES_READWRITE_ALL: "Sites.ReadWrite.All",

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

  // Online meetings and transcripts.
  //
  // Deliberately absent from DEFAULT_INTERACTIVE_SCOPES. Both require tenant admin consent, and a
  // non-admin user cannot consent past them — adding either to the defaults would break sign-in for
  // every existing deployment whose tenant has not granted it. Opt in per deployment with
  // MS365_EXTRA_SCOPES (see resolveExtraScopes below).
  //
  // Two distinct permissions, because the transcript tools make two distinct calls:
  //   ONLINE_MEETINGS_READ                 resolves a meeting (needed only for join_web_url lookup)
  //   ONLINE_MEETING_TRANSCRIPT_READ_ALL   lists transcripts and reads their content
  ONLINE_MEETINGS_READ: "OnlineMeetings.Read",
  ONLINE_MEETING_TRANSCRIPT_READ_ALL: "OnlineMeetingTranscript.Read.All",
} as const

// OIDC / OAuth2 scopes (not Graph permissions).
// OFFLINE_ACCESS: Azure AD issues a refresh token only when this scope is requested.
export const OIDC_SCOPES = {
  OFFLINE_ACCESS: "offline_access",
} as const

export const DEFAULT_INTERACTIVE_SCOPES: ReadonlyArray<string> = [
  OIDC_SCOPES.OFFLINE_ACCESS,
  GRAPH_SCOPES.USER_READ,
  GRAPH_SCOPES.USER_READ_ALL,
  GRAPH_SCOPES.MAIL_READ,
  GRAPH_SCOPES.MAIL_READWRITE,
  GRAPH_SCOPES.MAIL_SEND,
  GRAPH_SCOPES.CALENDARS_READWRITE,
  GRAPH_SCOPES.CALENDARS_READ_SHARED,
  GRAPH_SCOPES.CONTACTS_READ,
  GRAPH_SCOPES.FILES_READWRITE,
  GRAPH_SCOPES.TEAM_READ_BASIC_ALL,
  GRAPH_SCOPES.CHANNEL_READ_BASIC_ALL,
  GRAPH_SCOPES.CHANNEL_MESSAGE_SEND,
  GRAPH_SCOPES.TASKS_READWRITE,
  GRAPH_SCOPES.NOTES_READWRITE,
  GRAPH_SCOPES.GROUP_READ_ALL,
  GRAPH_SCOPES.CHAT_READWRITE,
  GRAPH_SCOPES.CHAT_MESSAGE_READ,
  GRAPH_SCOPES.CHAT_MESSAGE_SEND,
  GRAPH_SCOPES.CHANNEL_MESSAGE_READ_ALL,
  GRAPH_SCOPES.SITES_READ_ALL,
  GRAPH_SCOPES.SITES_READWRITE_ALL,
]

/**
 * Additional scopes requested on top of DEFAULT_INTERACTIVE_SCOPES, from MS365_EXTRA_SCOPES
 * (comma-separated). Exists so an operator can turn on an admin-consent permission — meeting
 * transcripts being the motivating case — without a code change and a release.
 *
 * Only the OAuth proxy path reads this. The credential-based modes request
 * `https://graph.microsoft.com/.default`, which means "whatever this app registration has already
 * been consented", so their scope set is governed by the Azure app registration rather than by the
 * server. There is nothing to wire on that path.
 *
 * Entries already in the defaults are dropped: Azure AD tolerates duplicates, but a duplicated
 * scope list is noise in the consent screen and in get_auth_status output.
 */
export const resolveExtraScopes = (raw: string | undefined = process.env.MS365_EXTRA_SCOPES): ReadonlyArray<string> => {
  if (!raw) return []

  const defaults = new Set<string>(DEFAULT_INTERACTIVE_SCOPES)
  const seen = new Set<string>()

  return raw
    .split(",")
    .map((scope) => scope.trim())
    .filter((scope) => {
      if (scope.length === 0 || defaults.has(scope) || seen.has(scope)) return false
      seen.add(scope)
      return true
    })
}

/** The full scope set to request in OAuth proxy mode: the defaults plus any MS365_EXTRA_SCOPES. */
export const resolveInteractiveScopes = (raw?: string): ReadonlyArray<string> => [
  ...DEFAULT_INTERACTIVE_SCOPES,
  ...resolveExtraScopes(raw),
]

// GRAPH_API_BASE is owned by @sapientsai/ms-graph-core; re-exported here for back-compat.
export { GRAPH_API_BASE } from "@sapientsai/ms-graph-core"
export const GRAPH_DEFAULT_SCOPE = "https://graph.microsoft.com/.default"
