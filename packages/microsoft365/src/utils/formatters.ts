import { formatBytes } from "@sapientsai/ms-graph-core"
import { Option } from "functype"

import type {
  GraphAttachment,
  GraphBucket,
  GraphCallTranscript,
  GraphChannel,
  GraphChannelMessage,
  GraphChat,
  GraphChatMessage,
  GraphContact,
  GraphDriveItem,
  GraphEvent,
  GraphGroup,
  GraphMailFolder,
  GraphMeetingTimeSuggestion,
  GraphMeetingTimeSuggestionsResult,
  GraphMessage,
  GraphNotebook,
  GraphPage,
  GraphPlan,
  GraphPlannerTask,
  GraphSection,
  GraphTodoList,
  GraphTodoTask,
  GraphUser,
} from "../types"

// Mail
export const formatMessageSummary = (msg: GraphMessage): string => {
  const from = Option(msg.from?.emailAddress.name).fold(
    () => msg.from?.emailAddress.address ?? "Unknown",
    (v) => v,
  )
  const read = msg.isRead ? "" : " [Unread]"
  const attachments = msg.hasAttachments ? " [Attachments]" : ""
  return `- **${msg.subject ?? "(No Subject)"}** from ${from} (${msg.receivedDateTime ?? ""})${read}${attachments} (ID: ${msg.id})`
}

const REFERENCE_ATTACHMENT = "#microsoft.graph.referenceAttachment"
const ITEM_ATTACHMENT = "#microsoft.graph.itemAttachment"

// The read_document path is included per attachment on purpose: it is the only way to get at the
// content, and deriving it by hand is easy to get wrong (the trailing /$value is required).
//
// It is only emitted for attachments that actually have bytes in the mailbox. read_document reads
// the /$value stream, which only a fileAttachment serves. A referenceAttachment is a OneDrive or
// Dropbox link and stores nothing; an itemAttachment is an embedded Outlook item, reachable as MIME
// but not as a document. Printing the path for either promises content that endpoint cannot return,
// and a caller that follows it gets an opaque failure instead of "there is nothing here to read".
//
// Graph returns @odata.type on every attachment whether or not it is $select-ed, so this costs
// nothing extra. An unrecognised type still gets the path — better to offer a read that might work
// than to hide a file attachment behind a type name we have not seen before.
export const formatAttachmentSummary = (messageId: string, att: GraphAttachment): string => {
  const inline = att.isInline ? " [inline]" : ""
  const type = att.contentType ?? "unknown type"
  const size = Option(att.size)
    .map(formatBytes)
    .fold(
      () => "unknown size",
      (v) => v,
    )
  const header = `- **${att.name ?? "(unnamed)"}** (${type}, ${size})${inline}`

  switch (att["@odata.type"]) {
    case REFERENCE_ATTACHMENT:
      return `${header}\n  cloud link — no file is stored in the mailbox, so read_document cannot fetch it`
    case ITEM_ATTACHMENT:
      return `${header}\n  embedded Outlook item — not readable with read_document`
    default:
      return `${header}\n  read_document path: /me/messages/${messageId}/attachments/${att.id}/$value`
  }
}

export const formatAttachmentList = (messageId: string, attachments: ReadonlyArray<GraphAttachment>): string =>
  attachments.length === 0
    ? "No attachments found."
    : `# Attachments\n\n${attachments.map((a) => formatAttachmentSummary(messageId, a)).join("\n")}`

export const formatMessageList = (messages: ReadonlyArray<GraphMessage>): string =>
  messages.length === 0 ? "No messages found." : `# Messages\n\n${messages.map(formatMessageSummary).join("\n")}`

export const formatMailFolderSummary = (folder: GraphMailFolder): string => {
  const counts = `${folder.totalItemCount ?? 0} items, ${folder.unreadItemCount ?? 0} unread`
  // Graph's /me/mailFolders returns immediate children of the root only, so a folder's own
  // subfolders are absent from this listing. Printing the count is what makes them discoverable —
  // otherwise a caller sees "Inbox" and has no way to learn that Inbox/Clients exists at all.
  const children = (folder.childFolderCount ?? 0) > 0 ? `, ${folder.childFolderCount} subfolders` : ""
  return `- **${folder.displayName ?? "(Unnamed)"}** (${counts}${children}) (ID: ${folder.id})`
}

export const formatMailFolderList = (folders: ReadonlyArray<GraphMailFolder>): string =>
  folders.length === 0
    ? "No mail folders found."
    : `# Mail Folders\n\n${folders.map(formatMailFolderSummary).join("\n")}\n\n` +
      `Top-level folders only. A folder showing subfolders has children that are not listed here; ` +
      `reach them with graph_query on /me/mailFolders/{id}/childFolders, then pass the subfolder's ` +
      `ID to move_message.`

export const formatMessageDetail = (msg: GraphMessage): string => {
  const from = Option(msg.from?.emailAddress)
    .map((e) => `${e.name ?? ""} <${e.address ?? ""}>`)
    .fold(
      () => "Unknown",
      (v) => v,
    )

  const to = Option(msg.toRecipients)
    .map((recipients) =>
      recipients.map((r) => `${r.emailAddress.name ?? ""} <${r.emailAddress.address ?? ""}>`.trim()).join(", "),
    )
    .fold(
      () => "",
      (v) => v,
    )

  const body = Option(msg.body?.content).fold(
    () => msg.bodyPreview ?? "",
    (v) => v,
  )

  return `# ${msg.subject ?? "(No Subject)"}

## Details
- ID: ${msg.id}
- From: ${from}
- To: ${to}
- Date: ${msg.receivedDateTime ?? ""}
- Read: ${msg.isRead ? "Yes" : "No"}
- Importance: ${msg.importance ?? "normal"}
- Has Attachments: ${msg.hasAttachments ? "Yes" : "No"}

## Body
${body}`
}

// Calendar
export const formatEventSummary = (event: GraphEvent): string => {
  const start = Option(event.start?.dateTime).fold(
    () => "",
    (v) => v,
  )
  const location = Option(event.location?.displayName)
    .map((loc) => ` @ ${loc}`)
    .fold(
      () => "",
      (v) => v,
    )
  const cancelled = event.isCancelled ? " [Cancelled]" : ""
  return `- **${event.subject ?? "(No Subject)"}** (${start})${location}${cancelled} (ID: ${event.id})`
}

export const formatEventList = (events: ReadonlyArray<GraphEvent>): string =>
  events.length === 0 ? "No events found." : `# Events\n\n${events.map(formatEventSummary).join("\n")}`

export const formatEventDetail = (event: GraphEvent): string => {
  const organizer = Option(event.organizer?.emailAddress)
    .map((e) => `${e.name ?? ""} <${e.address ?? ""}>`)
    .fold(
      () => "Unknown",
      (v) => v,
    )

  const attendees = Option(event.attendees)
    .map((atts) =>
      atts
        .map((a) => {
          const name = `${a.emailAddress.name ?? ""} <${a.emailAddress.address ?? ""}>`
          const status = Option(a.status?.response).fold(
            () => "",
            (r) => ` (${r})`,
          )
          return `  - ${name}${status}`
        })
        .join("\n"),
    )
    .fold(
      () => "None",
      (v) => v,
    )

  const meetingUrl = Option(event.onlineMeeting?.joinUrl)
    .map((url) => `\n- Meeting URL: ${url}`)
    .fold(
      () => "",
      (v) => v,
    )

  const body = Option(event.body?.content).fold(
    () => "",
    (v) => v,
  )

  return `# ${event.subject ?? "(No Subject)"}

## Details
- ID: ${event.id}
- Start: ${event.start?.dateTime ?? ""} (${event.start?.timeZone ?? ""})
- End: ${event.end?.dateTime ?? ""} (${event.end?.timeZone ?? ""})
- Location: ${event.location?.displayName ?? "None"}
- Organizer: ${organizer}
- All Day: ${event.isAllDay ? "Yes" : "No"}
- Cancelled: ${event.isCancelled ? "Yes" : "No"}${meetingUrl}

## Attendees
${attendees}

## Body
${body}`
}

const formatMeetingTimeSuggestion = (suggestion: GraphMeetingTimeSuggestion): string => {
  const slot = suggestion.meetingTimeSlot
  const start = slot?.start?.dateTime ?? ""
  const end = slot?.end?.dateTime ?? ""
  const confidence = Option(suggestion.confidence)
    .map((c) => ` — ${c}% confidence`)
    .fold(
      () => "",
      (v) => v,
    )
  const organizer = Option(suggestion.organizerAvailability)
    .map((a) => `\n  - Organizer: ${a}`)
    .fold(
      () => "",
      (v) => v,
    )
  // Pass each attendee's availability through, including "unknown" (external/cross-tenant
  // attendees Graph can't see free/busy on) — the slot is still bookable.
  const attendees = Option(suggestion.attendeeAvailability)
    .map((list) =>
      list
        .map((a) => `\n  - ${a.attendee?.emailAddress?.address ?? "unknown"}: ${a.availability ?? "unknown"}`)
        .join(""),
    )
    .fold(
      () => "",
      (v) => v,
    )
  return `- **${start} → ${end}**${confidence}${organizer}${attendees}`
}

export const formatMeetingTimeSuggestions = (result: GraphMeetingTimeSuggestionsResult): string => {
  const suggestions = result.meetingTimeSuggestions ?? []
  if (suggestions.length === 0) {
    // Surface emptySuggestionsReason (e.g. AttendeesUnavailable) so callers know WHY nothing
    // came back rather than getting a silent empty list.
    const reason = Option(result.emptySuggestionsReason)
      .filter((r) => r.trim() !== "")
      .map((r) => ` (reason: ${r})`)
      .fold(
        () => "",
        (v) => v,
      )
    return `No common availability found.${reason}`
  }
  return `# Meeting Time Suggestions\n\n${suggestions.map(formatMeetingTimeSuggestion).join("\n")}`
}

// Contacts
export const formatContactSummary = (contact: GraphContact): string => {
  const email = Option(contact.emailAddresses?.[0]?.address)
    .map((e) => ` <${e}>`)
    .fold(
      () => "",
      (v) => v,
    )
  const company = Option(contact.companyName)
    .map((c) => ` - ${c}`)
    .fold(
      () => "",
      (v) => v,
    )
  return `- **${contact.displayName ?? "Unknown"}**${email}${company} (ID: ${contact.id})`
}

export const formatContactList = (contacts: ReadonlyArray<GraphContact>): string =>
  contacts.length === 0 ? "No contacts found." : `# Contacts\n\n${contacts.map(formatContactSummary).join("\n")}`

export const formatContactDetail = (contact: GraphContact): string => {
  const emails = Option(contact.emailAddresses)
    .map((addrs) => addrs.map((e) => `  - ${e.address ?? ""}`).join("\n"))
    .fold(
      () => "None",
      (v) => v,
    )

  const phones =
    [
      ...(contact.businessPhones ?? []).map((p) => `  - Business: ${p}`),
      ...Option(contact.mobilePhone)
        .map((p) => `  - Mobile: ${p}`)
        .fold(
          () => [] as string[],
          (v) => [v],
        ),
    ].join("\n") || "None"

  return `# ${contact.displayName ?? "Unknown"}

## Details
- ID: ${contact.id}
- First Name: ${contact.givenName ?? ""}
- Last Name: ${contact.surname ?? ""}
- Company: ${contact.companyName ?? ""}
- Job Title: ${contact.jobTitle ?? ""}

## Email Addresses
${emails}

## Phone Numbers
${phones}`
}

// Files
export const formatDriveItemSummary = (item: GraphDriveItem): string => {
  const type = item.folder ? `Folder (${item.folder.childCount ?? 0} items)` : (item.file?.mimeType ?? "File")
  const size = Option(item.size)
    .map((s) => ` (${formatBytes(s)})`)
    .fold(
      () => "",
      (v) => v,
    )
  return `- **${item.name ?? "Untitled"}** (ID: ${item.id}) - ${type}${size}`
}

export const formatDriveItemList = (items: ReadonlyArray<GraphDriveItem>): string =>
  items.length === 0 ? "No files found." : `# Files\n\n${items.map(formatDriveItemSummary).join("\n")}`

export const formatDriveItemDetail = (item: GraphDriveItem): string => {
  const downloadUrl = Option(item["@microsoft.graph.downloadUrl"])
    .map((url) => `\n- Download URL: ${url}`)
    .fold(
      () => "",
      (v) => v,
    )

  return `# ${item.name ?? "Untitled"}

## Details
- ID: ${item.id}
- Type: ${item.folder ? "Folder" : "File"}
- Size: ${formatBytes(item.size ?? 0)}
- MIME Type: ${item.file?.mimeType ?? "N/A"}
- Last Modified: ${item.lastModifiedDateTime ?? ""}
- Modified By: ${item.lastModifiedBy?.user?.displayName ?? "Unknown"}
- Web URL: ${item.webUrl ?? ""}${downloadUrl}`
}

// Teams
export const formatTeamSummary = (team: { id: string; displayName?: string; description?: string }): string =>
  `- **${team.displayName ?? "Untitled"}** (ID: ${team.id})`

export const formatTeamList = (
  teams: ReadonlyArray<{ id: string; displayName?: string; description?: string }>,
): string => (teams.length === 0 ? "No teams found." : `# Teams\n\n${teams.map(formatTeamSummary).join("\n")}`)

export const formatChannelSummary = (channel: GraphChannel): string =>
  `- **${channel.displayName ?? "Untitled"}** (${channel.membershipType ?? "standard"}, ID: ${channel.id})`

export const formatChannelList = (channels: ReadonlyArray<GraphChannel>): string =>
  channels.length === 0 ? "No channels found." : `# Channels\n\n${channels.map(formatChannelSummary).join("\n")}`

export const formatTranscriptSummary = (transcript: GraphCallTranscript): string => {
  const organizer = transcript.meetingOrganizer?.user?.displayName ?? transcript.meetingOrganizer?.user?.id
  const when = transcript.createdDateTime ?? "unknown date"
  const organizerNote = organizer ? `, organizer: ${organizer}` : ""
  return `- **${when}**${organizerNote}\n  - transcript_id: \`${transcript.id}\`${
    transcript.meetingId ? `\n  - meeting_id: \`${transcript.meetingId}\`` : ""
  }`
}

// The ids are what get_meeting_transcript needs next, so they are printed verbatim rather than
// summarized — a truncated transcript id is useless to the caller.
export const formatTranscriptList = (transcripts: ReadonlyArray<GraphCallTranscript>): string =>
  transcripts.length === 0
    ? "No transcripts found for this meeting. Transcription must have been on during the meeting, and " +
      "the transcript can take a few minutes to appear after it ends."
    : `# Meeting transcripts (${transcripts.length})\n\n${transcripts.map(formatTranscriptSummary).join("\n")}`

export const formatChannelMessageSummary = (msg: GraphChannelMessage): string => {
  const from = Option(msg.from?.user?.displayName).fold(
    () => "Unknown",
    (v) => v,
  )
  const content = Option(msg.body?.content)
    .map((c) => c.substring(0, 100) + (c.length > 100 ? "..." : ""))
    .fold(
      () => "",
      (v) => v,
    )
  return `- **${from}** (${msg.createdDateTime ?? ""}): ${content}`
}

export const formatChannelMessageList = (msgs: ReadonlyArray<GraphChannelMessage>): string =>
  msgs.length === 0 ? "No messages found." : `# Channel Messages\n\n${msgs.map(formatChannelMessageSummary).join("\n")}`

// Users
export const formatUserSummary = (user: GraphUser): string => {
  const email = Option(user.mail)
    .map((e) => ` <${e}>`)
    .fold(
      () => "",
      (v) => v,
    )
  return `- **${user.displayName ?? "Unknown"}**${email}`
}

export const formatUserList = (users: ReadonlyArray<GraphUser>): string =>
  users.length === 0 ? "No users found." : `# Users\n\n${users.map(formatUserSummary).join("\n")}`

export const formatUserDetail = (user: GraphUser): string =>
  `# ${user.displayName ?? "Unknown"}

## Details
- ID: ${user.id}
- Email: ${user.mail ?? "N/A"}
- UPN: ${user.userPrincipalName ?? "N/A"}
- Job Title: ${user.jobTitle ?? "N/A"}
- Department: ${user.department ?? "N/A"}
- Office: ${user.officeLocation ?? "N/A"}
- Mobile: ${user.mobilePhone ?? "N/A"}`

// Groups
export const formatGroupSummary = (group: GraphGroup): string =>
  `- **${group.displayName ?? "Unknown"}** (${group.mail ?? "no mail"})`

export const formatGroupList = (groups: ReadonlyArray<GraphGroup>): string =>
  groups.length === 0 ? "No groups found." : `# Groups\n\n${groups.map(formatGroupSummary).join("\n")}`

export const formatGroupDetail = (group: GraphGroup): string => {
  const types = Option(group.groupTypes)
    .map((t) => t.join(", "))
    .fold(
      () => "None",
      (v) => v,
    )

  return `# ${group.displayName ?? "Unknown"}

## Details
- ID: ${group.id}
- Mail: ${group.mail ?? "N/A"}
- Description: ${group.description ?? "N/A"}
- Group Types: ${types}
- Membership Rule: ${group.membershipRule ?? "N/A"}`
}

// Planner
export const formatPlanSummary = (plan: GraphPlan): string => `- **${plan.title ?? "Untitled"}** (ID: ${plan.id})`

export const formatPlanList = (plans: ReadonlyArray<GraphPlan>): string =>
  plans.length === 0 ? "No plans found." : `# Plans\n\n${plans.map(formatPlanSummary).join("\n")}`

export const formatBucketList = (buckets: ReadonlyArray<GraphBucket>): string =>
  buckets.length === 0
    ? "No buckets found."
    : `# Buckets\n\n${buckets.map((b) => `- **${b.name ?? "Unnamed"}** (ID: ${b.id})`).join("\n")}`

export const formatPlannerTaskSummary = (task: GraphPlannerTask): string => {
  const due = Option(task.dueDateTime)
    .map((d) => ` (Due: ${d})`)
    .fold(
      () => "",
      (v) => v,
    )
  const pct = Option(task.percentComplete).fold(
    () => "",
    (p) => ` [${p}%]`,
  )
  return `- **${task.title ?? "Untitled"}**${pct}${due} (ID: ${task.id})`
}

export const formatPlannerTaskList = (tasks: ReadonlyArray<GraphPlannerTask>): string =>
  tasks.length === 0 ? "No tasks found." : `# Planner Tasks\n\n${tasks.map(formatPlannerTaskSummary).join("\n")}`

export const formatPlannerTaskDetail = (task: GraphPlannerTask): string =>
  `# ${task.title ?? "Untitled"}

## Details
- ID: ${task.id}
- Plan ID: ${task.planId ?? "N/A"}
- Bucket ID: ${task.bucketId ?? "N/A"}
- Progress: ${task.percentComplete ?? 0}%
- Priority: ${task.priority ?? "N/A"}
- Due: ${task.dueDateTime ?? "N/A"}
- Created: ${task.createdDateTime ?? "N/A"}`

// OneNote
export const formatNotebookSummary = (nb: GraphNotebook): string => {
  const def = nb.isDefault ? " [Default]" : ""
  return `- **${nb.displayName ?? "Untitled"}**${def} (ID: ${nb.id})`
}

export const formatNotebookList = (notebooks: ReadonlyArray<GraphNotebook>): string =>
  notebooks.length === 0 ? "No notebooks found." : `# Notebooks\n\n${notebooks.map(formatNotebookSummary).join("\n")}`

export const formatSectionSummary = (section: GraphSection): string =>
  `- **${section.displayName ?? "Untitled"}** (ID: ${section.id})`

export const formatSectionList = (sections: ReadonlyArray<GraphSection>): string =>
  sections.length === 0 ? "No sections found." : `# Sections\n\n${sections.map(formatSectionSummary).join("\n")}`

export const formatPageSummary = (page: GraphPage): string =>
  `- **${page.title ?? "Untitled"}** (${page.lastModifiedDateTime ?? ""}) (ID: ${page.id})`

export const formatPageList = (pages: ReadonlyArray<GraphPage>): string =>
  pages.length === 0 ? "No pages found." : `# Pages\n\n${pages.map(formatPageSummary).join("\n")}`

// To Do
export const formatTodoListSummary = (list: GraphTodoList): string => {
  const wellKnown = Option(list.wellknownListName)
    .map((n) => ` [${n}]`)
    .fold(
      () => "",
      (v) => v,
    )
  return `- **${list.displayName ?? "Untitled"}**${wellKnown} (ID: ${list.id})`
}

export const formatTodoListList = (lists: ReadonlyArray<GraphTodoList>): string =>
  lists.length === 0 ? "No To Do lists found." : `# To Do Lists\n\n${lists.map(formatTodoListSummary).join("\n")}`

export const formatTodoTaskSummary = (task: GraphTodoTask): string => {
  const status = task.status ?? "notStarted"
  const due = Option(task.dueDateTime?.dateTime)
    .map((d) => ` (Due: ${d})`)
    .fold(
      () => "",
      (v) => v,
    )
  return `- **${task.title ?? "Untitled"}** [${status}]${due} (ID: ${task.id})`
}

export const formatTodoTaskList = (tasks: ReadonlyArray<GraphTodoTask>): string =>
  tasks.length === 0 ? "No tasks found." : `# To Do Tasks\n\n${tasks.map(formatTodoTaskSummary).join("\n")}`

export const formatTodoTaskDetail = (task: GraphTodoTask): string => {
  const body = Option(task.body?.content).fold(
    () => "",
    (v) => v,
  )

  return `# ${task.title ?? "Untitled"}

## Details
- ID: ${task.id}
- Status: ${task.status ?? "notStarted"}
- Importance: ${task.importance ?? "normal"}
- Due: ${task.dueDateTime?.dateTime ?? "N/A"}
- Completed: ${task.completedDateTime?.dateTime ?? "N/A"}
- Reminder: ${task.isReminderOn ? "Yes" : "No"}
- Created: ${task.createdDateTime ?? ""}
- Modified: ${task.lastModifiedDateTime ?? ""}

## Body
${body}`
}

// Chats
export const formatChatSummary = (chat: GraphChat): string => {
  const topic = chat.topic ?? chat.chatType ?? "Chat"
  const updated = Option(chat.lastUpdatedDateTime)
    .map((d) => ` (updated: ${d})`)
    .fold(
      () => "",
      (v) => v,
    )
  return `- **${topic}**${updated} (${chat.chatType ?? "unknown"}, ID: ${chat.id})`
}

export const formatChatList = (chats: ReadonlyArray<GraphChat>): string =>
  chats.length === 0 ? "No chats found." : `# Chats\n\n${chats.map(formatChatSummary).join("\n")}`

export const formatChatMessageSummary = (msg: GraphChatMessage): string => {
  const from = Option(msg.from?.user?.displayName).fold(
    () => "Unknown",
    (v) => v,
  )
  const content = Option(msg.body?.content)
    .map((c) => c.substring(0, 100) + (c.length > 100 ? "..." : ""))
    .fold(
      () => "",
      (v) => v,
    )
  return `- **${from}** (${msg.createdDateTime ?? ""}): ${content}`
}

export const formatChatMessageList = (msgs: ReadonlyArray<GraphChatMessage>): string =>
  msgs.length === 0 ? "No chat messages found." : `# Chat Messages\n\n${msgs.map(formatChatMessageSummary).join("\n")}`

// Auth Status
export const formatAuthStatus = (status: {
  mode: string
  authenticated: boolean
  scopes: ReadonlyArray<string>
  expiresAt?: string
}): string =>
  `# Authentication Status

- Mode: ${status.mode}
- Authenticated: ${status.authenticated ? "Yes" : "No"}
- Expires: ${status.expiresAt ?? "N/A"}

## Scopes
${status.scopes.length > 0 ? status.scopes.map((s) => `- ${s}`).join("\n") : "No scopes available"}`
