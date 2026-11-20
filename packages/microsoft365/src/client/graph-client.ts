import {
  type AuthStrategy,
  createGraphRequest,
  drivePath,
  encodedId,
  GRAPH_API_BASE,
  graphPath,
  type GraphRequestOptions,
  mapHttpError,
  notSentError,
  raw,
  resilientFetch,
  uploadDriveContent,
} from "@sapientsai/ms-graph-core"
import { None, type Option, Ref, Some } from "functype"
import { type Either, Left, Right } from "functype/either"

import type {
  GraphApiError,
  GraphApiVersion,
  GraphAttachment,
  GraphAutomaticRepliesSetting,
  GraphBatchResponse,
  GraphBucket,
  GraphCalendar,
  GraphChannel,
  GraphChannelMessage,
  GraphChat,
  GraphChatMessage,
  GraphContact,
  GraphDrive,
  GraphDriveItem,
  GraphEvent,
  GraphGroup,
  GraphMailFolder,
  GraphMeetingTimeSuggestionsResult,
  GraphMessage,
  GraphMessageRule,
  GraphNotebook,
  GraphOutlookCategory,
  GraphPage,
  GraphPlan,
  GraphPlannerTask,
  GraphSection,
  GraphSite,
  GraphTodoList,
  GraphTodoTask,
  GraphUser,
  ODataParams,
  ODataResponse,
} from "../types"
import { isDeltaToken } from "../utils/delta-token"
import { safeDecode } from "../utils/safe-decode"
import { defaultVersion } from "./graph-version"
import { explainMailboxRefusal } from "./mailbox-hint"

/** A delta record: a task, or a removed one marked by "@removed". */
export type TodoDeltaRecord = GraphTodoTask & { readonly "@removed"?: { readonly reason?: string } }
type TodoDeltaPage = ODataResponse<TodoDeltaRecord> & { readonly "@odata.deltaLink"?: string }
export type TodoDeltaResult = { readonly records: ReadonlyArray<TodoDeltaRecord>; readonly deltaToken: string }

/** Tasks asked for per page of a To Do list, and the most pages read before giving up. */
export const TODO_TASK_PAGE_SIZE = 100
export const MAX_TODO_TASK_PAGES = 50

const createGraphClient = (auth: AuthStrategy) => {
  // Generic request machinery (auth injection, OData, error mapping, pagination) lives in
  // @sapientsai/ms-graph-core; this client builds the M365 domain methods on top of it.
  const core = createGraphRequest(auth, { defaultVersion })
  // A path is a plain string, or one graphPath built from caller IDs; a refused ID is a Left here,
  // so nothing is sent.
  const request = async <T>(
    method: string,
    path: string | Either<GraphApiError, string>,
    options?: GraphRequestOptions,
  ): Promise<Either<GraphApiError, T>> => {
    if (typeof path !== "string" && path.isLeft()) return Left<GraphApiError, T>(path.value as GraphApiError)
    const resolved = typeof path === "string" ? path : path.orThrow()
    return explainMailboxRefusal(method, resolved, await core.request<T>(method, resolved, options))
  }
  const requestPaginated = async <T>(
    path: string | Either<GraphApiError, string>,
    options?: GraphRequestOptions,
  ): Promise<Either<GraphApiError, ReadonlyArray<T>>> => {
    if (typeof path !== "string" && path.isLeft()) {
      return Left<GraphApiError, ReadonlyArray<T>>(path.value as GraphApiError)
    }
    const resolved = typeof path === "string" ? path : path.orThrow()
    return explainMailboxRefusal("GET", resolved, await core.requestPaginated<T>(resolved, options))
  }

  // Mail
  // Every mail path takes a prefix: "/me" for the signed-in user's own mailbox, or
  // "/users/{address}" for a delegated or shared one (see mail/mailbox.ts).
  // /me/messages spans every folder, Sent Items included. A folder takes a well-known name
  // (inbox, sentitems, ...) or a folder ID; Graph accepts either in the same path segment.
  const messagesPath = (folder?: string, prefix: string = "/me") =>
    folder ? graphPath`${raw(prefix)}/mailFolders/${encodedId(folder)}/messages` : graphPath`${raw(prefix)}/messages`

  const listMessages = (odataParams?: ODataParams, folder?: string, prefix: string = "/me") =>
    request<ODataResponse<GraphMessage>>("GET", messagesPath(folder, prefix), { odataParams })

  const listAllMessages = (odataParams?: ODataParams, folder?: string, prefix: string = "/me") =>
    requestPaginated<GraphMessage>(messagesPath(folder, prefix), { odataParams })

  // Prefer: outlook.body-content-type="text" makes Graph convert the body server-side.
  // Marketing mail is mostly CSS and layout tables — one newsletter measured 79,347
  // characters as HTML — so for a caller that only needs the words this is a ~95%
  // reduction, and better than stripping tags locally. Events honour the same header.
  const bodyFormat = (bodyContentType?: "text" | "html") =>
    bodyContentType ? { headers: { Prefer: `outlook.body-content-type="${bodyContentType}"` } } : undefined

  const getMessage = (id: string, bodyContentType?: "text" | "html", prefix: string = "/me") =>
    request<GraphMessage>("GET", graphPath`${raw(prefix)}/messages/${id}`, bodyFormat(bodyContentType))

  const listMailFolders = (odataParams?: ODataParams, prefix: string = "/me") =>
    request<ODataResponse<GraphMailFolder>>("GET", graphPath`${raw(prefix)}/mailFolders`, { odataParams })

  // Only follow Graph-issued mail collection links; never send a bearer token to another host.
  // Graph writes a folder-scoped nextLink in OData key form, /me/mailFolders('<id>')/messages, not
  // the /mailFolders/<id>/ form the first request used (confirmed live; 1.2.14 rejected it). A scan
  // of another mailbox continues under /users/{address}, in either form too — but only the mailbox
  // the scan addressed: a continuation that wandered into another mailbox is refused, not followed.
  const mailboxRoot = (root: string): string => {
    const keyed = /^\/users\('([^'/]+)'\)$/.exec(root)
    const address = keyed ? keyed[1]! : root.startsWith("/users/") ? root.slice("/users/".length) : undefined
    return address === undefined ? root : `/users/${safeDecode(address).toLowerCase()}`
  }

  const nextMessageScanPage = (nextLink: string, prefix: string = "/me") => {
    const url = URL.canParse(nextLink) ? new URL(nextLink) : undefined
    const match =
      /^\/(v1\.0|beta)((\/me|\/users(?:\/[^/]+|\('[^'/]+'\)))\/(?:mailFolders(?:\/[^/]+|\('[^'/]+'\))\/)?messages)$/.exec(
        url?.pathname ?? "",
      )
    if (
      url?.origin !== GRAPH_API_BASE ||
      url.username ||
      url.password ||
      url.hash ||
      !match ||
      mailboxRoot(match[3]!) !== mailboxRoot(prefix)
    ) {
      return Promise.resolve(
        Left<GraphApiError, ODataResponse<GraphMessage>>({
          type: "api",
          message: "Invalid Graph message continuation URL.",
        }),
      )
    }
    return request<ODataResponse<GraphMessage>>("GET", `${match[2]}${url.search}`, {
      version: match[1] as GraphApiVersion,
    })
  }

  const moveMessage = (id: string, destinationId: string, prefix: string = "/me") =>
    request<GraphMessage>("POST", graphPath`${raw(prefix)}/messages/${id}/move`, { body: { destinationId } })

  // JSON batching: up to 20 sub-requests per round-trip, each with its own status. The
  // only way to move thousands of messages without thousands of round-trips.
  const batchRequest = (requests: ReadonlyArray<Record<string, unknown>>) =>
    request<GraphBatchResponse>("POST", "/$batch", { body: { requests } })

  // $select omits contentBytes deliberately: fileAttachment includes the full base64 payload
  // by default, which would drag megabytes of binary through the model for a listing.
  const listAttachments = (messageId: string, prefix: string = "/me") =>
    request<ODataResponse<GraphAttachment>>("GET", graphPath`${raw(prefix)}/messages/${messageId}/attachments`, {
      odataParams: { $select: ["id", "name", "contentType", "size", "isInline", "lastModifiedDateTime"] },
    })

  const sendMessage = (message: Record<string, unknown>, prefix: string = "/me") =>
    request<Record<string, never>>("POST", graphPath`${raw(prefix)}/sendMail`, { body: message })

  const createDraft = (message: Record<string, unknown>, prefix: string = "/me") =>
    request<GraphMessage>("POST", graphPath`${raw(prefix)}/messages`, { body: message })

  const sendDraft = (messageId: string, prefix: string = "/me") =>
    request<Record<string, never>>("POST", graphPath`${raw(prefix)}/messages/${messageId}/send`)

  // Draft edits and deletes check isDraft first, so this fetches only what that check needs.
  const getDraftState = (id: string, prefix: string = "/me") =>
    request<GraphMessage>("GET", graphPath`${raw(prefix)}/messages/${id}`, {
      odataParams: { $select: ["id", "isDraft", "subject"] },
    })

  const updateMessage = (id: string, changes: Record<string, unknown>, prefix: string = "/me") =>
    request<GraphMessage>("PATCH", graphPath`${raw(prefix)}/messages/${id}`, { body: changes })

  const deleteMessage = (id: string, prefix: string = "/me") =>
    request<Record<string, never>>("DELETE", graphPath`${raw(prefix)}/messages/${id}`)

  // Mailbox settings: automatic replies, the master category list, and inbox rules.
  // All need MailboxSettings.Read or MailboxSettings.ReadWrite.
  const getAutomaticReplies = (prefix: string = "/me") =>
    request<GraphAutomaticRepliesSetting>("GET", graphPath`${raw(prefix)}/mailboxSettings/automaticRepliesSetting`)

  const setAutomaticReplies = (setting: Record<string, unknown>, prefix: string = "/me") =>
    request<{ automaticRepliesSetting?: GraphAutomaticRepliesSetting }>(
      "PATCH",
      graphPath`${raw(prefix)}/mailboxSettings`,
      {
        body: { automaticRepliesSetting: setting },
      },
    )

  const listMasterCategories = (prefix: string = "/me") =>
    request<ODataResponse<GraphOutlookCategory>>("GET", graphPath`${raw(prefix)}/outlook/masterCategories`)

  const createMasterCategory = (displayName: string, color: string, prefix: string = "/me") =>
    request<GraphOutlookCategory>("POST", graphPath`${raw(prefix)}/outlook/masterCategories`, {
      body: { displayName, color },
    })

  const getMessageCategories = (id: string, prefix: string = "/me") =>
    request<GraphMessage>("GET", graphPath`${raw(prefix)}/messages/${id}`, {
      odataParams: { $select: ["id", "subject", "categories"] },
    })

  // A rule's moveToFolder takes a real folder ID, so a well-known name such as "archive" is looked up.
  const getMailFolder = (idOrWellKnownName: string, prefix: string = "/me") =>
    request<GraphMailFolder>("GET", graphPath`${raw(prefix)}/mailFolders/${encodedId(idOrWellKnownName)}`, {
      odataParams: { $select: ["id", "displayName"] },
    })

  const rulesPath = (prefix: string) => graphPath`${raw(prefix)}/mailFolders/inbox/messageRules`

  const listMessageRules = (prefix: string = "/me") => requestPaginated<GraphMessageRule>(rulesPath(prefix))

  const createMessageRule = (rule: Record<string, unknown>, prefix: string = "/me") =>
    request<GraphMessageRule>("POST", rulesPath(prefix), { body: rule })

  // Rule IDs are base64 and may hold "/" or "+", so they are encoded as one path segment.
  const deleteMessageRule = (id: string, prefix: string = "/me") =>
    request<Record<string, never>>("DELETE", graphPath`${rulesPath(prefix)}/${encodedId(id)}`)

  const sendReply = (id: string, comment: string, prefix: string = "/me") =>
    request<Record<string, never>>("POST", graphPath`${raw(prefix)}/messages/${id}/reply`, { body: { comment } })

  // Draft-creating reply actions: return a threaded draft (original quoted) for review.
  const createReplyDraft = (id: string, comment: string, prefix: string = "/me") =>
    request<GraphMessage>("POST", graphPath`${raw(prefix)}/messages/${id}/createReply`, { body: { comment } })

  const createReplyAllDraft = (id: string, comment: string, prefix: string = "/me") =>
    request<GraphMessage>("POST", graphPath`${raw(prefix)}/messages/${id}/createReplyAll`, { body: { comment } })

  const createForwardDraft = (
    id: string,
    comment: string,
    toRecipients: ReadonlyArray<{ emailAddress: { address: string } }>,
    prefix: string = "/me",
  ) =>
    request<GraphMessage>("POST", graphPath`${raw(prefix)}/messages/${id}/createForward`, {
      body: { comment, toRecipients },
    })

  // Immediate-send reply actions: thread + quote, then send in one step.
  const sendReplyAll = (id: string, comment: string, prefix: string = "/me") =>
    request<Record<string, never>>("POST", graphPath`${raw(prefix)}/messages/${id}/replyAll`, { body: { comment } })

  const sendForward = (
    id: string,
    comment: string,
    toRecipients: ReadonlyArray<{ emailAddress: { address: string } }>,
    prefix: string = "/me",
  ) =>
    request<Record<string, never>>("POST", graphPath`${raw(prefix)}/messages/${id}/forward`, {
      body: { comment, toRecipients },
    })

  const searchMessages = (query: string, odataParams?: ODataParams, prefix: string = "/me") =>
    request<ODataResponse<GraphMessage>>("GET", graphPath`${raw(prefix)}/messages`, {
      odataParams: { ...odataParams, $search: query },
    })

  // Calendar
  // Without a calendar ID, events go to the default calendar. With one, the path goes through
  // /me/calendars/{id}, which is the only way to reach a calendar someone else shared with you:
  // /me/events/{id} does not find their events.
  const eventsPath = (calendarId?: string) =>
    calendarId ? graphPath`/me/calendars/${calendarId}/events` : graphPath`/me/events`

  // Graph answers event times in UTC unless Prefer: outlook.timezone names a zone, IANA or Windows
  // (both checked live, 2026-11-07, together with the body format in one comma-separated header).
  // An unknown zone is refused with "A valid TimeZone value must be specified".
  const eventView = (timeZone?: string, bodyContentType?: "text" | "html") => {
    const prefer = [
      timeZone === undefined ? undefined : `outlook.timezone="${timeZone}"`,
      bodyContentType === undefined ? undefined : `outlook.body-content-type="${bodyContentType}"`,
    ].filter((p): p is string => p !== undefined)
    return prefer.length === 0 ? {} : { headers: { Prefer: prefer.join(", ") } }
  }

  const listEvents = (odataParams?: ODataParams, calendarId?: string, timeZone?: string) =>
    request<ODataResponse<GraphEvent>>("GET", eventsPath(calendarId), { odataParams, ...eventView(timeZone) })

  const listAllEvents = (odataParams?: ODataParams, calendarId?: string, timeZone?: string) =>
    requestPaginated<GraphEvent>(eventsPath(calendarId), { odataParams, ...eventView(timeZone) })

  // calendarView expands recurring series into individual instances within [start, end].
  // Required by Graph: startDateTime/endDateTime as query params on the path.
  const listCalendarView = (
    startDateTime: string,
    endDateTime: string,
    odataParams?: ODataParams,
    calendarId?: string,
    timeZone?: string,
  ) => {
    const base = calendarId ? graphPath`/me/calendars/${calendarId}/calendarView` : graphPath`/me/calendarView`
    const query = `?startDateTime=${encodeURIComponent(startDateTime)}&endDateTime=${encodeURIComponent(endDateTime)}`
    const path = graphPath`${base}${raw(query)}`
    return request<ODataResponse<GraphEvent>>("GET", path, { odataParams, ...eventView(timeZone) })
  }

  const getEvent = (id: string, bodyContentType?: "text" | "html", calendarId?: string, timeZone?: string) =>
    request<GraphEvent>("GET", graphPath`${eventsPath(calendarId)}/${id}`, eventView(timeZone, bodyContentType))

  const getCalendar = (calendarId: string) =>
    request<GraphCalendar>("GET", graphPath`/me/calendars/${calendarId}`, {
      odataParams: { $select: ["name", "canEdit"] },
    })

  const createEvent = (event: Record<string, unknown>, calendarId?: string, timeZone?: string) =>
    request<GraphEvent>("POST", eventsPath(calendarId), { body: event, ...eventView(timeZone) })

  const updateEvent = (id: string, event: Record<string, unknown>, calendarId?: string, timeZone?: string) =>
    request<GraphEvent>("PATCH", graphPath`${eventsPath(calendarId)}/${id}`, { body: event, ...eventView(timeZone) })

  const deleteEvent = (id: string, calendarId?: string) =>
    request<Record<string, never>>("DELETE", graphPath`${eventsPath(calendarId)}/${id}`)

  // accept, decline and tentativelyAccept take the same body; Graph answers 202 with no content.
  const respondToEvent = (
    id: string,
    action: "accept" | "decline" | "tentativelyAccept",
    body: { comment?: string; sendResponse: boolean },
  ) => request<Record<string, never>>("POST", graphPath`/me/events/${id}/${raw(action)}`, { body })

  // findMeetingTimes suggests slots where attendees are free. POST body; Prefer header forces
  // UTC in the response so callers don't get mailbox-local times back unexpectedly.
  const findMeetingTimes = (body: Record<string, unknown>) =>
    request<GraphMeetingTimeSuggestionsResult>("POST", "/me/findMeetingTimes", {
      body,
      headers: { Prefer: 'outlook.timezone="UTC"' },
    })

  // Contacts
  const listContacts = (odataParams?: ODataParams) =>
    request<ODataResponse<GraphContact>>("GET", "/me/contacts", { odataParams })

  const getContact = (id: string) => request<GraphContact>("GET", graphPath`/me/contacts/${id}`)

  const createContact = (contact: Record<string, unknown>) =>
    request<GraphContact>("POST", "/me/contacts", { body: contact })

  const updateContact = (id: string, contact: Record<string, unknown>) =>
    request<GraphContact>("PATCH", graphPath`/me/contacts/${id}`, { body: contact })

  const deleteContact = (id: string) => request<Record<string, never>>("DELETE", graphPath`/me/contacts/${id}`)

  const searchContacts = (query: string, odataParams?: ODataParams) =>
    request<ODataResponse<GraphContact>>("GET", "/me/contacts", {
      odataParams: { ...odataParams, $search: query },
    })

  // Files (OneDrive)
  const listDriveItems = (folderId?: string) =>
    request<ODataResponse<GraphDriveItem>>(
      "GET",
      folderId ? graphPath`/me/drive/items/${folderId}/children` : "/me/drive/root/children",
    )

  const listDriveItemsByPath = (folderPath: string) =>
    request<ODataResponse<GraphDriveItem>>("GET", graphPath`/me/drive/root:/${drivePath(folderPath)}:/children`)

  const getDriveItem = (id: string) => request<GraphDriveItem>("GET", graphPath`/me/drive/items/${id}`)

  // q is an OData string literal: an apostrophe ends it unless doubled, so "O'Brien" broke the request.
  const searchLiteral = (query: string) => encodeURIComponent(query.replaceAll("'", "''"))

  const filesSearchPath = (query: string) => `/me/drive/root/search(q='${searchLiteral(query)}')`

  const searchFiles = (query: string, odataParams?: ODataParams) =>
    request<ODataResponse<GraphDriveItem>>("GET", filesSearchPath(query), { odataParams })

  const searchAllFiles = (query: string) => requestPaginated<GraphDriveItem>(filesSearchPath(query))

  // A drive item lives under /me/drive only when it is in the caller's own OneDrive. SharePoint
  // items need their own drive, which is why the /me/drive-only version returned "The resource
  // could not be found." for every SharePoint item ID.
  const driveItemPath = (id: string, driveId?: string) =>
    driveId ? graphPath`/drives/${driveId}/items/${id}` : graphPath`/me/drive/items/${id}`

  const downloadFile = (id: string, driveId?: string) => request<GraphDriveItem>("GET", driveItemPath(id, driveId))

  const downloadFileContent = async (id: string, driveId?: string): Promise<Either<GraphApiError, Uint8Array>> => {
    const tokenResult = await auth.getAccessToken()
    if (tokenResult.isLeft()) {
      return Left<GraphApiError, Uint8Array>({
        type: "auth",
        message: (tokenResult.value as { message: string }).message,
      })
    }
    const token = tokenResult.value as string
    const path = graphPath`${driveItemPath(id, driveId)}/content`
    if (path.isLeft()) return Left<GraphApiError, Uint8Array>(path.value as GraphApiError)
    const url = `${GRAPH_API_BASE}/${defaultVersion()}${path.orThrow()}`

    try {
      const response = await resilientFetch(url, {
        method: "GET",
        headers: { Authorization: `Bearer ${token}` },
        redirect: "follow",
      })
      if (!response.ok) {
        return mapHttpError<Uint8Array>(response)
      }
      // Bytes, not text: download_file has to sniff a file OneDrive typed as octet-stream, and a
      // decoded string has already replaced the evidence (NUL bytes, invalid UTF-8).
      return Right<GraphApiError, Uint8Array>(new Uint8Array(await response.arrayBuffer()))
    } catch (error) {
      const notSent = notSentError(error)
      if (notSent) return Left<GraphApiError, Uint8Array>(notSent)
      return Left<GraphApiError, Uint8Array>({
        type: "network",
        message: `Network error: ${error instanceof Error ? error.message : String(error)}`,
      })
    }
  }

  const createFolder = (parentId: string, name: string) =>
    request<GraphDriveItem>("POST", graphPath`/me/drive/items/${parentId}/children`, {
      body: { name, folder: {}, "@microsoft.graph.conflictBehavior": "rename" },
    })

  const ifMatchHeader = (ifMatch?: string) => (ifMatch === undefined ? {} : { headers: { "If-Match": ifMatch } })

  // A drive's root has a real ID, and a move or copy needs it: Graph refuses "root" in parentReference.
  // Its parentReference carries the drive's own ID, which a copy's parentReference should name.
  const getDriveRoot = (driveId?: string) =>
    request<GraphDriveItem>("GET", driveId ? graphPath`/drives/${driveId}/root` : "/me/drive/root", {
      odataParams: { $select: ["id", "parentReference"] },
    })

  // Move, rename, or both: one PATCH. Graph can't move an item to another drive this way.
  const updateDriveItem = (id: string, changes: Record<string, unknown>, driveId?: string, ifMatch?: string) =>
    request<GraphDriveItem>("PATCH", driveItemPath(id, driveId), { body: changes, ...ifMatchHeader(ifMatch) })

  // Graph answers 202 with an empty body and copies in the background. The conflict setting is a query
  // parameter, sent only when asked for: Graph's default is fail, and OneDrive personal doesn't support it.
  const copyDriveItem = (id: string, body: Record<string, unknown>, driveId?: string, conflictBehavior?: "rename") =>
    request<Record<string, never>>(
      "POST",
      graphPath`${driveItemPath(id, driveId)}/copy${raw(conflictBehavior ? `?@microsoft.graph.conflictBehavior=${conflictBehavior}` : "")}`,
      { body },
    )

  // Moves the item, and a folder's contents, to the recycle bin.
  const deleteDriveItem = (id: string, driveId?: string, ifMatch?: string) =>
    request<Record<string, never>>("DELETE", driveItemPath(id, driveId), ifMatchHeader(ifMatch))

  // SharePoint Sites
  const listFollowedSites = () => request<ODataResponse<GraphSite>>("GET", "/me/followedSites")

  const searchSites = (query: string) =>
    request<ODataResponse<GraphSite>>("GET", `/sites?search=${encodeURIComponent(query)}`)

  const getSite = (siteId: string) => request<GraphSite>("GET", graphPath`/sites/${siteId}`)

  const listSiteDrives = (siteId: string) =>
    request<ODataResponse<GraphDrive>>("GET", graphPath`/sites/${siteId}/drives`)

  const listSiteDriveItems = (siteId: string, driveId?: string, folderId?: string) => {
    if (driveId && folderId) {
      return request<ODataResponse<GraphDriveItem>>(
        "GET",
        graphPath`/sites/${siteId}/drives/${driveId}/items/${folderId}/children`,
      )
    }
    if (folderId) {
      return request<ODataResponse<GraphDriveItem>>("GET", graphPath`/sites/${siteId}/drive/items/${folderId}/children`)
    }
    if (driveId) {
      return request<ODataResponse<GraphDriveItem>>("GET", graphPath`/sites/${siteId}/drives/${driveId}/root/children`)
    }
    return request<ODataResponse<GraphDriveItem>>("GET", graphPath`/sites/${siteId}/drive/root/children`)
  }

  const listSiteDriveItemsByPath = (siteId: string, path: string, driveId?: string) => {
    if (driveId) {
      return request<ODataResponse<GraphDriveItem>>(
        "GET",
        graphPath`/sites/${siteId}/drives/${driveId}/root:/${drivePath(path)}:/children`,
      )
    }
    return request<ODataResponse<GraphDriveItem>>(
      "GET",
      graphPath`/sites/${siteId}/drive/root:/${drivePath(path)}:/children`,
    )
  }

  const siteDrivePath = (siteId: string, driveId?: string) =>
    driveId ? graphPath`/sites/${siteId}/drives/${driveId}` : graphPath`/sites/${siteId}/drive`

  const siteSearchPath = (siteId: string, query: string, driveId?: string) =>
    graphPath`${siteDrivePath(siteId, driveId)}/root/search(q='${raw(searchLiteral(query))}')`

  const searchSiteFiles = (siteId: string, query: string, driveId?: string, odataParams?: ODataParams) =>
    request<ODataResponse<GraphDriveItem>>("GET", siteSearchPath(siteId, query, driveId), { odataParams })

  const searchAllSiteFiles = (siteId: string, query: string, driveId?: string) =>
    requestPaginated<GraphDriveItem>(siteSearchPath(siteId, query, driveId))

  // Teams
  const listTeams = () =>
    request<ODataResponse<{ id: string; displayName?: string; description?: string }>>("GET", "/me/joinedTeams")

  const listChannels = (teamId: string) =>
    request<ODataResponse<GraphChannel>>("GET", graphPath`/teams/${teamId}/channels`)

  const listChannelMessages = (teamId: string, channelId: string, odataParams?: ODataParams) =>
    request<ODataResponse<GraphChannelMessage>>("GET", graphPath`/teams/${teamId}/channels/${channelId}/messages`, {
      odataParams,
    })

  // A message body is { contentType: "text" | "html", content }; see utils/teams-markdown.ts.
  type MessageBody = { readonly contentType: "text" | "html"; readonly content: string }

  const sendChannelMessage = (teamId: string, channelId: string, body: MessageBody) =>
    request<GraphChannelMessage>("POST", graphPath`/teams/${teamId}/channels/${channelId}/messages`, { body: { body } })

  // Graph only accepts replies to a top-level channel message, not to another reply.
  const replyToChannelMessage = (teamId: string, channelId: string, messageId: string, body: MessageBody) =>
    request<GraphChannelMessage>(
      "POST",
      graphPath`/teams/${teamId}/channels/${channelId}/messages/${messageId}/replies`,
      {
        body: { body },
      },
    )

  // Chats
  const listChats = (odataParams?: ODataParams) =>
    request<ODataResponse<GraphChat>>("GET", "/me/chats", { odataParams })

  const listChatMessages = (chatId: string, odataParams?: ODataParams) =>
    request<ODataResponse<GraphChatMessage>>("GET", graphPath`/chats/${chatId}/messages`, { odataParams })

  const sendChatMessage = (chatId: string, body: MessageBody) =>
    request<GraphChatMessage>("POST", graphPath`/chats/${chatId}/messages`, { body: { body } })

  // Graph returns the existing chat for a one-on-one pair, but creates a new group chat every call.
  const createChat = (chat: Record<string, unknown>) => request<GraphChat>("POST", "/chats", { body: chat })

  // Users & Groups
  const getMe = () => request<GraphUser>("GET", "/me")

  const listUsers = (odataParams?: ODataParams) => request<ODataResponse<GraphUser>>("GET", "/users", { odataParams })

  const getUser = (id: string) => request<GraphUser>("GET", graphPath`/users/${id}`)

  // /users/{key} accepts an ID or a user principal name, not an email address, so an address that
  // differs from the UPN needs the mail filter below. A guest UPN contains "#EXT#", and "#" would
  // otherwise end the path.
  const CHAT_MEMBER_FIELDS = ["id", "displayName", "mail", "userPrincipalName", "userType"]

  const getUserByKey = (key: string) =>
    request<GraphUser>("GET", graphPath`/users/${key}`, { odataParams: { $select: CHAT_MEMBER_FIELDS } })

  // userType is only returned when selected, and the caller needs it too: an in-tenant guest must
  // join a chat as a guest, the caller included.
  const getMeAsChatMember = () => request<GraphUser>("GET", "/me", { odataParams: { $select: CHAT_MEMBER_FIELDS } })

  const findUsersByMail = (address: string) =>
    request<ODataResponse<GraphUser>>("GET", "/users", {
      odataParams: { $filter: `mail eq '${address.replace(/'/g, "''")}'`, $select: CHAT_MEMBER_FIELDS },
    })

  const listGroups = (odataParams?: ODataParams) =>
    request<ODataResponse<GraphGroup>>("GET", "/groups", { odataParams })

  const getGroup = (id: string) => request<GraphGroup>("GET", graphPath`/groups/${id}`)

  const listGroupMembers = (id: string) => request<ODataResponse<GraphUser>>("GET", graphPath`/groups/${id}/members`)

  // Planner
  const listPlans = () => request<ODataResponse<GraphPlan>>("GET", "/me/planner/plans")

  // The user's group memberships (groups only, filtered from the memberOf directory objects) and a
  // group's plans — used to fan out plan discovery, since /me/planner/plans omits group-owned plans
  // the user wasn't explicitly added to.
  const listMyGroups = () =>
    request<ODataResponse<GraphGroup>>("GET", "/me/memberOf/microsoft.graph.group?$select=id,displayName")

  const listGroupPlans = (groupId: string) =>
    request<ODataResponse<GraphPlan>>("GET", graphPath`/groups/${groupId}/planner/plans`)

  const listPlannerBuckets = (planId: string) =>
    request<ODataResponse<GraphBucket>>("GET", graphPath`/planner/plans/${planId}/buckets`)

  const createPlannerBucket = (bucket: Record<string, unknown>) =>
    request<GraphBucket>("POST", "/planner/buckets", { body: bucket })

  const listPlannerTasks = (planId: string) =>
    request<ODataResponse<GraphPlannerTask>>("GET", graphPath`/planner/plans/${planId}/tasks`)

  const getPlannerTask = (id: string) => request<GraphPlannerTask>("GET", graphPath`/planner/tasks/${id}`)

  const createPlannerTask = (task: Record<string, unknown>) =>
    request<GraphPlannerTask>("POST", "/planner/tasks", { body: task })

  const updatePlannerTask = (id: string, task: Record<string, unknown>, etag: string) =>
    request<GraphPlannerTask>("PATCH", graphPath`/planner/tasks/${id}`, {
      body: task,
      headers: { "If-Match": etag },
    })

  // Task details (description / checklist / references) is a separate object with its own ETag,
  // so its writes need a distinct If-Match from the parent task's.
  const getPlannerTaskDetails = (id: string) =>
    request<{ "@odata.etag": string } & Record<string, unknown>>("GET", graphPath`/planner/tasks/${id}/details`)

  const updatePlannerTaskDetails = (id: string, details: Record<string, unknown>, etag: string) =>
    request<Record<string, unknown>>("PATCH", graphPath`/planner/tasks/${id}/details`, {
      body: details,
      headers: { "If-Match": etag },
    })

  // OneNote
  const listOnenoteNotebooks = () => request<ODataResponse<GraphNotebook>>("GET", "/me/onenote/notebooks")

  const listOnenoteSections = (notebookId: string) =>
    request<ODataResponse<GraphSection>>("GET", graphPath`/me/onenote/notebooks/${notebookId}/sections`)

  const listOnenotePages = (sectionId: string) =>
    request<ODataResponse<GraphPage>>("GET", graphPath`/me/onenote/sections/${sectionId}/pages`)

  // The /content endpoint returns text/html, not JSON — read it as raw text.
  const getOnenotePageContent = (pageId: string) =>
    request<string>("GET", graphPath`/me/onenote/pages/${pageId}/content`, { responseType: "text" })

  // OneNote writes. createOnenotePage sends raw text/html; the rest are JSON.
  const createOnenotePage = (sectionId: string, html: string) =>
    request<GraphPage>("POST", graphPath`/me/onenote/sections/${sectionId}/pages`, {
      body: html,
      contentType: "text/html",
    })

  const updateOnenotePageContent = (pageId: string, commands: readonly unknown[]) =>
    request<Record<string, never>>("PATCH", graphPath`/me/onenote/pages/${pageId}/content`, { body: commands })

  const createOnenoteSection = (notebookId: string, displayName: string) =>
    request<GraphSection>("POST", graphPath`/me/onenote/notebooks/${notebookId}/sections`, { body: { displayName } })

  const createOnenoteNotebook = (displayName: string) =>
    request<GraphNotebook>("POST", "/me/onenote/notebooks", { body: { displayName } })

  const copyOnenotePage = (pageId: string, sectionId: string) =>
    request<Record<string, never>>("POST", graphPath`/me/onenote/pages/${pageId}/copyToSection`, {
      body: { id: sectionId },
    })

  const deleteOnenotePage = (pageId: string) =>
    request<Record<string, never>>("DELETE", graphPath`/me/onenote/pages/${pageId}`)

  // To Do
  const listTodoLists = () => request<ODataResponse<GraphTodoList>>("GET", "/me/todo/lists")

  const createTodoList = (displayName: string) =>
    request<GraphTodoList>("POST", "/me/todo/lists", { body: { displayName } })

  const getTodoList = (listId: string) => request<GraphTodoList>("GET", graphPath`/me/todo/lists/${listId}`)

  const deleteTodoList = (listId: string) =>
    request<Record<string, never>>("DELETE", graphPath`/me/todo/lists/${listId}`)

  // Graph refuses $select on To Do task lists ("Invalid request", checked live 2026-10-30, while
  // $select on messages works), so these reads can't follow the AGENTS.md $select rule.
  // linkedResources is a relationship, so it has to be expanded to be read.
  const todoTasksPath = (listId: string) => graphPath`/me/todo/lists/${listId}/tasks`

  const listTodoTasks = (listId: string, odataParams?: ODataParams) =>
    request<ODataResponse<GraphTodoTask>>("GET", todoTasksPath(listId), { odataParams })

  // A continuation link is followed only if it is Graph's own, carries no credentials or fragment,
  // and stays on the same list's tasks: a link that points anywhere else is refused, not followed.
  // It must also keep the first page's $expand: a page without linkedResources would read as a
  // page of untagged tasks, and a duplicate check would miss the task it was looking for.
  const nextTodoTaskPage = (nextLink: string, listId: string, expand: string | null) => {
    const url = URL.canParse(nextLink) ? new URL(nextLink) : undefined
    const match = /^\/(v1\.0|beta)\/me\/todo\/lists\/([^/]+)\/tasks$/.exec(url?.pathname ?? "")
    if (
      url?.origin !== GRAPH_API_BASE ||
      url.username ||
      url.password ||
      url.hash ||
      !match ||
      safeDecode(match[2]!) !== safeDecode(listId) ||
      url.searchParams.get("$expand") !== expand
    ) {
      return Promise.resolve(
        Left<GraphApiError, ODataResponse<GraphTodoTask>>({
          type: "api",
          message: "Invalid Graph To Do continuation URL.",
        }),
      )
    }
    return request<ODataResponse<GraphTodoTask>>("GET", graphPath`${todoTasksPath(listId)}${raw(url.search)}`, {
      version: match[1] as GraphApiVersion,
    })
  }

  // Every page of a list, through the checked link above. Hitting the cap is an error, not a
  // shorter list: a caller checking for an existing task must never mistake "not read" for "absent".
  const listAllTodoTasks = async (
    listId: string,
    odataParams?: ODataParams,
  ): Promise<Either<GraphApiError, ReadonlyArray<GraphTodoTask>>> => {
    const expand = odataParams?.$expand ? odataParams.$expand.join(",") : null
    const step = async (
      page: Either<GraphApiError, ODataResponse<GraphTodoTask>>,
      read: ReadonlyArray<GraphTodoTask>,
      pages: number,
    ): Promise<Either<GraphApiError, ReadonlyArray<GraphTodoTask>>> => {
      if (page.isLeft()) return Left(page.value as GraphApiError)
      const { value, "@odata.nextLink": next } = page.value as Partial<ODataResponse<GraphTodoTask>>
      if (!Array.isArray(value)) return Left({ type: "parse", message: "A page of To Do tasks had no value array." })
      const all = [...read, ...value]
      if (next === undefined) return Right(all)
      if (pages >= MAX_TODO_TASK_PAGES) {
        return Left({
          type: "api",
          message: `The list has more than ${MAX_TODO_TASK_PAGES} pages of ${TODO_TASK_PAGE_SIZE} tasks; reading stopped there.`,
        })
      }
      return step(await nextTodoTaskPage(next, listId, expand), all, pages + 1)
    }
    return step(await listTodoTasks(listId, { ...odataParams, $top: TODO_TASK_PAGE_SIZE }), [], 1)
  }

  // What changed in a list since a token, through Graph's delta query. Checked live 2026-10-30:
  // delta refuses $select and the filtered extension expand, so it is sent with no query options at
  // all; its continuation links carry only $skiptoken, and the last page's deltaLink only
  // $deltatoken. A token is checked and encoded before it goes into a URL, and every link Graph
  // returns is checked before it is followed. Hitting the page cap or a bad link returns no token,
  // so a caller never moves past changes it didn't read.
  const todoDeltaPath = (listId: string) => graphPath`${todoTasksPath(listId)}/delta`
  const deltaPrefer = { headers: { Prefer: `odata.maxpagesize=${TODO_TASK_PAGE_SIZE}` } }

  const deltaLinkToken = (link: string, listId: string, key: "$skiptoken" | "$deltatoken") => {
    const url = URL.canParse(link) ? new URL(link) : undefined
    const match = /^\/(v1\.0|beta)\/me\/todo\/lists\/([^/]+)\/tasks\/delta$/.exec(url?.pathname ?? "")
    const keys = url ? [...url.searchParams.keys()] : []
    const token = url?.searchParams.get(key) ?? ""
    const ok =
      url?.origin === GRAPH_API_BASE &&
      !url.username &&
      !url.password &&
      !url.hash &&
      match !== null &&
      safeDecode(match[2]!) === safeDecode(listId) &&
      keys.length === 1 &&
      keys[0] === key &&
      isDeltaToken(token)
    return ok ? { token, version: match[1] as GraphApiVersion } : undefined
  }

  const todoTaskDelta = async (
    listId: string,
    deltaToken?: string,
  ): Promise<Either<GraphApiError, TodoDeltaResult>> => {
    if (deltaToken !== undefined && !isDeltaToken(deltaToken)) {
      return Left({ type: "api", status: 400, message: "Badly formed token." })
    }
    const page = (query: string, version?: GraphApiVersion) =>
      request<TodoDeltaPage>("GET", graphPath`${todoDeltaPath(listId)}${raw(query)}`, { ...deltaPrefer, version })
    const step = async (
      response: Either<GraphApiError, TodoDeltaPage>,
      read: ReadonlyArray<TodoDeltaRecord>,
      pages: number,
    ): Promise<Either<GraphApiError, TodoDeltaResult>> => {
      if (response.isLeft()) return Left(response.value as GraphApiError)
      const body = response.value as Partial<TodoDeltaPage>
      if (!Array.isArray(body.value))
        return Left({ type: "parse", message: "A page of To Do changes had no value array." })
      const records = [...read, ...body.value]
      const next = body["@odata.nextLink"]
      const done = body["@odata.deltaLink"]
      if (next === undefined && done === undefined) {
        return Left({ type: "parse", message: "Graph returned neither a next page nor a delta token." })
      }
      if (next === undefined) {
        const last = deltaLinkToken(done!, listId, "$deltatoken")
        return last === undefined
          ? Left({ type: "api", message: "Invalid Graph To Do delta link." })
          : Right({ records, deltaToken: last.token })
      }
      if (pages >= MAX_TODO_TASK_PAGES) {
        return Left({
          type: "api",
          message: `The changes run past ${MAX_TODO_TASK_PAGES} pages of ${TODO_TASK_PAGE_SIZE} tasks; reading stopped there.`,
        })
      }
      const skip = deltaLinkToken(next, listId, "$skiptoken")
      if (skip === undefined) return Left({ type: "api", message: "Invalid Graph To Do continuation URL." })
      return step(await page(`?$skiptoken=${encodeURIComponent(skip.token)}`, skip.version), records, pages + 1)
    }
    const first = deltaToken === undefined ? "" : `?$deltatoken=${encodeURIComponent(deltaToken)}`
    return step(await page(first), [], 1)
  }

  const createTodoTask = (listId: string, task: Record<string, unknown>) =>
    request<GraphTodoTask>("POST", graphPath`/me/todo/lists/${listId}/tasks`, { body: task })

  const updateTodoTask = (listId: string, taskId: string, task: Record<string, unknown>) =>
    request<GraphTodoTask>("PATCH", graphPath`/me/todo/lists/${listId}/tasks/${taskId}`, { body: task })

  const getTodoTask = (listId: string, taskId: string, odataParams?: ODataParams) =>
    request<GraphTodoTask>("GET", graphPath`/me/todo/lists/${listId}/tasks/${taskId}`, { odataParams })

  // A task's open extensions. A PATCH replaces the whole extension (checked live 2026-10-30), so a
  // caller that means to change some fields reads, merges, and sends them all.
  const extensionPath = (listId: string, taskId: string) => graphPath`${todoTasksPath(listId)}/${taskId}/extensions`

  const getTodoTaskExtension = (listId: string, taskId: string, name: string) =>
    request<Record<string, unknown>>("GET", graphPath`${extensionPath(listId, taskId)}/${encodedId(name)}`)

  const createTodoTaskExtension = (listId: string, taskId: string, body: Record<string, unknown>) =>
    request<Record<string, unknown>>("POST", extensionPath(listId, taskId), { body })

  const replaceTodoTaskExtension = (listId: string, taskId: string, name: string, body: Record<string, unknown>) =>
    request<Record<string, unknown>>("PATCH", graphPath`${extensionPath(listId, taskId)}/${encodedId(name)}`, {
      body,
    })

  const deleteTodoTask = (listId: string, taskId: string) =>
    request<Record<string, never>>("DELETE", graphPath`/me/todo/lists/${listId}/tasks/${taskId}`)

  // Text-only file upload. Binary uploads must use get_upload_config (httpStream) or upload_file_from_path (stdio).
  // With ifMatch the write is conditional: it goes through core's upload session, the route Graph
  // documents If-Match on, and lands only if the file still has that eTag.
  const uploadFile = async (
    path: string,
    content: string,
    contentType: string = "text/plain",
    conflictBehavior: "rename" | "replace" | "fail" = "rename",
    ifMatch?: string,
  ): Promise<Either<GraphApiError, GraphDriveItem>> => {
    const tokenResult = await auth.getAccessToken()
    if (tokenResult.isLeft()) {
      return Left<GraphApiError, GraphDriveItem>({
        type: "auth",
        message: (tokenResult.value as { message: string }).message,
      })
    }

    // Every upload goes through core, which checks the path, eTag, content type and conflict
    // behavior. Up to 4 MB, and without ifMatch, that is one PUT; upload_file caps text at 4 MB.
    const buffer = Buffer.from(content, "utf8")
    return uploadDriveContent(
      `${GRAPH_API_BASE}/${defaultVersion()}`,
      path,
      tokenResult.value as string,
      buffer,
      contentType,
      conflictBehavior,
      { ifMatch },
    )
  }

  // Generic escape hatch. Optional headers unlock header-gated ops (If-Match for Planner
  // task-details writes, Prefer, ConsistencyLevel, ...) that the typed tools don't cover.
  const graphQuery = <T = unknown>(
    method: string,
    path: string,
    body?: Record<string, unknown>,
    version?: GraphApiVersion,
    headers?: Record<string, string>,
  ) => request<T>(method, path, { body, version, headers })

  return Object.freeze({
    // Core
    request,
    requestPaginated,
    // Mail
    listMessages,
    listAllMessages,
    nextMessageScanPage,
    getMessage,
    listMailFolders,
    moveMessage,
    batchRequest,

    listAttachments,
    sendMessage,
    createDraft,
    sendDraft,
    getDraftState,
    updateMessage,
    deleteMessage,
    getAutomaticReplies,
    setAutomaticReplies,
    listMasterCategories,
    createMasterCategory,
    getMessageCategories,
    getMailFolder,
    listMessageRules,
    createMessageRule,
    deleteMessageRule,
    sendReply,
    sendReplyAll,
    sendForward,
    createReplyDraft,
    createReplyAllDraft,
    createForwardDraft,
    searchMessages,
    // Calendar
    listEvents,
    listAllEvents,
    listCalendarView,
    getEvent,
    getCalendar,
    createEvent,
    updateEvent,
    deleteEvent,
    respondToEvent,
    findMeetingTimes,
    // Contacts
    listContacts,
    getContact,
    createContact,
    updateContact,
    deleteContact,
    searchContacts,
    // Files
    listDriveItems,
    listDriveItemsByPath,
    getDriveItem,
    searchFiles,
    searchAllFiles,
    downloadFile,
    downloadFileContent,
    createFolder,
    getDriveRoot,
    updateDriveItem,
    copyDriveItem,
    deleteDriveItem,
    // SharePoint
    listFollowedSites,
    searchSites,
    getSite,
    listSiteDrives,
    listSiteDriveItems,
    listSiteDriveItemsByPath,
    searchSiteFiles,
    searchAllSiteFiles,
    // Chats
    listChats,
    listChatMessages,
    sendChatMessage,
    createChat,
    // Teams
    listTeams,
    listChannels,
    listChannelMessages,
    sendChannelMessage,
    replyToChannelMessage,
    // Users & Groups
    getMe,
    listUsers,
    getUser,
    getUserByKey,
    getMeAsChatMember,
    findUsersByMail,
    listGroups,
    getGroup,
    listGroupMembers,
    // Planner
    listPlans,
    listMyGroups,
    listGroupPlans,
    listPlannerBuckets,
    createPlannerBucket,
    listPlannerTasks,
    getPlannerTask,
    createPlannerTask,
    updatePlannerTask,
    getPlannerTaskDetails,
    updatePlannerTaskDetails,
    // OneNote
    listOnenoteNotebooks,
    listOnenoteSections,
    listOnenotePages,
    getOnenotePageContent,
    createOnenotePage,
    updateOnenotePageContent,
    createOnenoteSection,
    createOnenoteNotebook,
    copyOnenotePage,
    deleteOnenotePage,
    // To Do
    listTodoLists,
    createTodoList,
    getTodoList,
    deleteTodoList,
    listTodoTasks,
    listAllTodoTasks,
    todoTaskDelta,
    getTodoTaskExtension,
    createTodoTaskExtension,
    replaceTodoTaskExtension,
    createTodoTask,
    updateTodoTask,
    getTodoTask,
    deleteTodoTask,
    // Upload
    uploadFile,
    // Generic
    graphQuery,
  })
}

export type GraphClient = ReturnType<typeof createGraphClient>

const clientRef = Ref<Option<GraphClient>>(None())

export const initializeGraphClient = (auth: AuthStrategy): GraphClient => {
  const c = createGraphClient(auth)
  clientRef.set(Some(c))
  return c
}

export const getGraphClient = (): Option<GraphClient> => clientRef.get()
