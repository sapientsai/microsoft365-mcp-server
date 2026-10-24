import {
  type AuthStrategy,
  createGraphRequest,
  GRAPH_API_BASE,
  mapHttpError,
  parseJsonResponse,
  resilientFetch,
} from "@sapientsai/ms-graph-core"
import { None, type Option, Ref, Some } from "functype"
import { type Either, Left, Right } from "functype/either"

import type {
  GraphApiError,
  GraphApiVersion,
  GraphAttachment,
  GraphBucket,
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
  GraphNotebook,
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
import { defaultVersion } from "./graph-version"

const createGraphClient = (auth: AuthStrategy) => {
  // Generic request machinery (auth injection, OData, error mapping, pagination) lives in
  // @sapientsai/ms-graph-core; this client builds the M365 domain methods on top of it.
  const { request, requestPaginated } = createGraphRequest(auth, { defaultVersion })

  // Mail
  // /me/messages spans every folder, Sent Items included. A folder takes a well-known name
  // (inbox, sentitems, ...) or a folder ID; Graph accepts either in the same path segment.
  const messagesPath = (folder?: string) =>
    folder ? `/me/mailFolders/${encodeURIComponent(folder)}/messages` : "/me/messages"

  const listMessages = (odataParams?: ODataParams, folder?: string) =>
    request<ODataResponse<GraphMessage>>("GET", messagesPath(folder), { odataParams })

  const listAllMessages = (odataParams?: ODataParams, folder?: string) =>
    requestPaginated<GraphMessage>(messagesPath(folder), { odataParams })

  // Prefer: outlook.body-content-type="text" makes Graph convert the body server-side.
  // Marketing mail is mostly CSS and layout tables — one newsletter measured 79,347
  // characters as HTML — so for a caller that only needs the words this is a ~95%
  // reduction, and better than stripping tags locally. Events honour the same header.
  const bodyFormat = (bodyContentType?: "text" | "html") =>
    bodyContentType ? { headers: { Prefer: `outlook.body-content-type="${bodyContentType}"` } } : undefined

  const getMessage = (id: string, bodyContentType?: "text" | "html") =>
    request<GraphMessage>("GET", `/me/messages/${id}`, bodyFormat(bodyContentType))

  const listMailFolders = (odataParams?: ODataParams) =>
    request<ODataResponse<GraphMailFolder>>("GET", "/me/mailFolders", { odataParams })

  // Only follow Graph-issued mail collection links; never send a bearer token to another host.
  // Graph writes a folder-scoped nextLink in OData key form, /me/mailFolders('<id>')/messages, not
  // the /mailFolders/<id>/ form the first request used (confirmed live; 1.2.14 rejected it).
  const nextMessageScanPage = (nextLink: string) => {
    const url = URL.canParse(nextLink) ? new URL(nextLink) : undefined
    const match = /^\/(v1\.0|beta)(\/me\/(?:mailFolders(?:\/[^/]+|\('[^'/]+'\))\/)?messages)$/.exec(url?.pathname ?? "")
    if (url?.origin !== GRAPH_API_BASE || url.username || url.password || url.hash || !match) {
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

  const moveMessage = (id: string, destinationId: string) =>
    request<GraphMessage>("POST", `/me/messages/${id}/move`, { body: { destinationId } })

  // $select omits contentBytes deliberately: fileAttachment includes the full base64 payload
  // by default, which would drag megabytes of binary through the model for a listing.
  const listAttachments = (messageId: string) =>
    request<ODataResponse<GraphAttachment>>("GET", `/me/messages/${messageId}/attachments`, {
      odataParams: { $select: ["id", "name", "contentType", "size", "isInline", "lastModifiedDateTime"] },
    })

  const sendMessage = (message: Record<string, unknown>) =>
    request<Record<string, never>>("POST", "/me/sendMail", { body: message })

  const createDraft = (message: Record<string, unknown>) =>
    request<GraphMessage>("POST", "/me/messages", { body: message })

  const sendDraft = (messageId: string) => request<Record<string, never>>("POST", `/me/messages/${messageId}/send`)

  // Draft edits and deletes check isDraft first, so this fetches only what that check needs.
  const getDraftState = (id: string) =>
    request<GraphMessage>("GET", `/me/messages/${id}`, { odataParams: { $select: ["id", "isDraft", "subject"] } })

  const updateMessage = (id: string, changes: Record<string, unknown>) =>
    request<GraphMessage>("PATCH", `/me/messages/${id}`, { body: changes })

  const deleteMessage = (id: string) => request<Record<string, never>>("DELETE", `/me/messages/${id}`)

  const sendReply = (id: string, comment: string) =>
    request<Record<string, never>>("POST", `/me/messages/${id}/reply`, { body: { comment } })

  // Draft-creating reply actions: return a threaded draft (original quoted) for review.
  const createReplyDraft = (id: string, comment: string) =>
    request<GraphMessage>("POST", `/me/messages/${id}/createReply`, { body: { comment } })

  const createReplyAllDraft = (id: string, comment: string) =>
    request<GraphMessage>("POST", `/me/messages/${id}/createReplyAll`, { body: { comment } })

  const createForwardDraft = (
    id: string,
    comment: string,
    toRecipients: ReadonlyArray<{ emailAddress: { address: string } }>,
  ) => request<GraphMessage>("POST", `/me/messages/${id}/createForward`, { body: { comment, toRecipients } })

  // Immediate-send reply actions: thread + quote, then send in one step.
  const sendReplyAll = (id: string, comment: string) =>
    request<Record<string, never>>("POST", `/me/messages/${id}/replyAll`, { body: { comment } })

  const sendForward = (
    id: string,
    comment: string,
    toRecipients: ReadonlyArray<{ emailAddress: { address: string } }>,
  ) => request<Record<string, never>>("POST", `/me/messages/${id}/forward`, { body: { comment, toRecipients } })

  const searchMessages = (query: string, odataParams?: ODataParams) =>
    request<ODataResponse<GraphMessage>>("GET", "/me/messages", {
      odataParams: { ...odataParams, $search: query },
    })

  // Calendar
  const listEvents = (odataParams?: ODataParams) =>
    request<ODataResponse<GraphEvent>>("GET", "/me/events", { odataParams })

  // calendarView expands recurring series into individual instances within [start, end].
  // Required by Graph: startDateTime/endDateTime as query params on the path.
  const listCalendarView = (startDateTime: string, endDateTime: string, odataParams?: ODataParams) => {
    const path = `/me/calendarView?startDateTime=${encodeURIComponent(startDateTime)}&endDateTime=${encodeURIComponent(endDateTime)}`
    return request<ODataResponse<GraphEvent>>("GET", path, { odataParams })
  }

  const getEvent = (id: string, bodyContentType?: "text" | "html") =>
    request<GraphEvent>("GET", `/me/events/${id}`, bodyFormat(bodyContentType))

  const createEvent = (event: Record<string, unknown>) => request<GraphEvent>("POST", "/me/events", { body: event })

  const updateEvent = (id: string, event: Record<string, unknown>) =>
    request<GraphEvent>("PATCH", `/me/events/${id}`, { body: event })

  const deleteEvent = (id: string) => request<Record<string, never>>("DELETE", `/me/events/${id}`)

  // accept, decline and tentativelyAccept take the same body; Graph answers 202 with no content.
  const respondToEvent = (
    id: string,
    action: "accept" | "decline" | "tentativelyAccept",
    body: { comment?: string; sendResponse: boolean },
  ) => request<Record<string, never>>("POST", `/me/events/${id}/${action}`, { body })

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

  const getContact = (id: string) => request<GraphContact>("GET", `/me/contacts/${id}`)

  const createContact = (contact: Record<string, unknown>) =>
    request<GraphContact>("POST", "/me/contacts", { body: contact })

  const updateContact = (id: string, contact: Record<string, unknown>) =>
    request<GraphContact>("PATCH", `/me/contacts/${id}`, { body: contact })

  const deleteContact = (id: string) => request<Record<string, never>>("DELETE", `/me/contacts/${id}`)

  const searchContacts = (query: string, odataParams?: ODataParams) =>
    request<ODataResponse<GraphContact>>("GET", "/me/contacts", {
      odataParams: { ...odataParams, $search: query },
    })

  // Files (OneDrive)
  const listDriveItems = (folderId?: string) =>
    request<ODataResponse<GraphDriveItem>>(
      "GET",
      folderId ? `/me/drive/items/${folderId}/children` : "/me/drive/root/children",
    )

  const listDriveItemsByPath = (folderPath: string) =>
    request<ODataResponse<GraphDriveItem>>("GET", `/me/drive/root:/${folderPath}:/children`)

  const getDriveItem = (id: string) => request<GraphDriveItem>("GET", `/me/drive/items/${id}`)

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
    driveId ? `/drives/${driveId}/items/${id}` : `/me/drive/items/${id}`

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
    const version = defaultVersion()
    const url = `${GRAPH_API_BASE}/${version}${driveItemPath(id, driveId)}/content`

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
      return Left<GraphApiError, Uint8Array>({
        type: "network",
        message: `Network error: ${error instanceof Error ? error.message : String(error)}`,
      })
    }
  }

  const createFolder = (parentId: string, name: string) =>
    request<GraphDriveItem>("POST", `/me/drive/items/${parentId}/children`, {
      body: { name, folder: {}, "@microsoft.graph.conflictBehavior": "rename" },
    })

  // SharePoint Sites
  const listFollowedSites = () => request<ODataResponse<GraphSite>>("GET", "/me/followedSites")

  const searchSites = (query: string) =>
    request<ODataResponse<GraphSite>>("GET", `/sites?search=${encodeURIComponent(query)}`)

  const getSite = (siteId: string) => request<GraphSite>("GET", `/sites/${siteId}`)

  const listSiteDrives = (siteId: string) => request<ODataResponse<GraphDrive>>("GET", `/sites/${siteId}/drives`)

  const listSiteDriveItems = (siteId: string, driveId?: string, folderId?: string) => {
    if (driveId && folderId) {
      return request<ODataResponse<GraphDriveItem>>(
        "GET",
        `/sites/${siteId}/drives/${driveId}/items/${folderId}/children`,
      )
    }
    if (folderId) {
      return request<ODataResponse<GraphDriveItem>>("GET", `/sites/${siteId}/drive/items/${folderId}/children`)
    }
    if (driveId) {
      return request<ODataResponse<GraphDriveItem>>("GET", `/sites/${siteId}/drives/${driveId}/root/children`)
    }
    return request<ODataResponse<GraphDriveItem>>("GET", `/sites/${siteId}/drive/root/children`)
  }

  const listSiteDriveItemsByPath = (siteId: string, path: string, driveId?: string) => {
    const cleanPath = path.replace(/^\/+/, "")
    if (driveId) {
      return request<ODataResponse<GraphDriveItem>>(
        "GET",
        `/sites/${siteId}/drives/${driveId}/root:/${cleanPath}:/children`,
      )
    }
    return request<ODataResponse<GraphDriveItem>>("GET", `/sites/${siteId}/drive/root:/${cleanPath}:/children`)
  }

  const siteSearchPath = (siteId: string, query: string, driveId?: string) =>
    `${driveId ? `/sites/${siteId}/drives/${driveId}` : `/sites/${siteId}/drive`}/root/search(q='${searchLiteral(query)}')`

  const searchSiteFiles = (siteId: string, query: string, driveId?: string, odataParams?: ODataParams) =>
    request<ODataResponse<GraphDriveItem>>("GET", siteSearchPath(siteId, query, driveId), { odataParams })

  const searchAllSiteFiles = (siteId: string, query: string, driveId?: string) =>
    requestPaginated<GraphDriveItem>(siteSearchPath(siteId, query, driveId))

  // Teams
  const listTeams = () =>
    request<ODataResponse<{ id: string; displayName?: string; description?: string }>>("GET", "/me/joinedTeams")

  const listChannels = (teamId: string) => request<ODataResponse<GraphChannel>>("GET", `/teams/${teamId}/channels`)

  const listChannelMessages = (teamId: string, channelId: string, odataParams?: ODataParams) =>
    request<ODataResponse<GraphChannelMessage>>("GET", `/teams/${teamId}/channels/${channelId}/messages`, {
      odataParams,
    })

  const sendChannelMessage = (teamId: string, channelId: string, content: string) =>
    request<GraphChannelMessage>("POST", `/teams/${teamId}/channels/${channelId}/messages`, {
      body: { body: { content } },
    })

  // Graph only accepts replies to a top-level channel message, not to another reply.
  const replyToChannelMessage = (teamId: string, channelId: string, messageId: string, content: string) =>
    request<GraphChannelMessage>("POST", `/teams/${teamId}/channels/${channelId}/messages/${messageId}/replies`, {
      body: { body: { content } },
    })

  // Chats
  const listChats = (odataParams?: ODataParams) =>
    request<ODataResponse<GraphChat>>("GET", "/me/chats", { odataParams })

  const listChatMessages = (chatId: string, odataParams?: ODataParams) =>
    request<ODataResponse<GraphChatMessage>>("GET", `/chats/${chatId}/messages`, { odataParams })

  const sendChatMessage = (chatId: string, content: string, contentType: string = "text") =>
    request<GraphChatMessage>("POST", `/chats/${chatId}/messages`, {
      body: { body: { contentType, content } },
    })

  // Graph returns the existing chat for a one-on-one pair, but creates a new group chat every call.
  const createChat = (chat: Record<string, unknown>) => request<GraphChat>("POST", "/chats", { body: chat })

  // Users & Groups
  const getMe = () => request<GraphUser>("GET", "/me")

  const listUsers = (odataParams?: ODataParams) => request<ODataResponse<GraphUser>>("GET", "/users", { odataParams })

  const getUser = (id: string) => request<GraphUser>("GET", `/users/${id}`)

  // /users/{key} accepts an ID or a user principal name, not an email address, so an address that
  // differs from the UPN needs the mail filter below. A guest UPN contains "#EXT#", and "#" would
  // otherwise end the path.
  const CHAT_MEMBER_FIELDS = ["id", "displayName", "mail", "userPrincipalName", "userType"]

  const getUserByKey = (key: string) =>
    request<GraphUser>("GET", `/users/${key.replace(/#/g, "%23")}`, { odataParams: { $select: CHAT_MEMBER_FIELDS } })

  // userType is only returned when selected, and the caller needs it too: an in-tenant guest must
  // join a chat as a guest, the caller included.
  const getMeAsChatMember = () => request<GraphUser>("GET", "/me", { odataParams: { $select: CHAT_MEMBER_FIELDS } })

  const findUsersByMail = (address: string) =>
    request<ODataResponse<GraphUser>>("GET", "/users", {
      odataParams: { $filter: `mail eq '${address.replace(/'/g, "''")}'`, $select: CHAT_MEMBER_FIELDS },
    })

  const listGroups = (odataParams?: ODataParams) =>
    request<ODataResponse<GraphGroup>>("GET", "/groups", { odataParams })

  const getGroup = (id: string) => request<GraphGroup>("GET", `/groups/${id}`)

  const listGroupMembers = (id: string) => request<ODataResponse<GraphUser>>("GET", `/groups/${id}/members`)

  // Planner
  const listPlans = () => request<ODataResponse<GraphPlan>>("GET", "/me/planner/plans")

  // The user's group memberships (groups only, filtered from the memberOf directory objects) and a
  // group's plans — used to fan out plan discovery, since /me/planner/plans omits group-owned plans
  // the user wasn't explicitly added to.
  const listMyGroups = () =>
    request<ODataResponse<GraphGroup>>("GET", "/me/memberOf/microsoft.graph.group?$select=id,displayName")

  const listGroupPlans = (groupId: string) =>
    request<ODataResponse<GraphPlan>>("GET", `/groups/${groupId}/planner/plans`)

  const listPlannerBuckets = (planId: string) =>
    request<ODataResponse<GraphBucket>>("GET", `/planner/plans/${planId}/buckets`)

  const createPlannerBucket = (bucket: Record<string, unknown>) =>
    request<GraphBucket>("POST", "/planner/buckets", { body: bucket })

  const listPlannerTasks = (planId: string) =>
    request<ODataResponse<GraphPlannerTask>>("GET", `/planner/plans/${planId}/tasks`)

  const getPlannerTask = (id: string) => request<GraphPlannerTask>("GET", `/planner/tasks/${id}`)

  const createPlannerTask = (task: Record<string, unknown>) =>
    request<GraphPlannerTask>("POST", "/planner/tasks", { body: task })

  const updatePlannerTask = (id: string, task: Record<string, unknown>, etag: string) =>
    request<GraphPlannerTask>("PATCH", `/planner/tasks/${id}`, {
      body: task,
      headers: { "If-Match": etag },
    })

  // Task details (description / checklist / references) is a separate object with its own ETag,
  // so its writes need a distinct If-Match from the parent task's.
  const getPlannerTaskDetails = (id: string) =>
    request<{ "@odata.etag": string } & Record<string, unknown>>("GET", `/planner/tasks/${id}/details`)

  const updatePlannerTaskDetails = (id: string, details: Record<string, unknown>, etag: string) =>
    request<Record<string, unknown>>("PATCH", `/planner/tasks/${id}/details`, {
      body: details,
      headers: { "If-Match": etag },
    })

  // OneNote
  const listOnenoteNotebooks = () => request<ODataResponse<GraphNotebook>>("GET", "/me/onenote/notebooks")

  const listOnenoteSections = (notebookId: string) =>
    request<ODataResponse<GraphSection>>("GET", `/me/onenote/notebooks/${notebookId}/sections`)

  const listOnenotePages = (sectionId: string) =>
    request<ODataResponse<GraphPage>>("GET", `/me/onenote/sections/${sectionId}/pages`)

  // The /content endpoint returns text/html, not JSON — read it as raw text.
  const getOnenotePageContent = (pageId: string) =>
    request<string>("GET", `/me/onenote/pages/${pageId}/content`, { responseType: "text" })

  // OneNote writes. createOnenotePage sends raw text/html; the rest are JSON.
  const createOnenotePage = (sectionId: string, html: string) =>
    request<GraphPage>("POST", `/me/onenote/sections/${sectionId}/pages`, { body: html, contentType: "text/html" })

  const updateOnenotePageContent = (pageId: string, commands: readonly unknown[]) =>
    request<Record<string, never>>("PATCH", `/me/onenote/pages/${pageId}/content`, { body: commands })

  const createOnenoteSection = (notebookId: string, displayName: string) =>
    request<GraphSection>("POST", `/me/onenote/notebooks/${notebookId}/sections`, { body: { displayName } })

  const createOnenoteNotebook = (displayName: string) =>
    request<GraphNotebook>("POST", "/me/onenote/notebooks", { body: { displayName } })

  const copyOnenotePage = (pageId: string, sectionId: string) =>
    request<Record<string, never>>("POST", `/me/onenote/pages/${pageId}/copyToSection`, { body: { id: sectionId } })

  const deleteOnenotePage = (pageId: string) => request<Record<string, never>>("DELETE", `/me/onenote/pages/${pageId}`)

  // To Do
  const listTodoLists = () => request<ODataResponse<GraphTodoList>>("GET", "/me/todo/lists")

  const listTodoTasks = (listId: string) =>
    request<ODataResponse<GraphTodoTask>>("GET", `/me/todo/lists/${listId}/tasks`)

  const createTodoTask = (listId: string, task: Record<string, unknown>) =>
    request<GraphTodoTask>("POST", `/me/todo/lists/${listId}/tasks`, { body: task })

  const updateTodoTask = (listId: string, taskId: string, task: Record<string, unknown>) =>
    request<GraphTodoTask>("PATCH", `/me/todo/lists/${listId}/tasks/${taskId}`, { body: task })

  const getTodoTask = (listId: string, taskId: string) =>
    request<GraphTodoTask>("GET", `/me/todo/lists/${listId}/tasks/${taskId}`)

  const deleteTodoTask = (listId: string, taskId: string) =>
    request<Record<string, never>>("DELETE", `/me/todo/lists/${listId}/tasks/${taskId}`)

  // Text-only file upload. Binary uploads must use get_upload_config (httpStream) or upload_file_from_path (stdio).
  const uploadFile = async (
    path: string,
    content: string,
    contentType: string = "text/plain",
    conflictBehavior: "rename" | "replace" | "fail" = "rename",
  ): Promise<Either<GraphApiError, GraphDriveItem>> => {
    const tokenResult = await auth.getAccessToken()
    if (tokenResult.isLeft()) {
      return Left<GraphApiError, GraphDriveItem>({
        type: "auth",
        message: (tokenResult.value as { message: string }).message,
      })
    }

    const token = tokenResult.value as string
    const version = defaultVersion()
    const separator = path.includes("?") ? "&" : "?"
    const url = `${GRAPH_API_BASE}/${version}${path}${separator}@microsoft.graph.conflictBehavior=${conflictBehavior}`

    try {
      const response = await resilientFetch(
        url,
        {
          method: "PUT",
          headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": contentType,
          },
          body: content,
        },
        { retryTransient: false },
      )

      if (!response.ok) {
        return mapHttpError<GraphDriveItem>(response)
      }

      const text = await response.text()
      return parseJsonResponse<GraphDriveItem>(text)
    } catch (error) {
      return Left<GraphApiError, GraphDriveItem>({
        type: "network",
        message: `Network error: ${error instanceof Error ? error.message : String(error)}`,
      })
    }
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

    listAttachments,
    sendMessage,
    createDraft,
    sendDraft,
    getDraftState,
    updateMessage,
    deleteMessage,
    sendReply,
    sendReplyAll,
    sendForward,
    createReplyDraft,
    createReplyAllDraft,
    createForwardDraft,
    searchMessages,
    // Calendar
    listEvents,
    listCalendarView,
    getEvent,
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
    listTodoTasks,
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
