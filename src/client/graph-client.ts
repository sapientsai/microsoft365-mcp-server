import { None, type Option, Ref, Some } from "functype"
import { type Either, Left, Right } from "functype/either"

import { getAccessToken } from "../auth"
import { GRAPH_API_BASE } from "../auth/scopes"
import type {
  GraphApiError,
  GraphApiVersion,
  GraphChannel,
  GraphChannelMessage,
  GraphChat,
  GraphChatMessage,
  GraphContact,
  GraphDrive,
  GraphDriveItem,
  GraphEvent,
  GraphGroup,
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
import { appendODataQuery, buildODataQuery } from "../utils/odata-helpers"
import { fetchAllPages, parseJsonResponse } from "../utils/pagination"

type RequestOptions = {
  readonly version?: GraphApiVersion
  readonly body?: Record<string, unknown> | readonly unknown[] | string
  readonly contentType?: string
  readonly responseType?: "json" | "text"
  readonly odataParams?: ODataParams
  readonly headers?: Record<string, string>
}

const defaultVersion = (): GraphApiVersion => (process.env.MS365_GRAPH_VERSION === "beta" ? "beta" : "v1.0")

const createGraphClient = () => {
  const request = async <T>(
    method: string,
    path: string,
    options?: RequestOptions,
  ): Promise<Either<GraphApiError, T>> => {
    const tokenResult = await getAccessToken()

    if (tokenResult.isLeft()) {
      return Left<GraphApiError, T>({
        type: "auth",
        message: (tokenResult.value as { message: string }).message,
      })
    }

    const token = tokenResult.value as string
    const version = options?.version ?? defaultVersion()
    const queryString = buildODataQuery(options?.odataParams)
    const url = `${GRAPH_API_BASE}/${version}${appendODataQuery(path, queryString)}`

    // eslint-disable-next-line functype/prefer-either -- boundary between throwing fetch API and Either-returning client
    try {
      const fetchOptions: RequestInit = {
        method,
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": options?.contentType ?? "application/json",
          ...(options?.headers ?? {}),
        },
      }

      if (options?.body !== undefined && (method === "POST" || method === "PUT" || method === "PATCH")) {
        fetchOptions.body = typeof options.body === "string" ? options.body : JSON.stringify(options.body)
      }

      const response = await fetch(url, fetchOptions)

      if (!response.ok) {
        return mapHttpError<T>(response)
      }

      // Handle 204 No Content
      if (response.status === 204) {
        return Right<GraphApiError, T>({} as T)
      }

      const text = await response.text()
      if (!text || text.trim() === "") {
        return Right<GraphApiError, T>({} as T)
      }

      // Some endpoints return raw text (e.g. OneNote page /content is text/html, not JSON).
      if (options?.responseType === "text") {
        return Right<GraphApiError, T>(text as T)
      }

      return parseJsonResponse<T>(text)
    } catch (error) {
      return Left<GraphApiError, T>({
        type: "network",
        message: `Network error: ${error instanceof Error ? error.message : String(error)}`,
      })
    }
  }

  const mapHttpError = async <T>(response: Response): Promise<Either<GraphApiError, T>> => {
    const fallbackMessage = `Microsoft Graph API error: ${response.status} ${response.statusText}`

    const { message, graphErrorCode } = await (async (): Promise<{ message: string; graphErrorCode?: string }> => {
      try {
        const errorBody = await response.json()
        return {
          message: (errorBody?.error?.message as string | undefined) ?? fallbackMessage,
          graphErrorCode: errorBody?.error?.code as string | undefined,
        }
      } catch {
        return { message: fallbackMessage }
      }
    })()

    const retryAfter = response.headers.get("Retry-After")

    switch (response.status) {
      case 401:
        return Left<GraphApiError, T>({ type: "auth", message, status: 401, graphErrorCode })
      case 403:
        return Left<GraphApiError, T>({ type: "forbidden", message, status: 403, graphErrorCode })
      case 404:
        return Left<GraphApiError, T>({ type: "not_found", message, status: 404, graphErrorCode })
      case 429:
        return Left<GraphApiError, T>({
          type: "throttle",
          message,
          status: 429,
          graphErrorCode,
          retryAfter: retryAfter ? parseInt(retryAfter, 10) : undefined,
        })
      default:
        return Left<GraphApiError, T>({ type: "api", message, status: response.status, graphErrorCode })
    }
  }

  const requestPaginated = async <T>(
    path: string,
    options?: RequestOptions,
  ): Promise<Either<GraphApiError, ReadonlyArray<T>>> => {
    const version = options?.version ?? defaultVersion()
    const queryString = buildODataQuery(options?.odataParams)
    const initialUrl = `${GRAPH_API_BASE}/${version}${appendODataQuery(path, queryString)}`

    return fetchAllPages<T>(async (url: string) => {
      const tokenResult = await getAccessToken()

      if (tokenResult.isLeft()) {
        return Left<GraphApiError, ODataResponse<T>>({
          type: "auth",
          message: (tokenResult.value as { message: string }).message,
        })
      }

      const token = tokenResult.value as string

      // eslint-disable-next-line functype/prefer-either -- boundary: fetch API
      try {
        const response = await fetch(url, {
          headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": "application/json",
            ...(options?.headers ?? {}),
          },
        })

        if (!response.ok) {
          return mapHttpError<ODataResponse<T>>(response)
        }

        const text = await response.text()
        return parseJsonResponse<ODataResponse<T>>(text)
      } catch (error) {
        return Left<GraphApiError, ODataResponse<T>>({
          type: "network",
          message: `Network error: ${error instanceof Error ? error.message : String(error)}`,
        })
      }
    }, initialUrl)
  }

  // Mail
  const listMessages = (odataParams?: ODataParams) =>
    request<ODataResponse<GraphMessage>>("GET", "/me/messages", { odataParams })

  const getMessage = (id: string) => request<GraphMessage>("GET", `/me/messages/${id}`)

  const sendMessage = (message: Record<string, unknown>) =>
    request<Record<string, never>>("POST", "/me/sendMail", { body: message })

  const createDraft = (message: Record<string, unknown>) =>
    request<GraphMessage>("POST", "/me/messages", { body: message })

  const sendDraft = (messageId: string) => request<Record<string, never>>("POST", `/me/messages/${messageId}/send`)

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

  const getEvent = (id: string) => request<GraphEvent>("GET", `/me/events/${id}`)

  const createEvent = (event: Record<string, unknown>) => request<GraphEvent>("POST", "/me/events", { body: event })

  const updateEvent = (id: string, event: Record<string, unknown>) =>
    request<GraphEvent>("PATCH", `/me/events/${id}`, { body: event })

  const deleteEvent = (id: string) => request<Record<string, never>>("DELETE", `/me/events/${id}`)

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

  const searchFiles = (query: string) =>
    request<ODataResponse<GraphDriveItem>>("GET", `/me/drive/root/search(q='${encodeURIComponent(query)}')`)

  const downloadFile = (id: string) => request<GraphDriveItem>("GET", `/me/drive/items/${id}`)

  const downloadFileContent = async (id: string): Promise<Either<GraphApiError, string>> => {
    const tokenResult = await getAccessToken()
    if (tokenResult.isLeft()) {
      return Left<GraphApiError, string>({
        type: "auth",
        message: (tokenResult.value as { message: string }).message,
      })
    }
    const token = tokenResult.value as string
    const version = defaultVersion()
    const url = `${GRAPH_API_BASE}/${version}/me/drive/items/${id}/content`
    // eslint-disable-next-line functype/prefer-either -- boundary between throwing fetch API and Either-returning client
    try {
      const response = await fetch(url, {
        method: "GET",
        headers: { Authorization: `Bearer ${token}` },
        redirect: "follow",
      })
      if (!response.ok) {
        return mapHttpError<string>(response)
      }
      const text = await response.text()
      return Right<GraphApiError, string>(text)
    } catch (error) {
      return Left<GraphApiError, string>({
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

  const searchSiteFiles = (siteId: string, query: string, driveId?: string) => {
    if (driveId) {
      return request<ODataResponse<GraphDriveItem>>(
        "GET",
        `/sites/${siteId}/drives/${driveId}/root/search(q='${encodeURIComponent(query)}')`,
      )
    }
    return request<ODataResponse<GraphDriveItem>>(
      "GET",
      `/sites/${siteId}/drive/root/search(q='${encodeURIComponent(query)}')`,
    )
  }

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

  // Chats
  const listChats = (odataParams?: ODataParams) =>
    request<ODataResponse<GraphChat>>("GET", "/me/chats", { odataParams })

  const listChatMessages = (chatId: string, odataParams?: ODataParams) =>
    request<ODataResponse<GraphChatMessage>>("GET", `/chats/${chatId}/messages`, { odataParams })

  const sendChatMessage = (chatId: string, content: string, contentType: string = "text") =>
    request<GraphChatMessage>("POST", `/chats/${chatId}/messages`, {
      body: { body: { contentType, content } },
    })

  // Users & Groups
  const getMe = () => request<GraphUser>("GET", "/me")

  const listUsers = (odataParams?: ODataParams) => request<ODataResponse<GraphUser>>("GET", "/users", { odataParams })

  const getUser = (id: string) => request<GraphUser>("GET", `/users/${id}`)

  const listGroups = (odataParams?: ODataParams) =>
    request<ODataResponse<GraphGroup>>("GET", "/groups", { odataParams })

  const getGroup = (id: string) => request<GraphGroup>("GET", `/groups/${id}`)

  const listGroupMembers = (id: string) => request<ODataResponse<GraphUser>>("GET", `/groups/${id}/members`)

  // Planner
  const listPlans = () => request<ODataResponse<GraphPlan>>("GET", "/me/planner/plans")

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

  // Text-only file upload. Binary uploads must use get_upload_config (httpStream) or upload_file_from_path (stdio).
  const uploadFile = async (
    path: string,
    content: string,
    contentType: string = "text/plain",
    conflictBehavior: "rename" | "replace" | "fail" = "rename",
  ): Promise<Either<GraphApiError, GraphDriveItem>> => {
    const tokenResult = await getAccessToken()
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

    const controller = new AbortController()
    const timeout = setTimeout(() => controller.abort(), 60_000)

    // eslint-disable-next-line functype/prefer-either -- boundary: fetch API
    try {
      const response = await fetch(url, {
        method: "PUT",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": contentType,
        },
        body: content,
        signal: controller.signal,
      })

      if (!response.ok) {
        return mapHttpError<GraphDriveItem>(response)
      }

      const text = await response.text()
      return parseJsonResponse<GraphDriveItem>(text)
    } catch (error) {
      const isAbort = error instanceof Error && error.name === "AbortError"
      return Left<GraphApiError, GraphDriveItem>({
        type: "network",
        message: isAbort
          ? "Upload timed out after 60s"
          : `Network error: ${error instanceof Error ? error.message : String(error)}`,
      })
    } finally {
      clearTimeout(timeout)
    }
  }

  // Generic escape hatch
  const graphQuery = <T = unknown>(
    method: string,
    path: string,
    body?: Record<string, unknown>,
    version?: GraphApiVersion,
  ) => request<T>(method, path, { body, version })

  return Object.freeze({
    // Core
    request,
    requestPaginated,
    // Mail
    listMessages,
    getMessage,
    sendMessage,
    createDraft,
    sendDraft,
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
    findMeetingTimes,
    // Contacts
    listContacts,
    getContact,
    createContact,
    searchContacts,
    // Files
    listDriveItems,
    listDriveItemsByPath,
    getDriveItem,
    searchFiles,
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
    // Chats
    listChats,
    listChatMessages,
    sendChatMessage,
    // Teams
    listTeams,
    listChannels,
    listChannelMessages,
    sendChannelMessage,
    // Users & Groups
    getMe,
    listUsers,
    getUser,
    listGroups,
    getGroup,
    listGroupMembers,
    // Planner
    listPlans,
    listPlannerTasks,
    getPlannerTask,
    createPlannerTask,
    updatePlannerTask,
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
    // Upload
    uploadFile,
    // Generic
    graphQuery,
  })
}

export type GraphClient = ReturnType<typeof createGraphClient>

const clientRef = Ref<Option<GraphClient>>(None())

export const initializeGraphClient = (): GraphClient => {
  const c = createGraphClient()
  clientRef.set(Some(c))
  return c
}

export const getGraphClient = (): Option<GraphClient> => clientRef.get()
