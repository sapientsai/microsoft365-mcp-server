import { None, type Option, Some } from "functype"
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
  GraphDriveItem,
  GraphEvent,
  GraphGroup,
  GraphMessage,
  GraphNotebook,
  GraphPage,
  GraphPlan,
  GraphPlannerTask,
  GraphSection,
  GraphTodoList,
  GraphTodoTask,
  GraphUser,
  ODataParams,
  ODataResponse,
} from "../types"
import { buildODataQuery } from "../utils/odata-helpers"
import { fetchAllPages, parseJsonResponse } from "../utils/pagination"

type RequestOptions = {
  readonly version?: GraphApiVersion
  readonly body?: Record<string, unknown>
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
    const url = `${GRAPH_API_BASE}/${version}${path}${queryString}`

    // eslint-disable-next-line functype/prefer-either -- boundary between throwing fetch API and Either-returning client
    try {
      const fetchOptions: RequestInit = {
        method,
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
          ...(options?.headers ?? {}),
        },
      }

      if (options?.body && (method === "POST" || method === "PUT" || method === "PATCH")) {
        fetchOptions.body = JSON.stringify(options.body)
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

      return parseJsonResponse<T>(text)
    } catch (error) {
      return Left<GraphApiError, T>({
        type: "network",
        message: `Network error: ${error instanceof Error ? error.message : String(error)}`,
      })
    }
  }

  const mapHttpError = async <T>(response: Response): Promise<Either<GraphApiError, T>> => {
    let graphErrorCode: string | undefined
    let message = `Microsoft Graph API error: ${response.status} ${response.statusText}`

    // eslint-disable-next-line functype/prefer-either -- boundary: parsing error response body
    try {
      const errorBody = await response.json()
      if (errorBody?.error?.message) {
        message = errorBody.error.message as string
      }
      if (errorBody?.error?.code) {
        graphErrorCode = errorBody.error.code as string
      }
    } catch {
      // Could not parse error body
    }

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
    const initialUrl = `${GRAPH_API_BASE}/${version}${path}${queryString}`

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

  const replyToMessage = (id: string, comment: string) =>
    request<Record<string, never>>("POST", `/me/messages/${id}/reply`, { body: { comment } })

  const searchMessages = (query: string, odataParams?: ODataParams) =>
    request<ODataResponse<GraphMessage>>("GET", "/me/messages", {
      odataParams: { ...odataParams, $search: query },
    })

  // Calendar
  const listEvents = (odataParams?: ODataParams) =>
    request<ODataResponse<GraphEvent>>("GET", "/me/events", { odataParams })

  const getEvent = (id: string) => request<GraphEvent>("GET", `/me/events/${id}`)

  const createEvent = (event: Record<string, unknown>) => request<GraphEvent>("POST", "/me/events", { body: event })

  const updateEvent = (id: string, event: Record<string, unknown>) =>
    request<GraphEvent>("PATCH", `/me/events/${id}`, { body: event })

  const deleteEvent = (id: string) => request<Record<string, never>>("DELETE", `/me/events/${id}`)

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
  const listNotebooks = () => request<ODataResponse<GraphNotebook>>("GET", "/me/onenote/notebooks")

  const listSections = (notebookId: string) =>
    request<ODataResponse<GraphSection>>("GET", `/me/onenote/notebooks/${notebookId}/sections`)

  const listPages = (sectionId: string) =>
    request<ODataResponse<GraphPage>>("GET", `/me/onenote/sections/${sectionId}/pages`)

  const getPageContent = (pageId: string) => request<string>("GET", `/me/onenote/pages/${pageId}/content`)

  // To Do
  const listTodoLists = () => request<ODataResponse<GraphTodoList>>("GET", "/me/todo/lists")

  const listTodoTasks = (listId: string) =>
    request<ODataResponse<GraphTodoTask>>("GET", `/me/todo/lists/${listId}/tasks`)

  const createTodoTask = (listId: string, task: Record<string, unknown>) =>
    request<GraphTodoTask>("POST", `/me/todo/lists/${listId}/tasks`, { body: task })

  const updateTodoTask = (listId: string, taskId: string, task: Record<string, unknown>) =>
    request<GraphTodoTask>("PATCH", `/me/todo/lists/${listId}/tasks/${taskId}`, { body: task })

  // File upload (raw content, not JSON)
  const uploadFile = async (
    path: string,
    content: string,
    contentType: string = "text/plain",
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
    const url = `${GRAPH_API_BASE}/${version}${path}`

    // eslint-disable-next-line functype/prefer-either -- boundary: fetch API
    try {
      const body = contentType === "application/octet-stream" ? Buffer.from(content, "base64") : content
      const response = await fetch(url, {
        method: "PUT",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": contentType,
        },
        body,
      })

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
    replyToMessage,
    searchMessages,
    // Calendar
    listEvents,
    getEvent,
    createEvent,
    updateEvent,
    deleteEvent,
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
    listNotebooks,
    listSections,
    listPages,
    getPageContent,
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

let client: Option<GraphClient> = None()

export const initializeGraphClient = (): GraphClient => {
  const c = createGraphClient()
  client = Some(c)
  return c
}

export const getGraphClient = (): Option<GraphClient> => client
