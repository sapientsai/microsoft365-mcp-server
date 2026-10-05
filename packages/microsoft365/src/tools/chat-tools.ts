import { UserError } from "fastmcp"
import type { Either } from "functype/either"
import { Left, Right } from "functype/either"

import { getGraphClient } from "../client/graph-client"
import type { GraphChat, GraphChatMessage, ODataResponse } from "../types"
import { formatChatList, formatChatMessageList, isReadableChatMessage } from "../utils/formatters"

const requireClient = () => {
  const client = getGraphClient()
  if (client.isNone()) return null
  return client.orThrow()
}

// Members give an untitled chat its name; lastMessagePreview gives the time it was last active, and
// is the only sort Graph supports here (descending only). Both expands are documented for /me/chats.
const CHAT_LIST_PARAMS = {
  $expand: ["members", "lastMessagePreview"],
  $orderby: "lastMessagePreview/createdDateTime desc",
}

type Client = NonNullable<ReturnType<typeof requireClient>>

// A since value is parsed and re-serialized before it reaches a $filter, so only a real timestamp
// can: "2026-10-01 or true" would otherwise widen the filter.
// A time zone is required: JavaScript reads a date-time without one in the server's local zone, so
// "2026-10-01T00:00:00" would mean a different instant in the container than on a laptop.
const parseSince = (since: string): Either<UserError, string> => {
  const time = Date.parse(since)
  return Number.isNaN(time) || !/(Z|[+-]\d{2}:?\d{2})$/i.test(since.trim())
    ? Left(
        new UserError(
          `since must be an ISO 8601 date-time with a time zone, e.g. 2026-10-01T00:00:00Z (got "${since}")`,
        ),
      )
    : Right(new Date(time).toISOString())
}

// Graph's @odata.nextLink is absolute, while client.request takes a path under the API version.
export const nextLinkPath = (link: string): string => link.replace(/^https:\/\/graph\.microsoft\.com\/(v1\.0|beta)/, "")

const MAX_SINCE_PAGES = 20

const previewTime = (chat: GraphChat): number | undefined => {
  const time = Date.parse(chat.lastMessagePreview?.createdDateTime ?? "")
  return Number.isNaN(time) ? undefined : time
}

type SincePage = { readonly chats: ReadonlyArray<GraphChat>; readonly reachedOlder: boolean }

// Chats arrive newest first, so the first one whose last message is not after since ends the scan.
// A chat with no preview (no messages) is skipped, not treated as the end: its position in the sort
// is undocumented, and stopping on one placed first would return nothing.
const chatsAfter = (chats: ReadonlyArray<GraphChat>, sinceMs: number): SincePage => {
  const firstOlder = chats.findIndex((chat) => {
    const time = previewTime(chat)
    return time !== undefined && time <= sinceMs
  })
  const scanned = firstOlder === -1 ? chats : chats.slice(0, firstOlder)
  return { chats: scanned.filter((chat) => previewTime(chat) !== undefined), reachedOlder: firstOlder !== -1 }
}

const listChatsSince = async (
  client: Client,
  sinceMs: number,
  page: number = 0,
  link?: string,
): Promise<Either<UserError, { readonly chats: ReadonlyArray<GraphChat>; readonly truncated: boolean }>> => {
  const result =
    link === undefined
      ? await client.listChats({ ...CHAT_LIST_PARAMS, $top: 50 })
      : await client.request<ODataResponse<GraphChat>>("GET", nextLinkPath(link))
  if (result.isLeft())
    return Left(new UserError(`Failed to list chats: ${(result.value as { message: string }).message}`))

  const response = result.value as ODataResponse<GraphChat>
  const { chats, reachedOlder } = chatsAfter(response.value, sinceMs)
  const next = response["@odata.nextLink"]
  if (reachedOlder || next === undefined) return Right({ chats, truncated: false })
  if (page + 1 >= MAX_SINCE_PAGES) return Right({ chats, truncated: true })

  const rest = await listChatsSince(client, sinceMs, page + 1, next)
  return rest.map((later) => ({ chats: [...chats, ...later.chats], truncated: later.truncated }))
}

export const listChats = async (params?: {
  top?: number
  since?: string
  fetch_all_pages?: boolean
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  if (params?.since !== undefined) {
    const since = parseSince(params.since)
    if (since.isLeft()) return Left(since.value as UserError)
    const found = await listChatsSince(client, Date.parse(since.value as string))
    return found.map(({ chats, truncated }) => {
      const kept = params.top === undefined ? chats : chats.slice(0, params.top)
      const note = truncated
        ? `\n\nNote: stopped after ${MAX_SINCE_PAGES} pages; chats active earlier in the window may be missing.`
        : ""
      return `${formatChatList(kept)}${note}`
    })
  }

  if (params?.fetch_all_pages) {
    const result = await client.requestPaginated<GraphChat>("/me/chats", { odataParams: CHAT_LIST_PARAMS })
    return result
      .mapLeft((error) => new UserError(`Failed to list chats: ${error.message}`))
      .map((items) => formatChatList(items))
  }

  const result = await client.listChats({ ...CHAT_LIST_PARAMS, $top: params?.top ?? 25 })
  return result
    .mapLeft((error) => new UserError(`Failed to list chats: ${error.message}`))
    .map((response) => formatChatList((response as ODataResponse<never>).value))
}

const DEFAULT_MESSAGE_TOP = 25

const messagesSince = async (
  client: Client,
  chatId: string,
  since: string,
  top: number,
  page: number = 0,
  link?: string,
): Promise<Either<UserError, { readonly messages: ReadonlyArray<GraphChatMessage>; readonly more: boolean }>> => {
  const result =
    link === undefined
      ? await client.listChatMessages(chatId, {
          $top: Math.min(top, 50),
          $orderby: "lastModifiedDateTime desc",
          $filter: `lastModifiedDateTime gt ${since}`,
        })
      : await client.request<ODataResponse<GraphChatMessage>>("GET", nextLinkPath(link))
  if (result.isLeft())
    return Left(new UserError(`Failed to list chat messages: ${(result.value as { message: string }).message}`))

  const response = result.value as ODataResponse<GraphChatMessage>
  const messages = response.value.filter(isReadableChatMessage)
  const next = response["@odata.nextLink"]
  if (messages.length >= top || next === undefined) return Right({ messages, more: next !== undefined })
  if (page + 1 >= MAX_SINCE_PAGES) return Right({ messages, more: true })

  const rest = await messagesSince(client, chatId, since, top - messages.length, page + 1, next)
  return rest.map((later) => ({ messages: [...messages, ...later.messages], more: later.more }))
}

export const listChatMessages = async (params: {
  chat_id: string
  top?: number
  since?: string
  max_chars?: number
  fetch_all_pages?: boolean
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const since = params.since === undefined ? undefined : parseSince(params.since)
  if (since?.isLeft()) return Left(since.value as UserError)

  // Looked up on every call, never cached: one server serves many users, and a shared id would tag
  // another person's messages [You]. A failed lookup drops the two flags rather than the listing.
  const me = await client.getMe()
  const meId = me.fold(
    () => undefined,
    (user) => (user as { id?: string }).id,
  )
  const options = { meId, maxChars: params.max_chars, selfUnresolved: meId === undefined }

  // Graph ignores a lastModifiedDateTime filter unless the request also orders by it (List messages
  // in a chat docs, "Optional query parameters"). top is the caller's budget: pages are read only
  // until that many readable messages are in hand, and a note says when more matched.
  if (since !== undefined) {
    const top = params.top ?? DEFAULT_MESSAGE_TOP
    const found = await messagesSince(client, params.chat_id, since.value as string, top)
    return found.map(({ messages, more }) =>
      formatChatMessageList(messages.slice(0, top), { ...options, moreRemain: more || messages.length > top }),
    )
  }

  if (params.fetch_all_pages) {
    const result = await client.requestPaginated<GraphChatMessage>(`/chats/${params.chat_id}/messages`)
    return result
      .mapLeft((error) => new UserError(`Failed to list chat messages: ${error.message}`))
      .map((items) => formatChatMessageList(items, options))
  }

  const result = await client.listChatMessages(params.chat_id, { $top: params.top ?? DEFAULT_MESSAGE_TOP })
  return result
    .mapLeft((error) => new UserError(`Failed to list chat messages: ${error.message}`))
    .map((response) => formatChatMessageList((response as ODataResponse<GraphChatMessage>).value, options))
}

export const sendChatMessage = async (params: {
  chat_id: string
  content: string
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.sendChatMessage(params.chat_id, params.content)
  return result
    .mapLeft((error) => new UserError(`Failed to send chat message: ${error.message}`))
    .map(() => "Chat message sent.")
}
