import { UserError } from "fastmcp"
import type { Either } from "functype/either"
import { Left, Right } from "functype/either"

import { getGraphClient } from "../client/graph-client"
import type { GraphChat, GraphChatMessage, GraphUser, ODataResponse } from "../types"
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

// Members bind by directory ID. In-tenant guests must carry the "guest" role and everyone else
// "owner", per Graph's create-chat rules.
const memberBind = (user: GraphUser): Record<string, unknown> => ({
  "@odata.type": "#microsoft.graph.aadUserConversationMember",
  roles: [user.userType === "Guest" ? "guest" : "owner"],
  "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${user.id}')`,
})

const parseChatRecipients = (to: string): ReadonlyArray<string> => [
  ...new Map(
    to
      .split(",")
      .map((address) => address.trim())
      .filter((address) => address.length > 0)
      .map((address) => [address.toLowerCase(), address] as const),
  ).values(),
]

// A sign-in name resolves directly; an email address that differs from it (or an alias held in
// mail) needs the filter. Only people in the caller's directory, guests included, can be found.
const resolveRecipient = async (client: Client, address: string): Promise<Either<UserError, GraphUser>> => {
  const direct = await client.getUserByKey(address)
  if (direct.isRight()) return Right(direct.value as GraphUser)

  const byMail = await client.findUsersByMail(address)
  if (byMail.isLeft())
    return Left(new UserError(`Couldn't look up ${address}: ${(byMail.value as { message: string }).message}`))
  const match = (byMail.value as ODataResponse<GraphUser>).value.at(0)
  return match === undefined
    ? Left(new UserError(`No one in your organization matches ${address}. Use their work email or sign-in name.`))
    : Right(match)
}

const resolveRecipients = async (
  client: Client,
  addresses: ReadonlyArray<string>,
): Promise<Either<UserError, ReadonlyArray<GraphUser>>> => {
  const resolved = await Promise.all(addresses.map((address) => resolveRecipient(client, address)))
  const failed = resolved.find((result) => result.isLeft())
  if (failed !== undefined) return Left(failed.value as UserError)
  return Right(resolved.map((result) => result.value as GraphUser))
}

// Every member, the caller included, has to be listed when a chat is created. Graph hands back the
// existing chat for a one-on-one pair, but creates a new group chat on every call.
const findOrCreateChat = async (
  client: Client,
  to: string,
  topic: string | undefined,
): Promise<Either<UserError, GraphChat>> => {
  const me = await client.getMe()
  if (me.isLeft())
    return Left(new UserError(`Failed to read your profile: ${(me.value as { message: string }).message}`))
  const self = me.value as GraphUser

  const looked = await resolveRecipients(client, parseChatRecipients(to))
  if (looked.isLeft()) return Left(looked.value as UserError)
  // Matched by ID, so the caller's own aliases and two spellings of one person both collapse.
  const others = [
    ...new Map(
      (looked.value as ReadonlyArray<GraphUser>).filter((user) => user.id !== self.id).map((user) => [user.id, user]),
    ).values(),
  ]
  if (others.length === 0) {
    return Left(new UserError("Name at least one other person in 'to'. To message yourself, use chat_id '48:notes'."))
  }

  const oneOnOne = others.length === 1
  if (oneOnOne && topic !== undefined) {
    return Left(new UserError("A topic can only be set on a group chat, which needs two or more people in 'to'."))
  }

  const chat = {
    chatType: oneOnOne ? "oneOnOne" : "group",
    members: [memberBind(self), ...others.map(memberBind)],
    ...(topic === undefined ? {} : { topic }),
  }
  const result = await client.createChat(chat)
  return result.mapLeft((error) => new UserError(`Failed to start chat: ${error.message}`))
}

export const sendChatMessage = async (params: {
  chat_id?: string
  to?: string
  topic?: string
  content: string
}): Promise<Either<UserError, string>> => {
  if ((params.chat_id === undefined) === (params.to === undefined)) {
    return Left(new UserError("Pass exactly one of chat_id (an existing chat) or to (people to message)."))
  }
  if (params.chat_id !== undefined && params.topic !== undefined) {
    return Left(new UserError("topic only applies when starting a group chat with 'to'."))
  }

  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  if (params.chat_id !== undefined) {
    const result = await client.sendChatMessage(params.chat_id, params.content)
    return result
      .mapLeft((error) => new UserError(`Failed to send chat message: ${error.message}`))
      .map(() => "Chat message sent.")
  }

  const chat = await findOrCreateChat(client, params.to as string, params.topic)
  if (chat.isLeft()) return Left(chat.value as UserError)
  const { id, chatType } = chat.value as GraphChat

  const sent = await client.sendChatMessage(id, params.content)
  const where =
    chatType === "group"
      ? `Started a new group chat (chat ID: ${id}). Use this chat_id for follow-ups; 'to' would start another one.`
      : `Chat ID: ${id}.`
  return sent
    .mapLeft((error) => new UserError(`Chat is ready (chat ID: ${id}), but sending failed: ${error.message}`))
    .map(() => `Chat message sent. ${where}`)
}
