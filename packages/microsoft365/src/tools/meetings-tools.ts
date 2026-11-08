import { resilientFetch } from "@sapientsai/ms-graph-core"
import type { Either } from "functype/either"
import { Left, Right } from "functype/either"
import { UserError } from "somamcp"

import { getAccessToken } from "../auth"
import { GRAPH_API_BASE } from "../auth/scopes"
import { heldScope, tokenScopes, tokenTenant } from "../auth/token-facts"
import { getGraphClient } from "../client/graph-client"
import { defaultVersion } from "../client/graph-version"
import type { GraphCallTranscript, ODataResponse } from "../types"
import { formatTranscriptList } from "../utils/formatters"

const requireClient = () => {
  const client = getGraphClient()
  if (client.isNone()) return null
  return client.orThrow()
}

const DEFAULT_MAX_CHARS = 50000

// Transcript content formats. text/vtt carries `<v Speaker>` voice tags; the +text form carries the
// same timestamped utterances without them. Only text/vtt is selectable via the $format query
// parameter — the unattributed form must be asked for with the Accept header, which is why both go
// through Accept here.
const SPEAKER_ATTRIBUTED = "text/vtt"
const SPEAKER_UNATTRIBUTED = "application/vnd.microsoft.graph.transcript+text"

// A tenant control that took effect end of July 2026 can disable speaker-attributed transcripts.
// Asking for text/vtt in such a tenant returns 403 with this inner-error code, and the documented
// remedy is to re-request the unattributed format — so this is a retry, not an error to surface.
const SPEAKER_ATTRIBUTION_DISABLED = "SpeakerAttributionNotAllowed"

// The other 403 on this endpoint, and the one there is no request-side workaround for. Kept
// distinct so an operator is told to go change a tenant setting instead of chasing a scope grant.
const GRAPH_TRANSCRIPT_ACCESS_DISABLED = "GraphAccessToTranscriptsDisabled"

const TRANSCRIPT_SCOPE = "OnlineMeetingTranscript.Read.All"

// Nothing about the request can get past this one, so it must never carry scope advice: the caller
// may well hold the scope already, and telling them to go get it sends them after a permission that
// would change nothing.
const TENANT_ACCESS_DISABLED =
  "This tenant has Graph API access to transcripts turned off, so no transcript can be read through " +
  "this server regardless of permissions — the scope you hold is not the problem. A Teams " +
  "administrator re-enables it in the Teams Admin Center (Set-CsTeamsMeetingConfiguration)."

/**
 * How to actually obtain the transcript permission, which differs by auth mode.
 *
 * MS365_EXTRA_SCOPES is read only in oauth-proxy mode. Every other mode requests
 * `https://graph.microsoft.com/.default` and takes its scope set from the Azure app registration,
 * so naming the env var there sends an operator down a dead end — pointing at a variable that is
 * never consulted, while the real fix is a permission on the app registration.
 */
const scopeGuidance = (): string =>
  (process.env.MS365_AUTH_MODE ?? "interactive") === "oauth-proxy"
    ? `Add ${TRANSCRIPT_SCOPE} to MS365_EXTRA_SCOPES, have a tenant admin consent to it, then sign in again.`
    : `This deployment requests \`.default\`, so its scopes come from the Azure app registration and ` +
      `MS365_EXTRA_SCOPES is not consulted. Add ${TRANSCRIPT_SCOPE} as a delegated permission on the ` +
      `app registration and consent to it — tenant-wide, or for one user with an oauth2PermissionGrant ` +
      `using consentType "Principal" — then sign in again.`

// A Teams join URL carries the organizer's tenant in its context parameter, {"Tid":"…","Oid":"…"}.
// Newer short links (/meet/…) carry none, and then nothing here can tell.
export const joinUrlTenant = (joinWebUrl: string): string | undefined => {
  try {
    const context = new URL(joinWebUrl).searchParams.get("context")
    if (context === null) return undefined
    const tid = (JSON.parse(context) as { Tid?: unknown }).Tid
    return typeof tid === "string" && /^[0-9a-f-]{36}$/i.test(tid) ? tid.toLowerCase() : undefined
  } catch {
    return undefined
  }
}

// Graph won't let a delegated token read a meeting another organization hosts, nor its transcripts,
// and no permission on this side changes that. Said only when the join URL and the token prove it.
const externalMeeting = async (params: MeetingRef): Promise<string | undefined> => {
  if (!params.join_web_url) return undefined
  const organizer = joinUrlTenant(params.join_web_url)
  if (organizer === undefined) return undefined
  const own = await tokenTenant()
  if (own === undefined || own === organizer) return undefined
  return (
    `This meeting was organized in another organization (tenant ${organizer}), and Graph doesn't let a ` +
    "signed-in user read another organization's meetings or transcripts. No permission on this side changes that."
  )
}

// Scope advice only when the token is missing the scope. When it holds it, the advice would send the
// caller after a permission they already have; when its scopes can't be read, the advice stays.
const scopeAdvice = async (): Promise<string> => {
  const scopes = await tokenScopes()
  if (scopes.length === 0) return scopeGuidance()
  const held = heldScope(scopes, [TRANSCRIPT_SCOPE])
  return held === undefined ? scopeGuidance() : `The token already holds ${held}, so the permission is not the cause.`
}

type MeetingRef = {
  readonly meeting_id?: string
  readonly join_web_url?: string
}

/**
 * The meeting lookup path for a join URL.
 *
 * A joinWebUrl is a URL containing `?`, `&`, `#`, `%`-escapes and a JSON `context` blob, all of
 * which are hostile to being pasted into a query string. Two escapes, in this order:
 *
 *  1. OData string literals escape a single quote by doubling it. A joinWebUrl with an apostrophe
 *     would otherwise terminate the literal early and produce a filter parse error.
 *  2. The whole filter expression is then percent-encoded, so the URL's own `?`/`&`/`%` land as
 *     data rather than as query-string structure. Graph's own documented example shows the value
 *     double-encoded this way.
 */
export const meetingByJoinUrlPath = (joinWebUrl: string): string =>
  `/me/onlineMeetings?$filter=${encodeURIComponent(`JoinWebUrl eq '${joinWebUrl.replace(/'/g, "''")}'`)}`

const NEITHER_REF = new UserError("Provide either meeting_id or join_web_url.")

const NO_MEETING_FOUND = (joinWebUrl: string) =>
  new UserError(
    `No online meeting matches that join URL: ${joinWebUrl}\n\n` +
      "Transcripts are only available for meetings that have a calendar event — a meeting created " +
      "through the create-onlineMeeting API without one is not supported, and neither are live " +
      "events. An expired meeting also drops off this API.",
  )

/**
 * Resolve the caller's meeting reference to an onlineMeeting id.
 *
 * `meeting_id` is passed straight through, so the common case costs no extra round trip and needs
 * no OnlineMeetings.Read grant — that permission is only required for the join-URL lookup.
 */
const resolveMeetingId = async (params: MeetingRef): Promise<Either<UserError, string>> => {
  if (params.meeting_id) return Right(params.meeting_id)
  if (!params.join_web_url) return Left(NEITHER_REF)

  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.graphQuery<ODataResponse<{ id?: string }>>(
    "GET",
    meetingByJoinUrlPath(params.join_web_url),
    undefined,
    defaultVersion(),
  )
  if (result.isLeft()) {
    const message = `Failed to resolve meeting: ${(result.value as { message: string }).message}`
    const external = await externalMeeting(params)
    return Left(new UserError(external === undefined ? message : `${message}\n\n${external}`))
  }

  const id = (result.value as ODataResponse<{ id?: string }>).value[0]?.id
  return id ? Right(id) : Left(NO_MEETING_FOUND(params.join_web_url))
}

export const listMeetingTranscripts = async (params: MeetingRef): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const meetingId = await resolveMeetingId(params)
  if (meetingId.isLeft()) return meetingId as Either<UserError, string>

  const result = await client.graphQuery<ODataResponse<GraphCallTranscript>>(
    "GET",
    `/me/onlineMeetings/${encodeURIComponent(meetingId.value as string)}/transcripts`,
    undefined,
    defaultVersion(),
  )

  if (result.isLeft()) {
    return Left(
      await transcriptError(result.value as { message: string; status?: number; graphInnerErrorCode?: string }, params),
    )
  }
  return Right(formatTranscriptList((result.value as ODataResponse<GraphCallTranscript>).value))
}

const transcriptError = async (
  error: { message: string; status?: number; graphInnerErrorCode?: string },
  params: MeetingRef,
): Promise<UserError> => {
  const message = `Failed to list transcripts: ${error.message}`
  const external = await externalMeeting(params)
  if (external !== undefined) return new UserError(`${message}\n\n${external}`)
  if (error.status !== 403) return new UserError(message)

  // Both 403s here answer with outer code "Forbidden"; only the inner code separates "you lack the
  // scope" from "the tenant switched the feature off". Appending scope advice to the latter is
  // actively misleading — it was, until a live tenant with the switch off produced exactly that.
  if (error.graphInnerErrorCode === GRAPH_TRANSCRIPT_ACCESS_DISABLED) return new UserError(TENANT_ACCESS_DISABLED)

  return new UserError(`${message}\n\n${await scopeAdvice()}`)
}

type ContentAttempt = {
  readonly ok: boolean
  readonly status: number
  readonly text: string
  readonly innerErrorCode?: string
}

const fetchTranscriptContent = async (url: string, accept: string, token: string): Promise<ContentAttempt> => {
  const response = await resilientFetch(url, { headers: { Authorization: `Bearer ${token}`, Accept: accept } })
  const text = await response.text()

  if (response.ok) return { ok: true, status: response.status, text }

  // The inner-error code is what distinguishes the two 403s; the outer `code` is "Forbidden" for
  // both, so the shared error mapper cannot tell them apart.
  const innerErrorCode = ((): string | undefined => {
    try {
      return (JSON.parse(text) as { error?: { innerError?: { code?: string } } }).error?.innerError?.code
    } catch {
      return undefined
    }
  })()

  return { ok: false, status: response.status, text, innerErrorCode }
}

const contentHttpError = async (attempt: ContentAttempt, params: MeetingRef): Promise<UserError> => {
  if (attempt.status === 403 && attempt.innerErrorCode === GRAPH_TRANSCRIPT_ACCESS_DISABLED) {
    return new UserError(TENANT_ACCESS_DISABLED)
  }

  const message = ((): string => {
    try {
      return (JSON.parse(attempt.text) as { error?: { message?: string } }).error?.message ?? attempt.text
    } catch {
      return attempt.text
    }
  })()

  const external = await externalMeeting(params)
  if (external !== undefined)
    return new UserError(`Failed to get transcript (${attempt.status}): ${message}\n\n${external}`)

  if (attempt.status === 403) {
    return new UserError(`Failed to get transcript (403): ${message}\n\n${await scopeAdvice()}`)
  }

  return new UserError(`Failed to get transcript (${attempt.status}): ${message}`)
}

/**
 * Cut one window out of a transcript, starting at `offset`.
 *
 * An hour's meeting runs to about 60,000 chars of WebVTT, so the default window cuts off ordinary
 * meetings; the marker therefore carries the exact `offset` to pass next. It is printed as a raw
 * number because "50,000" cannot be pasted back as a parameter.
 */
export const sliceTranscript = (text: string, offset: number, maxChars: number): Either<UserError, string> => {
  const total = text.length
  if (offset > 0 && offset >= total) {
    return Left(new UserError(`offset ${offset} is past the end of the transcript (${total.toLocaleString()} chars)`))
  }

  // Never end a window between the two halves of an emoji (a UTF-16 surrogate pair): a lone half
  // breaks UTF-8 encoding in the client. Step forward, not back — stepping back with max_chars 1
  // would name the same offset forever.
  const cut = Math.min(offset + maxChars, total)
  const end = cut < total && (text.charCodeAt(cut - 1) & 0xfc00) === 0xd800 ? cut + 1 : cut
  const window = text.slice(offset, end)
  const range = `chars ${offset.toLocaleString()}–${end.toLocaleString()} of ${total.toLocaleString()}`

  if (end < total) return Right(`${window}\n\n[truncated: ${range}; call again with offset: ${end}]`)
  if (offset > 0) return Right(`${window}\n\n[${range}; end of transcript]`)
  return Right(window)
}

/**
 * Fetch one transcript's text.
 *
 * Uses fetch directly rather than the shared Graph client because this endpoint returns WebVTT, not
 * JSON, and because the retry below has to read the *inner* error code of a 403 — which the shared
 * error mapper discards.
 */
export const getMeetingTranscript = async (
  params: MeetingRef & {
    transcript_id: string
    include_speaker_names?: boolean
    max_chars?: number
    offset?: number
  },
): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const meetingId = await resolveMeetingId(params)
  if (meetingId.isLeft()) return meetingId as Either<UserError, string>

  const tokenResult = await getAccessToken()
  if (tokenResult.isLeft()) return Left(new UserError((tokenResult.value as { message: string }).message))
  const token = tokenResult.value as string

  const url =
    `${GRAPH_API_BASE}/${defaultVersion()}/me/onlineMeetings/${encodeURIComponent(meetingId.value as string)}` +
    `/transcripts/${encodeURIComponent(params.transcript_id)}/content`

  const wantSpeakers = params.include_speaker_names ?? true
  const first = await fetchTranscriptContent(url, wantSpeakers ? SPEAKER_ATTRIBUTED : SPEAKER_UNATTRIBUTED, token)

  // Fall back rather than surface the 403: the tenant has disabled speaker attribution, and the
  // unattributed transcript is exactly what the caller wanted minus the names.
  const attempt =
    !first.ok && first.status === 403 && first.innerErrorCode === SPEAKER_ATTRIBUTION_DISABLED
      ? await fetchTranscriptContent(url, SPEAKER_UNATTRIBUTED, token)
      : first

  if (!attempt.ok) return Left(await contentHttpError(attempt, params))

  const speakersDropped = wantSpeakers && attempt !== first
  const body = sliceTranscript(attempt.text, params.offset ?? 0, params.max_chars ?? DEFAULT_MAX_CHARS)
  if (body.isLeft()) return body

  const note = speakersDropped
    ? "\n\nNote: this tenant has speaker attribution disabled, so utterances carry timestamps but no speaker names."
    : ""

  return Right(`# Transcript ${params.transcript_id}${note}\n\n${body.value as string}`)
}
