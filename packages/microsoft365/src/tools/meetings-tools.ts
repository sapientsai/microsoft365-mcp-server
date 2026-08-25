import { UserError } from "fastmcp"
import type { Either } from "functype/either"
import { Left, Right } from "functype/either"

import { getAccessToken } from "../auth"
import { GRAPH_API_BASE } from "../auth/scopes"
import { getGraphClient } from "../client/graph-client"
import type { GraphCallTranscript, ODataResponse } from "../types"
import { formatTranscriptList } from "../utils/formatters"

const requireClient = () => {
  const client = getGraphClient()
  if (client.isNone()) return null
  return client.orThrow()
}

const VERSION = "v1.0"
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
    VERSION,
  )
  if (result.isLeft()) {
    return Left(new UserError(`Failed to resolve meeting: ${(result.value as { message: string }).message}`))
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
    VERSION,
  )

  return result
    .mapLeft((error) => transcriptError(error as { message: string; graphErrorCode?: string }))
    .map((response) => formatTranscriptList((response as ODataResponse<GraphCallTranscript>).value))
}

const transcriptError = (error: { message: string; graphErrorCode?: string }): UserError =>
  new UserError(`Failed to list transcripts: ${error.message}`)

type ContentAttempt = {
  readonly ok: boolean
  readonly status: number
  readonly text: string
  readonly innerErrorCode?: string
}

const fetchTranscriptContent = async (url: string, accept: string, token: string): Promise<ContentAttempt> => {
  const response = await fetch(url, { headers: { Authorization: `Bearer ${token}`, Accept: accept } })
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

const contentHttpError = (attempt: ContentAttempt): UserError => {
  if (attempt.status === 403 && attempt.innerErrorCode === GRAPH_TRANSCRIPT_ACCESS_DISABLED) {
    return new UserError(
      "This tenant has Graph API access to transcripts turned off, so no transcript can be read " +
        "through this server regardless of permissions. A Teams administrator changes it in the " +
        "Teams Admin Center (Set-CsTeamsMeetingConfiguration).",
    )
  }

  const message = ((): string => {
    try {
      return (JSON.parse(attempt.text) as { error?: { message?: string } }).error?.message ?? attempt.text
    } catch {
      return attempt.text
    }
  })()

  if (attempt.status === 403) {
    return new UserError(
      `Failed to get transcript (403): ${message}\n\n` +
        "If this deployment has never been granted OnlineMeetingTranscript.Read.All, add it to " +
        "MS365_EXTRA_SCOPES and have a tenant admin consent, then sign in again.",
    )
  }

  return new UserError(`Failed to get transcript (${attempt.status}): ${message}`)
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
    `${GRAPH_API_BASE}/${VERSION}/me/onlineMeetings/${encodeURIComponent(meetingId.value as string)}` +
    `/transcripts/${encodeURIComponent(params.transcript_id)}/content`

  const wantSpeakers = params.include_speaker_names ?? true
  const first = await fetchTranscriptContent(url, wantSpeakers ? SPEAKER_ATTRIBUTED : SPEAKER_UNATTRIBUTED, token)

  // Fall back rather than surface the 403: the tenant has disabled speaker attribution, and the
  // unattributed transcript is exactly what the caller wanted minus the names.
  const attempt =
    !first.ok && first.status === 403 && first.innerErrorCode === SPEAKER_ATTRIBUTION_DISABLED
      ? await fetchTranscriptContent(url, SPEAKER_UNATTRIBUTED, token)
      : first

  if (!attempt.ok) return Left(contentHttpError(attempt))

  const speakersDropped = wantSpeakers && attempt !== first
  const maxChars = params.max_chars ?? DEFAULT_MAX_CHARS
  const body =
    attempt.text.length > maxChars
      ? `${attempt.text.slice(0, maxChars)}\n\n[truncated at ${maxChars.toLocaleString()} chars — full transcript is ${attempt.text.length.toLocaleString()} chars]`
      : attempt.text

  const note = speakersDropped
    ? "\n\nNote: this tenant has speaker attribution disabled, so utterances carry timestamps but no speaker names."
    : ""

  return Right(`# Transcript ${params.transcript_id}${note}\n\n${body}`)
}
