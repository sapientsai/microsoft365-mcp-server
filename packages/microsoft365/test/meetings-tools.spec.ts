import { Some } from "functype"
import { Left, Right } from "functype/either"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

vi.mock("../src/client/graph-client", () => ({
  getGraphClient: vi.fn(),
}))

vi.mock("../src/auth", () => ({
  getAccessToken: vi.fn(),
}))

import { getAccessToken } from "../src/auth"
import { getGraphClient } from "../src/client/graph-client"
import { getMeetingTranscript, listMeetingTranscripts, meetingByJoinUrlPath } from "../src/tools/meetings-tools"

const MEETING_ID = "MSo1N2Y5ZGFjYy03MWJmLTQ3NDMtYjQxMy01M2E"
const TRANSCRIPT_ID = "MSMjMCMjNzU3ODc2ZDYtOTcwMi00MDhkLWFkNDItOTE2ZDNmZjkwZGY4"

const VTT = "WEBVTT\n\n00:00:01.000 --> 00:00:04.000\n<v Jordan Burke>Let's start."
const UNATTRIBUTED = "WEBVTT\n\n00:00:01.000 --> 00:00:04.000\nLet's start."

const mockClient = { graphQuery: vi.fn() }

/** A fetch stub that answers each call from `responses` in order. */
const stubFetch = (responses: ReadonlyArray<{ ok: boolean; status: number; body: string }>) => {
  const spy = vi.fn()
  for (const r of responses) {
    spy.mockResolvedValueOnce({ ok: r.ok, status: r.status, text: () => Promise.resolve(r.body) })
  }
  vi.stubGlobal("fetch", spy)
  return spy
}

const forbidden = (innerCode: string, message = "Forbidden.") => ({
  ok: false,
  status: 403,
  body: JSON.stringify({ error: { code: "Forbidden", message, innerError: { code: innerCode } } }),
})

const acceptHeaderOf = (spy: ReturnType<typeof vi.fn>, call: number): string =>
  (spy.mock.calls[call]?.[1] as { headers: Record<string, string> }).headers.Accept

beforeEach(() => {
  vi.clearAllMocks()
  vi.mocked(getGraphClient).mockReturnValue(Some(mockClient) as never)
  vi.mocked(getAccessToken).mockResolvedValue(Right("token-abc") as never)
})

afterEach(() => {
  vi.unstubAllGlobals()
})

describe("meetingByJoinUrlPath", () => {
  // A joinWebUrl is a URL: it carries `?`, `&`, `#`, existing %-escapes and a JSON context blob.
  // Interpolating it raw would end the query string at its own `?` and silently filter on nothing.
  it("percent-encodes the filter so the join URL's own query string is data, not structure", () => {
    const joinUrl =
      "https://teams.microsoft.com/l/meetup-join/19%3ameeting_MGQ4@thread.v2/0?context=%7b%22Tid%22%3a%22909c%22%7d"

    const path = meetingByJoinUrlPath(joinUrl)

    expect(path.startsWith("/me/onlineMeetings?$filter=")).toBe(true)
    // Exactly one `?`, the one that opens our own query string.
    expect(path.split("?")).toHaveLength(2)
    expect(path).not.toContain("&")
    // Already-escaped characters in the join URL get escaped again, as Graph's own example shows.
    expect(path).toContain("19%253ameeting_MGQ4")
    expect(decodeURIComponent(path.split("$filter=")[1] as string)).toBe(`JoinWebUrl eq '${joinUrl}'`)
  })

  // OData ends a string literal at a single quote; the escape is to double it. Without this, a join
  // URL containing an apostrophe produces a filter parse error rather than a lookup.
  it("doubles single quotes so they cannot terminate the OData literal", () => {
    const path = meetingByJoinUrlPath("https://teams.microsoft.com/l/o'brien")

    expect(decodeURIComponent(path.split("$filter=")[1] as string)).toBe(
      "JoinWebUrl eq 'https://teams.microsoft.com/l/o''brien'",
    )
  })
})

describe("listMeetingTranscripts", () => {
  it("lists transcripts for a meeting id without a lookup round trip", async () => {
    mockClient.graphQuery.mockResolvedValue(
      Right({
        value: [
          {
            id: TRANSCRIPT_ID,
            meetingId: MEETING_ID,
            createdDateTime: "2026-08-20T15:04:00Z",
            meetingOrganizer: { user: { displayName: "Jordan Burke" } },
          },
        ],
      }),
    )

    const result = await listMeetingTranscripts({ meeting_id: MEETING_ID })

    expect(result.isRight()).toBe(true)
    expect(result.value).toContain(TRANSCRIPT_ID)
    expect(result.value).toContain("Jordan Burke")
    expect(mockClient.graphQuery).toHaveBeenCalledTimes(1)
    expect(mockClient.graphQuery).toHaveBeenCalledWith(
      "GET",
      `/me/onlineMeetings/${encodeURIComponent(MEETING_ID)}/transcripts`,
      undefined,
      "v1.0",
    )
  })

  it("resolves a join URL to a meeting id first, then lists", async () => {
    mockClient.graphQuery
      .mockResolvedValueOnce(Right({ value: [{ id: MEETING_ID }] }))
      .mockResolvedValueOnce(Right({ value: [] }))

    const result = await listMeetingTranscripts({ join_web_url: "https://teams.microsoft.com/l/meetup-join/19%3ax" })

    expect(result.isRight()).toBe(true)
    expect(mockClient.graphQuery).toHaveBeenCalledTimes(2)
    expect(mockClient.graphQuery.mock.calls[1]?.[1]).toBe(
      `/me/onlineMeetings/${encodeURIComponent(MEETING_ID)}/transcripts`,
    )
  })

  it("explains the calendar-event requirement when a join URL matches nothing", async () => {
    mockClient.graphQuery.mockResolvedValue(Right({ value: [] }))

    const result = await listMeetingTranscripts({ join_web_url: "https://teams.microsoft.com/l/meetup-join/19%3ax" })

    expect(result.isLeft()).toBe(true)
    expect((result.value as Error).message).toContain("calendar event")
  })

  it("rejects a call that names neither a meeting id nor a join URL", async () => {
    const result = await listMeetingTranscripts({})

    expect(result.isLeft()).toBe(true)
    expect((result.value as Error).message).toContain("meeting_id or join_web_url")
    expect(mockClient.graphQuery).not.toHaveBeenCalled()
  })

  it("says so when a meeting has no transcripts rather than returning an empty string", async () => {
    mockClient.graphQuery.mockResolvedValue(Right({ value: [] }))

    const result = await listMeetingTranscripts({ meeting_id: MEETING_ID })

    expect(result.isRight()).toBe(true)
    expect(result.value).toContain("No transcripts found")
  })

  it("surfaces a Graph error", async () => {
    mockClient.graphQuery.mockResolvedValue(Left({ type: "forbidden", message: "Access denied", status: 403 }))

    const result = await listMeetingTranscripts({ meeting_id: MEETING_ID })

    expect(result.isLeft()).toBe(true)
    expect((result.value as Error).message).toContain("Access denied")
  })
})

describe("getMeetingTranscript", () => {
  it("returns speaker-attributed text when the tenant allows it", async () => {
    const spy = stubFetch([{ ok: true, status: 200, body: VTT }])

    const result = await getMeetingTranscript({ meeting_id: MEETING_ID, transcript_id: TRANSCRIPT_ID })

    expect(result.isRight()).toBe(true)
    expect(result.value).toContain("<v Jordan Burke>")
    expect(spy).toHaveBeenCalledTimes(1)
    expect(acceptHeaderOf(spy, 0)).toBe("text/vtt")
    expect(spy.mock.calls[0]?.[0]).toContain(
      `/me/onlineMeetings/${encodeURIComponent(MEETING_ID)}/transcripts/${encodeURIComponent(TRANSCRIPT_ID)}/content`,
    )
  })

  // The acceptance criterion: a tenant with speaker attribution disabled gets transcript text, not
  // a 403. The retry is the whole point — asserting only on the returned text would pass even if
  // the first request had succeeded, so assert the two Accept headers too.
  it("retries unattributed when the tenant disallows speaker attribution", async () => {
    const spy = stubFetch([forbidden("SpeakerAttributionNotAllowed"), { ok: true, status: 200, body: UNATTRIBUTED }])

    const result = await getMeetingTranscript({ meeting_id: MEETING_ID, transcript_id: TRANSCRIPT_ID })

    expect(result.isRight()).toBe(true)
    expect(result.value).toContain("Let's start.")
    expect(spy).toHaveBeenCalledTimes(2)
    expect(acceptHeaderOf(spy, 0)).toBe("text/vtt")
    expect(acceptHeaderOf(spy, 1)).toBe("application/vnd.microsoft.graph.transcript+text")
  })

  it("tells the caller the names are missing after falling back", async () => {
    stubFetch([forbidden("SpeakerAttributionNotAllowed"), { ok: true, status: 200, body: UNATTRIBUTED }])

    const result = await getMeetingTranscript({ meeting_id: MEETING_ID, transcript_id: TRANSCRIPT_ID })

    expect(result.value).toContain("speaker attribution disabled")
  })

  it("asks for the unattributed format directly when speaker names are not wanted", async () => {
    const spy = stubFetch([{ ok: true, status: 200, body: UNATTRIBUTED }])

    const result = await getMeetingTranscript({
      meeting_id: MEETING_ID,
      transcript_id: TRANSCRIPT_ID,
      include_speaker_names: false,
    })

    expect(result.isRight()).toBe(true)
    expect(spy).toHaveBeenCalledTimes(1)
    expect(acceptHeaderOf(spy, 0)).toBe("application/vnd.microsoft.graph.transcript+text")
    // Nothing was dropped, so there is nothing to warn about.
    expect(result.value).not.toContain("speaker attribution disabled")
  })

  // The other 403 on this endpoint has no request-side workaround, so retrying it would burn a
  // round trip and then report the wrong remedy.
  it("does not retry when Graph transcript access is disabled tenant-wide", async () => {
    const spy = stubFetch([forbidden("GraphAccessToTranscriptsDisabled")])

    const result = await getMeetingTranscript({ meeting_id: MEETING_ID, transcript_id: TRANSCRIPT_ID })

    expect(result.isLeft()).toBe(true)
    expect(spy).toHaveBeenCalledTimes(1)
    expect((result.value as Error).message).toContain("Teams Admin Center")
  })

  it("points at MS365_EXTRA_SCOPES on a plain 403", async () => {
    stubFetch([{ ok: false, status: 403, body: JSON.stringify({ error: { message: "Insufficient privileges." } }) }])

    const result = await getMeetingTranscript({ meeting_id: MEETING_ID, transcript_id: TRANSCRIPT_ID })

    expect(result.isLeft()).toBe(true)
    expect((result.value as Error).message).toContain("MS365_EXTRA_SCOPES")
    expect((result.value as Error).message).toContain("Insufficient privileges.")
  })

  it("surfaces a non-403 failure with its status", async () => {
    stubFetch([{ ok: false, status: 404, body: JSON.stringify({ error: { message: "Not found." } }) }])

    const result = await getMeetingTranscript({ meeting_id: MEETING_ID, transcript_id: TRANSCRIPT_ID })

    expect(result.isLeft()).toBe(true)
    expect((result.value as Error).message).toContain("404")
  })

  it("truncates a long transcript and says by how much", async () => {
    stubFetch([{ ok: true, status: 200, body: "x".repeat(200) }])

    const result = await getMeetingTranscript({
      meeting_id: MEETING_ID,
      transcript_id: TRANSCRIPT_ID,
      max_chars: 50,
    })

    expect(result.isRight()).toBe(true)
    expect(result.value).toContain("truncated at 50 chars")
    expect(result.value).toContain("full transcript is 200 chars")
  })

  it("resolves a join URL before fetching content", async () => {
    mockClient.graphQuery.mockResolvedValue(Right({ value: [{ id: MEETING_ID }] }))
    const spy = stubFetch([{ ok: true, status: 200, body: VTT }])

    const result = await getMeetingTranscript({
      join_web_url: "https://teams.microsoft.com/l/meetup-join/19%3ax",
      transcript_id: TRANSCRIPT_ID,
    })

    expect(result.isRight()).toBe(true)
    expect(spy.mock.calls[0]?.[0]).toContain(`/me/onlineMeetings/${encodeURIComponent(MEETING_ID)}/transcripts/`)
  })

  it("does not fetch when the meeting reference is missing", async () => {
    const spy = stubFetch([{ ok: true, status: 200, body: VTT }])

    const result = await getMeetingTranscript({ transcript_id: TRANSCRIPT_ID })

    expect(result.isLeft()).toBe(true)
    expect(spy).not.toHaveBeenCalled()
  })
})
