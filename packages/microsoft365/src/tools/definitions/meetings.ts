// Meeting transcript tool definitions.

import { z } from "zod"

import { getMeetingTranscript, listMeetingTranscripts } from ".."
import type { ToolDefinition } from "../tool-definitions"
import { unwrapResult } from "./shared"

export const meetingsTools: ReadonlyArray<ToolDefinition> = [
  //
  // Gated by admin-consent scopes that are deliberately not in the defaults, so on a deployment
  // that has not opted in via MS365_EXTRA_SCOPES these are visible but return a 403 explaining
  // what to grant. See the README's "Meeting transcripts" section.
  {
    name: "list_meeting_transcripts",
    description:
      "List the transcripts of a Teams meeting. Identify the meeting by meeting_id, or by " +
      "join_web_url taken from a calendar event's onlineMeeting.joinUrl. Returns transcript IDs " +
      "for get_meeting_transcript.",
    parameters: z.object({
      meeting_id: z.string().optional().describe("Online meeting ID. Preferred — needs no lookup."),
      join_web_url: z
        .string()
        .optional()
        .describe("Teams join URL (an event's onlineMeeting.joinUrl). Requires the OnlineMeetings.Read scope."),
    }),
    execute: async (params) => unwrapResult(await listMeetingTranscripts(params)),
    domain: "meetings",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "get_meeting_transcript",
    description:
      "Get the text of one Teams meeting transcript, as timestamped WebVTT utterances. Works for " +
      "attendees, not just the organizer. Falls back to a speaker-less transcript automatically " +
      "when the tenant has speaker attribution disabled.",
    parameters: z.object({
      transcript_id: z.string().describe("Transcript ID from list_meeting_transcripts"),
      meeting_id: z.string().optional().describe("Online meeting ID. Preferred — needs no lookup."),
      join_web_url: z
        .string()
        .optional()
        .describe("Teams join URL, as an alternative to meeting_id. Requires the OnlineMeetings.Read scope."),
      include_speaker_names: z
        .boolean()
        .optional()
        .describe("Ask for speaker-attributed text (default: true). Ignored if the tenant disallows it."),
      max_chars: z.number().optional().describe("Truncate the transcript at this length (default: 50000)"),
    }),
    execute: async (params) => unwrapResult(await getMeetingTranscript(params)),
    domain: "meetings",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
]
