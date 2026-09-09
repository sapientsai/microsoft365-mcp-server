// Calendar tool definitions.

import { z } from "zod"

import {
  createEvent,
  deleteEvent,
  findMeetingAvailability,
  getEvent,
  listCalendarView,
  listEvents,
  updateEvent,
} from ".."
import type { ToolDefinition } from "../tool-definitions"
import { FETCH_ALL_PAGES_PARAM, unwrapResult } from "./shared"

export const calendarTools: ReadonlyArray<ToolDefinition> = [
  {
    name: "list_events",
    description:
      "List calendar event resources (/me/events). Returns series masters for recurring meetings, NOT individual occurrences. " +
      "For 'what's on my calendar between X and Y' use list_calendar_view instead — it expands recurrences into instances.",
    parameters: z.object({
      top: z.number().optional().describe("Number of events to return (default: 25)"),
      filter: z.string().optional().describe("OData filter expression"),
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listEvents(params)),
    domain: "calendar",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "list_calendar_view",
    description:
      "List event instances on the calendar between start_date_time and end_date_time. Expands recurring series into " +
      "individual occurrences (unlike list_events which returns series masters). Use this for 'what's on my calendar this week'.",
    parameters: z.object({
      start_date_time: z.string().describe("Window start (ISO 8601, e.g. 2026-05-22T00:00:00Z)"),
      end_date_time: z.string().describe("Window end (ISO 8601, e.g. 2026-05-29T00:00:00Z)"),
      top: z.number().optional().describe("Max events to return (default: 50)"),
    }),
    execute: async (params) => unwrapResult(await listCalendarView(params)),
    domain: "calendar",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "find_meeting_availability",
    description:
      "Find meeting time suggestions where all participants are free, ranked by confidence. Searches between " +
      "after_date_time and before_date_time for a slot of duration_minutes. Read-only: suggests times only — " +
      "use create_event to book one.",
    parameters: z.object({
      participants: z
        .array(z.string())
        .min(1)
        .describe("Attendee email addresses. The signed-in user is automatically the organizer."),
      after_date_time: z.string().describe("Search window start (ISO 8601, e.g. 2026-06-04T00:00:00Z)"),
      before_date_time: z.string().describe("Search window end (ISO 8601, e.g. 2026-06-06T00:00:00Z)"),
      duration_minutes: z
        .number()
        .int()
        .min(15)
        .max(480)
        .optional()
        .describe("Meeting length in minutes (default: 30)"),
      max_candidates: z.number().int().min(1).max(50).optional().describe("Max slots to return (default: 3)"),
      is_organizer_optional: z
        .boolean()
        .optional()
        .describe("Whether the organizer's attendance is optional (default: false)"),
    }),
    execute: async (params) => unwrapResult(await findMeetingAvailability(params)),
    domain: "calendar",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "get_event",
    description: "Get detailed information about a calendar event",
    parameters: z.object({
      event_id: z.string().describe("The event ID"),
    }),
    execute: async (params) => unwrapResult(await getEvent(params)),
    domain: "calendar",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "create_event",
    description: "Create a new calendar event",
    parameters: z.object({
      subject: z.string().describe("Event subject/title"),
      start: z.string().describe("Start date/time (ISO format)"),
      end: z.string().describe("End date/time (ISO format)"),
      time_zone: z.string().optional().describe("Time zone (default: UTC)"),
      location: z.string().optional().describe("Event location"),
      body: z.string().optional().describe("Event description"),
      content_type: z.string().optional().describe("Body content type: Text or HTML (default: Text)"),
      attendees: z.string().optional().describe("Comma-separated email addresses of attendees"),
      is_draft: z
        .boolean()
        .optional()
        .describe(
          "Save to your calendar without inviting attendees (any attendees param is ignored). " +
            "Add attendees and send the meeting later from Outlook. Default: false.",
        ),
      online_meeting: z
        .boolean()
        .optional()
        .describe("Add a Teams meeting to the event; joinUrl is returned in the event details (default: false)"),
    }),
    execute: async (params) => unwrapResult(await createEvent(params)),
    domain: "calendar",
    readOnly: false,
  },
  {
    name: "update_event",
    description: "Update an existing calendar event",
    parameters: z.object({
      event_id: z.string().describe("The event ID to update"),
      subject: z.string().optional().describe("New subject"),
      start: z.string().optional().describe("New start date/time (ISO format)"),
      end: z.string().optional().describe("New end date/time (ISO format)"),
      time_zone: z.string().optional().describe("Time zone (default: UTC)"),
      location: z.string().optional().describe("New location"),
      body: z.string().optional().describe("New description"),
      content_type: z.string().optional().describe("Body content type: Text or HTML (default: Text)"),
      attendees: z
        .string()
        .optional()
        .describe("Comma-separated email addresses; replaces the current attendee list when provided"),
      online_meeting: z
        .boolean()
        .optional()
        .describe(
          "Add a Teams meeting to the event. One-way per Graph: once enabled it cannot be turned off or re-provisioned. Default: false.",
        ),
    }),
    execute: async (params) => unwrapResult(await updateEvent(params)),
    domain: "calendar",
    readOnly: false,
  },
  {
    name: "delete_event",
    description: "Delete a calendar event",
    parameters: z.object({
      event_id: z.string().describe("The event ID to delete"),
    }),
    execute: async (params) => unwrapResult(await deleteEvent(params)),
    domain: "calendar",
    readOnly: false,
    annotations: { destructiveHint: true },
  },
]
