import { UserError } from "fastmcp"
import type { Either } from "functype/either"
import { Left } from "functype/either"

import { getGraphClient } from "../client/graph-client"
import type { GraphEvent, ODataResponse } from "../types"
import { formatEventDetail, formatEventList, formatMeetingTimeSuggestions } from "../utils/formatters"

const requireClient = () => {
  const client = getGraphClient()
  if (client.isNone()) return null
  return client.orThrow()
}

export const listEvents = async (params: {
  top?: number
  filter?: string
  fetch_all_pages?: boolean
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  if (params.fetch_all_pages) {
    const result = await client.requestPaginated<GraphEvent>("/me/events", {
      odataParams: { $filter: params.filter, $orderby: "start/dateTime" },
    })
    return result
      .mapLeft((error) => new UserError(`Failed to list events: ${error.message}`))
      .map((items) => formatEventList(items))
  }

  const result = await client.listEvents({
    $top: params.top ?? 25,
    $filter: params.filter,
    $orderby: "start/dateTime",
  })
  return result
    .mapLeft((error) => new UserError(`Failed to list events: ${error.message}`))
    .map((response) => formatEventList((response as ODataResponse<never>).value))
}

export const getEvent = async (params: { event_id: string }): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.getEvent(params.event_id)
  return result.mapLeft((error) => new UserError(`Failed to get event: ${error.message}`)).map(formatEventDetail)
}

export const listCalendarView = async (params: {
  start_date_time: string
  end_date_time: string
  top?: number
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  if (!params.start_date_time || !params.end_date_time) {
    return Left(new UserError("start_date_time and end_date_time are required (ISO 8601, e.g. 2026-05-22T00:00:00Z)."))
  }

  const result = await client.listCalendarView(params.start_date_time, params.end_date_time, {
    $top: params.top ?? 50,
    $orderby: "start/dateTime",
  })
  return result
    .mapLeft((error) => new UserError(`Failed to list calendar view: ${error.message}`))
    .map((response) => formatEventList((response as ODataResponse<never>).value))
}

export const createEvent = async (params: {
  subject: string
  start: string
  end: string
  time_zone?: string
  location?: string
  body?: string
  content_type?: string
  attendees?: string
  is_draft?: boolean
  online_meeting?: boolean
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const timeZone = params.time_zone ?? "UTC"
  const event: Record<string, unknown> = {
    subject: params.subject,
    start: { dateTime: params.start, timeZone },
    end: { dateTime: params.end, timeZone },
  }

  if (params.location) event.location = { displayName: params.location }
  if (params.body) event.body = { contentType: params.content_type ?? "Text", content: params.body }
  if (params.online_meeting) {
    event.isOnlineMeeting = true
    event.onlineMeetingProvider = "teamsForBusiness"
  }
  // When is_draft=true, omit attendees so Graph does not send invitations on POST.
  // The caller can add attendees later in Outlook and send the meeting from there;
  // Graph has no API to dispatch deferred invites and `isDraft` is read-only.
  if (params.attendees && !params.is_draft) {
    event.attendees = params.attendees.split(",").map((email) => ({
      emailAddress: { address: email.trim() },
      type: "required",
    }))
  }

  const result = await client.createEvent(event)
  const draftSuffix =
    params.is_draft && params.attendees
      ? "\n\n_Saved as draft: attendees were not invited. Open in Outlook to add attendees and send the meeting._"
      : ""
  return result
    .mapLeft((error) => new UserError(`Failed to create event: ${error.message}`))
    .map((evt) => `Event created.\n\n${formatEventDetail(evt)}${draftSuffix}`)
}

export const updateEvent = async (params: {
  event_id: string
  subject?: string
  start?: string
  end?: string
  time_zone?: string
  location?: string
  body?: string
  content_type?: string
  attendees?: string
  online_meeting?: boolean
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const updates: Record<string, unknown> = {}
  const timeZone = params.time_zone ?? "UTC"

  if (params.subject) updates.subject = params.subject
  if (params.start) updates.start = { dateTime: params.start, timeZone }
  if (params.end) updates.end = { dateTime: params.end, timeZone }
  if (params.location) updates.location = { displayName: params.location }
  if (params.body) updates.body = { contentType: params.content_type ?? "Text", content: params.body }
  if (params.attendees) {
    updates.attendees = params.attendees.split(",").map((email) => ({
      emailAddress: { address: email.trim() },
      type: "required",
    }))
  }
  // Graph constraint: once isOnlineMeeting is true, it cannot be set false or the provider changed.
  // We only honor online_meeting=true; explicit false is a no-op (omitted from PATCH).
  if (params.online_meeting) {
    updates.isOnlineMeeting = true
    updates.onlineMeetingProvider = "teamsForBusiness"
  }

  const result = await client.updateEvent(params.event_id, updates)
  return result
    .mapLeft((error) => new UserError(`Failed to update event: ${error.message}`))
    .map((evt) => `Event updated.\n\n${formatEventDetail(evt)}`)
}

export const deleteEvent = async (params: { event_id: string }): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.deleteEvent(params.event_id)
  return result
    .mapLeft((error) => new UserError(`Failed to delete event: ${error.message}`))
    .map(() => "Event deleted successfully.")
}

export const findMeetingAvailability = async (params: {
  participants: string[]
  after_date_time: string
  before_date_time: string
  duration_minutes?: number
  max_candidates?: number
  is_organizer_optional?: boolean
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  if (!params.participants?.length) {
    return Left(new UserError("participants is required (at least one attendee email)."))
  }
  if (!params.after_date_time || !params.before_date_time) {
    return Left(
      new UserError("after_date_time and before_date_time are required (ISO 8601, e.g. 2026-06-04T00:00:00Z)."),
    )
  }
  if (params.before_date_time <= params.after_date_time) {
    return Left(new UserError("before_date_time must be after after_date_time."))
  }

  const body = {
    attendees: params.participants.map((address) => ({
      type: "required",
      emailAddress: { address: address.trim() },
    })),
    timeConstraint: {
      activityDomain: "work",
      timeSlots: [
        {
          start: { dateTime: params.after_date_time, timeZone: "UTC" },
          end: { dateTime: params.before_date_time, timeZone: "UTC" },
        },
      ],
    },
    // ISO 8601 duration (PT30M), NOT minutes — easy to get wrong.
    meetingDuration: `PT${params.duration_minutes ?? 30}M`,
    maxCandidates: params.max_candidates ?? 3,
    isOrganizerOptional: params.is_organizer_optional ?? false,
    // 100 => only return slots where every attendee is free (mirrors the generic connector).
    minimumAttendeePercentage: 100,
    returnSuggestionReasons: true,
  }

  const result = await client.findMeetingTimes(body)
  return result
    .mapLeft((error) => new UserError(`Failed to find meeting times: ${error.message}`))
    .map((res) => formatMeetingTimeSuggestions(res))
}
