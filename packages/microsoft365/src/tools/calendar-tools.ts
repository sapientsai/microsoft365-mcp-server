import type { Either } from "functype/either"
import { Left, Right } from "functype/either"
import { UserError } from "somamcp"

import { getGraphClient } from "../client/graph-client"
import type { GraphCalendar, ODataResponse } from "../types"
import {
  formatCalendarList,
  formatEventDetail,
  formatEventList,
  formatMeetingTimeSuggestions,
} from "../utils/formatters"
import { isGraphId } from "../utils/graph-id"

const requireClient = () => {
  const client = getGraphClient()
  if (client.isNone()) return null
  return client.orThrow()
}

const invalidCalendarId = (calendarId: string | undefined): UserError | undefined =>
  calendarId === undefined || isGraphId(calendarId)
    ? undefined
    : new UserError("calendar_id is not a calendar ID. Use an ID exactly as list_calendars returns it.")

// time_zone goes into a Prefer header inside quotes, so only the characters zone names use get
// through: IANA (America/New_York, Etc/GMT+5) and Windows (Eastern Standard Time, UTC-02).
const ZONE_NAME = /^[A-Za-z0-9_+\-/ .()]{1,100}$/

const invalidTimeZone = (timeZone: string | undefined): UserError | undefined =>
  timeZone === undefined || ZONE_NAME.test(timeZone)
    ? undefined
    : new UserError(
        "time_zone is not a time zone name. Use an IANA name such as America/New_York, or a Windows one such as " +
          "Eastern Standard Time.",
      )

// Graph's refusal of a write to someone else's read-only calendar reads like a bad sign-in, so the
// calendar is read first and the write refused here with the real reason. If that read fails, the
// write goes ahead: Graph still enforces the rule, and a guard that blocks on its own failure would
// stop writes that would have worked. The default calendar is always the user's own, so it isn't read.
const readOnlyCalendar = async (
  client: NonNullable<ReturnType<typeof requireClient>>,
  calendarId: string | undefined,
): Promise<UserError | undefined> => {
  if (calendarId === undefined) return undefined
  const calendar = await client.getCalendar(calendarId)
  if (calendar.isLeft()) return undefined
  const { name, canEdit } = calendar.value as GraphCalendar
  if (canEdit !== false) return undefined
  const label = (name ?? "This calendar").replace(/[\r\n]+/g, " ").slice(0, 120)
  return new UserError(
    `"${label}" is read-only for you, so nothing was changed. Its owner has shared it without edit rights; ` +
      "list_calendars marks such calendars [read-only].",
  )
}

const invalidEventId = (eventId: string): UserError | undefined =>
  isGraphId(eventId)
    ? undefined
    : new UserError("event_id is not an event ID. Use an ID exactly as list_events or list_calendar_view returns it.")

export const listCalendars = async (): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.requestPaginated<GraphCalendar>("/me/calendars", {
    odataParams: { $select: ["id", "name", "isDefaultCalendar", "canEdit", "owner"] },
  })
  return result.mapLeft((error) => new UserError(`Failed to list calendars: ${error.message}`)).map(formatCalendarList)
}

export const listEvents = async (params: {
  top?: number
  filter?: string
  fetch_all_pages?: boolean
  calendar_id?: string
  time_zone?: string
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))
  const badCalendar = invalidCalendarId(params.calendar_id)
  if (badCalendar) return Left(badCalendar)
  const badZone = invalidTimeZone(params.time_zone)
  if (badZone) return Left(badZone)

  if (params.fetch_all_pages) {
    const result = await client.listAllEvents(
      { $filter: params.filter, $orderby: "start/dateTime" },
      params.calendar_id,
      params.time_zone,
    )
    return result
      .mapLeft((error) => new UserError(`Failed to list events: ${error.message}`))
      .map((items) => formatEventList(items))
  }

  const result = await client.listEvents(
    {
      $top: params.top ?? 25,
      $filter: params.filter,
      $orderby: "start/dateTime",
    },
    params.calendar_id,
    params.time_zone,
  )
  return result
    .mapLeft((error) => new UserError(`Failed to list events: ${error.message}`))
    .map((response) => formatEventList((response as ODataResponse<never>).value))
}

export const getEvent = async (params: {
  event_id: string
  body_format?: "text" | "html"
  calendar_id?: string
  time_zone?: string
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))
  const badCalendar = invalidCalendarId(params.calendar_id)
  if (badCalendar) return Left(badCalendar)
  const badEvent = invalidEventId(params.event_id)
  if (badEvent) return Left(badEvent)
  const badZone = invalidTimeZone(params.time_zone)
  if (badZone) return Left(badZone)

  // Text by default, unlike get_message: an event body is almost always a Teams invite, whose HTML
  // is layout markup around a join link and the same details the event already carries.
  const result = await client.getEvent(
    params.event_id,
    params.body_format ?? "text",
    params.calendar_id,
    params.time_zone,
  )
  return result.mapLeft((error) => new UserError(`Failed to get event: ${error.message}`)).map(formatEventDetail)
}

export const listCalendarView = async (params: {
  start_date_time: string
  end_date_time: string
  top?: number
  calendar_id?: string
  time_zone?: string
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))
  const badCalendar = invalidCalendarId(params.calendar_id)
  if (badCalendar) return Left(badCalendar)
  const badZone = invalidTimeZone(params.time_zone)
  if (badZone) return Left(badZone)

  if (!params.start_date_time || !params.end_date_time) {
    return Left(new UserError("start_date_time and end_date_time are required (ISO 8601, e.g. 2026-05-22T00:00:00Z)."))
  }

  const result = await client.listCalendarView(
    params.start_date_time,
    params.end_date_time,
    {
      $top: params.top ?? 50,
      $orderby: "start/dateTime",
    },
    params.calendar_id,
    params.time_zone,
  )
  return result
    .mapLeft((error) => new UserError(`Failed to list calendar view: ${error.message}`))
    .map((response) => formatEventList((response as ODataResponse<never>).value))
}

// Graph wants an all-day event to run midnight to midnight in its own time zone, ending at the
// start of the day after the last one. Callers give the last day itself, as a person would say
// "6 to 8 November", so the exclusive end is worked out here rather than left to the caller.
const ALL_DAY_DATE = /^(\d{4}-\d{2}-\d{2})(?:T00:00(?::00(?:\.0+)?)?)?$/

// Date.parse accepts 2026-02-30 and rolls it into March, so a date is real only if it survives
// the round trip unchanged.
const isRealDate = (date: string): boolean => {
  const d = new Date(`${date}T00:00:00Z`)
  return !Number.isNaN(d.getTime()) && d.toISOString().slice(0, 10) === date
}

const nextDay = (date: string): string => {
  const d = new Date(`${date}T00:00:00Z`)
  d.setUTCDate(d.getUTCDate() + 1)
  return d.toISOString().slice(0, 10)
}

const allDayRange = (
  start: string,
  lastDay: string | undefined,
): { readonly start: string; readonly end: string } | UserError => {
  const first = ALL_DAY_DATE.exec(start)?.[1]
  const last = ALL_DAY_DATE.exec(lastDay ?? start)?.[1]
  if (!first || !last || !isRealDate(first) || !isRealDate(last))
    return new UserError(
      "For an all-day event, start and end are dates (YYYY-MM-DD), with no time of day or time zone.",
    )
  if (last < first) return new UserError("For an all-day event, end (the last day) cannot be before start.")
  return { start: `${first}T00:00:00`, end: `${nextDay(last)}T00:00:00` }
}

export const SHOW_AS = ["free", "tentative", "busy", "oof", "workingElsewhere"] as const
export type ShowAs = (typeof SHOW_AS)[number]

// Graph marks every new event busy. Outlook marks an all-day event free, and back to busy when it
// becomes timed again; a staff day or school carnival marked busy would block the whole day in
// everyone's free/busy view. Follow Outlook unless the caller says otherwise.
const showAsFor = (showAs: ShowAs | undefined, allDay: boolean | undefined): ShowAs | undefined =>
  showAs ?? (allDay === true ? "free" : allDay === false ? "busy" : undefined)

// all_day also has to be passed to move an all-day event, because that is what works out the
// exclusive end. So only a real change of state takes the Outlook default: moving a vacation marked
// oof must not turn it free. Knowing which it is takes one read, and only when show_as is left out.
const showAsOnUpdate = async (
  client: NonNullable<ReturnType<typeof requireClient>>,
  params: { event_id: string; all_day?: boolean; show_as?: ShowAs; calendar_id?: string },
): Promise<Either<UserError, ShowAs | undefined>> => {
  if (params.show_as !== undefined || params.all_day === undefined) return Right(params.show_as)
  const current = await client.getEvent(params.event_id, "text", params.calendar_id)
  return current
    .mapLeft((error) => new UserError(`Failed to read the event before updating it: ${error.message}`))
    .map((event) => ((event.isAllDay ?? false) === params.all_day ? undefined : showAsFor(undefined, params.all_day)))
}

// Graph gives every new event a 15-minute reminder, which for an all-day event fires at 23:45 the
// night before. Outlook's own default comes from a mailbox setting the API cannot read, so the
// caller sets it: reminder_minutes turns the reminder on with that lead time, reminder: false
// turns it off. Left out, Graph's default (on create) or the event's current reminder (on update)
// stays.
const reminderFields = (
  reminder: boolean | undefined,
  minutes: number | undefined,
): Record<string, unknown> | UserError => {
  if (minutes !== undefined) {
    if (reminder === false) return new UserError("reminder: false cannot be combined with reminder_minutes.")
    return { isReminderOn: true, reminderMinutesBeforeStart: minutes }
  }
  return reminder === undefined ? {} : { isReminderOn: reminder }
}

export const createEvent = async (params: {
  subject: string
  start: string
  end?: string
  all_day?: boolean
  show_as?: ShowAs
  reminder?: boolean
  reminder_minutes?: number
  time_zone?: string
  location?: string
  body?: string
  content_type?: string
  attendees?: string
  is_draft?: boolean
  online_meeting?: boolean
  calendar_id?: string
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))
  const badCalendar = invalidCalendarId(params.calendar_id)
  if (badCalendar) return Left(badCalendar)
  const badZone = invalidTimeZone(params.time_zone)
  if (badZone) return Left(badZone)

  const timeZone = params.time_zone ?? "UTC"
  const range = params.all_day
    ? allDayRange(params.start, params.end)
    : params.end
      ? { start: params.start, end: params.end }
      : new UserError("end is required unless all_day is true.")
  if (range instanceof UserError) return Left(range)
  const event: Record<string, unknown> = {
    subject: params.subject,
    start: { dateTime: range.start, timeZone },
    end: { dateTime: range.end, timeZone },
  }
  if (params.all_day) event.isAllDay = true
  const showAs = showAsFor(params.show_as, params.all_day ? true : undefined)
  if (showAs) event.showAs = showAs
  const reminder = reminderFields(params.reminder, params.reminder_minutes)
  if (reminder instanceof UserError) return Left(reminder)
  Object.assign(event, reminder)

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

  const readOnly = await readOnlyCalendar(client, params.calendar_id)
  if (readOnly) return Left(readOnly)

  const result = await client.createEvent(event, params.calendar_id, params.time_zone)
  const draftSuffix =
    params.is_draft && params.attendees
      ? "\n\n_Saved as draft: attendees were not invited. Open in Outlook to add attendees and send the meeting._"
      : ""
  return result
    .mapLeft((error) => new UserError(`Failed to create event: ${error.message}`))
    .map((evt) => `Event created.\n\n${formatEventDetail(evt, { body: false })}${draftSuffix}`)
}

export const updateEvent = async (params: {
  event_id: string
  subject?: string
  start?: string
  end?: string
  all_day?: boolean
  show_as?: ShowAs
  reminder?: boolean
  reminder_minutes?: number
  time_zone?: string
  location?: string
  body?: string
  content_type?: string
  attendees?: string
  online_meeting?: boolean
  calendar_id?: string
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))
  const badCalendar = invalidCalendarId(params.calendar_id)
  if (badCalendar) return Left(badCalendar)
  const badEvent = invalidEventId(params.event_id)
  if (badEvent) return Left(badEvent)
  const badZone = invalidTimeZone(params.time_zone)
  if (badZone) return Left(badZone)

  const updates: Record<string, unknown> = {}
  const timeZone = params.time_zone ?? "UTC"

  if (params.subject) updates.subject = params.subject
  if (params.all_day === true) {
    // The event's current dates are not known here, so an all-day change has to name its days.
    if (!params.start) return Left(new UserError("all_day: true needs start (the first day, YYYY-MM-DD)."))
    const allDay = allDayRange(params.start, params.end)
    if (allDay instanceof UserError) return Left(allDay)
    updates.isAllDay = true
    updates.start = { dateTime: allDay.start, timeZone }
    updates.end = { dateTime: allDay.end, timeZone }
  } else {
    if (params.all_day === false) {
      // A now-timed event needs its own times, or it would keep its old midnight-to-midnight span.
      if (!params.start || !params.end)
        return Left(new UserError("all_day: false needs both start and end times for the now-timed event."))
      updates.isAllDay = false
    }
    if (params.start) updates.start = { dateTime: params.start, timeZone }
    if (params.end) updates.end = { dateTime: params.end, timeZone }
  }
  const reminder = reminderFields(params.reminder, params.reminder_minutes)
  if (reminder instanceof UserError) return Left(reminder)
  Object.assign(updates, reminder)
  const showAs = await showAsOnUpdate(client, params)
  if (showAs.isLeft()) return Left(showAs.value as UserError)
  if (showAs.value) updates.showAs = showAs.value
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

  const readOnly = await readOnlyCalendar(client, params.calendar_id)
  if (readOnly) return Left(readOnly)

  const result = await client.updateEvent(params.event_id, updates, params.calendar_id, params.time_zone)
  return result
    .mapLeft((error) => new UserError(`Failed to update event: ${error.message}`))
    .map((evt) => `Event updated.\n\n${formatEventDetail(evt, { body: false })}`)
}

export const deleteEvent = async (params: {
  event_id: string
  calendar_id?: string
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))
  const badCalendar = invalidCalendarId(params.calendar_id)
  if (badCalendar) return Left(badCalendar)
  const badEvent = invalidEventId(params.event_id)
  if (badEvent) return Left(badEvent)

  const readOnly = await readOnlyCalendar(client, params.calendar_id)
  if (readOnly) return Left(readOnly)

  const result = await client.deleteEvent(params.event_id, params.calendar_id)
  return result
    .mapLeft((error) => new UserError(`Failed to delete event: ${error.message}`))
    .map(() => "Event deleted successfully.")
}

const EVENT_RESPONSE_ACTIONS = {
  accept: "accept",
  decline: "decline",
  tentative: "tentativelyAccept",
} as const

const EVENT_RESPONSE_WORDS = {
  accept: "Accepted",
  decline: "Declined",
  tentative: "Tentatively accepted",
} as const

export const respondToEvent = async (params: {
  event_id: string
  response: keyof typeof EVENT_RESPONSE_ACTIONS
  comment?: string
  send_response?: boolean
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))
  const badEvent = invalidEventId(params.event_id)
  if (badEvent) return Left(badEvent)

  const sendResponse = params.send_response ?? true
  const body = params.comment === undefined ? { sendResponse } : { comment: params.comment, sendResponse }
  const result = await client.respondToEvent(params.event_id, EVENT_RESPONSE_ACTIONS[params.response], body)
  return result
    .mapLeft((error) => new UserError(`Failed to respond to event: ${error.message}`))
    .map(
      () =>
        `${EVENT_RESPONSE_WORDS[params.response]}. ${
          sendResponse ? "A response was sent to the organizer." : "No response was sent to the organizer."
        }`,
    )
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

  if (!params.participants.length) {
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
