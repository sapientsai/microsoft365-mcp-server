import { UserError } from "fastmcp"
import type { Either } from "functype/either"
import { Left } from "functype/either"

import { getGraphClient } from "../client/graph-client"
import type { GraphEvent, ODataResponse } from "../types"
import { formatEventDetail, formatEventList } from "../utils/formatters"

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

export const createEvent = async (params: {
  subject: string
  start: string
  end: string
  time_zone?: string
  location?: string
  body?: string
  attendees?: string
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
  if (params.body) event.body = { contentType: "Text", content: params.body }
  if (params.attendees) {
    event.attendees = params.attendees.split(",").map((email) => ({
      emailAddress: { address: email.trim() },
      type: "required",
    }))
  }

  const result = await client.createEvent(event)
  return result
    .mapLeft((error) => new UserError(`Failed to create event: ${error.message}`))
    .map((evt) => `Event created.\n\n${formatEventDetail(evt)}`)
}

export const updateEvent = async (params: {
  event_id: string
  subject?: string
  start?: string
  end?: string
  time_zone?: string
  location?: string
  body?: string
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const updates: Record<string, unknown> = {}
  const timeZone = params.time_zone ?? "UTC"

  if (params.subject) updates.subject = params.subject
  if (params.start) updates.start = { dateTime: params.start, timeZone }
  if (params.end) updates.end = { dateTime: params.end, timeZone }
  if (params.location) updates.location = { displayName: params.location }
  if (params.body) updates.body = { contentType: "Text", content: params.body }

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
