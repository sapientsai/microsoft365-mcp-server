// Builds a Graph patternedRecurrence from the flat parameters a tool exposes.
//
// Graph is strict here in a way that is easy to get wrong: "Any property that you
// include that does not have a supported value would result in an error." Sending
// dayOfMonth on a weekly pattern fails the whole create, so each pattern type is
// assembled from only the properties that type allows rather than by spreading
// every optional the caller happened to pass.

import { Option } from "functype"

import type { GraphPatternedRecurrence, GraphRecurrencePattern, GraphRecurrenceRange } from "../types"

export const RECURRENCE_PATTERN_TYPES = [
  "daily",
  "weekly",
  "absoluteMonthly",
  "relativeMonthly",
  "absoluteYearly",
  "relativeYearly",
] as const

export const RECURRENCE_RANGE_TYPES = ["noEnd", "endDate", "numbered"] as const

export const DAYS_OF_WEEK = ["sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"] as const

export const WEEK_INDEXES = ["first", "second", "third", "fourth", "last"] as const

export type RecurrencePatternType = (typeof RECURRENCE_PATTERN_TYPES)[number]
export type RecurrenceRangeType = (typeof RECURRENCE_RANGE_TYPES)[number]

export type RecurrenceInput = {
  readonly pattern: RecurrencePatternType
  readonly interval?: number
  readonly days_of_week?: ReadonlyArray<string>
  readonly day_of_month?: number
  readonly month?: number
  readonly index?: string
  readonly first_day_of_week?: string
  readonly range_type?: RecurrenceRangeType
  readonly start_date?: string
  readonly end_date?: string
  readonly number_of_occurrences?: number
  readonly recurrence_time_zone?: string
}

// Graph's recurrenceRange dates are Date, not DateTimeOffset — passing a full
// datetime through is rejected. Callers naturally hand over the task's due date,
// which is a datetime, so it is trimmed here rather than at each call site.
export const toDateOnly = (value: string): string => value.slice(0, 10)

const buildPattern = (input: RecurrenceInput): GraphRecurrencePattern | string => {
  const interval = input.interval ?? 1
  if (!Number.isInteger(interval) || interval < 1) {
    return `interval must be a positive whole number, got ${String(input.interval)}.`
  }

  switch (input.pattern) {
    case "daily":
      return { type: "daily", interval }

    case "weekly": {
      if (!input.days_of_week || input.days_of_week.length === 0) {
        return "A weekly recurrence requires days_of_week, e.g. ['monday']."
      }
      return {
        type: "weekly",
        interval,
        daysOfWeek: input.days_of_week,
        // Graph lists firstDayOfWeek as required for weekly. It defaults to sunday,
        // which shifts which week an every-N-weeks pattern lands in, so it is sent
        // explicitly rather than left to the service default.
        firstDayOfWeek: input.first_day_of_week ?? "monday",
      }
    }

    case "absoluteMonthly": {
      if (input.day_of_month === undefined) {
        return "An absoluteMonthly recurrence requires day_of_month, e.g. 15."
      }
      return { type: "absoluteMonthly", interval, dayOfMonth: input.day_of_month }
    }

    case "relativeMonthly": {
      if (!input.days_of_week || input.days_of_week.length === 0) {
        return "A relativeMonthly recurrence requires days_of_week, e.g. ['saturday']."
      }
      return {
        type: "relativeMonthly",
        interval,
        daysOfWeek: input.days_of_week,
        index: input.index ?? "first",
      }
    }

    case "absoluteYearly": {
      if (input.day_of_month === undefined || input.month === undefined) {
        return "An absoluteYearly recurrence requires both day_of_month and month, e.g. 15 and 3."
      }
      return { type: "absoluteYearly", interval, dayOfMonth: input.day_of_month, month: input.month }
    }

    case "relativeYearly": {
      if (!input.days_of_week || input.days_of_week.length === 0 || input.month === undefined) {
        return "A relativeYearly recurrence requires both days_of_week and month."
      }
      return {
        type: "relativeYearly",
        interval,
        daysOfWeek: input.days_of_week,
        month: input.month,
        index: input.index ?? "first",
      }
    }
  }
}

const buildRange = (input: RecurrenceInput, fallbackStart?: string): GraphRecurrenceRange | string => {
  const start = input.start_date ?? fallbackStart
  if (!start) {
    return "A recurring task needs a start date. Pass start_date, or a due_date for it to start from."
  }

  const startDate = toDateOnly(start)
  const rangeType = input.range_type ?? "noEnd"

  switch (rangeType) {
    case "noEnd":
      return { type: "noEnd", startDate }

    case "endDate": {
      if (!input.end_date) return "A range_type of 'endDate' requires end_date."
      return { type: "endDate", startDate, endDate: toDateOnly(input.end_date) }
    }

    case "numbered": {
      const count = input.number_of_occurrences
      if (count === undefined || !Number.isInteger(count) || count < 1) {
        return "A range_type of 'numbered' requires a positive number_of_occurrences."
      }
      return { type: "numbered", startDate, numberOfOccurrences: count }
    }
  }
}

/**
 * Returns the Graph patternedRecurrence, or a message naming what was missing.
 * The caller decides how to surface the failure; this stays free of transport
 * and error types so it can be unit tested on its own.
 */
export const buildRecurrence = (input: RecurrenceInput, fallbackStart?: string): GraphPatternedRecurrence | string => {
  const pattern = buildPattern(input)
  if (typeof pattern === "string") return pattern

  const range = buildRange(input, fallbackStart)
  if (typeof range === "string") return range

  // The time zone is optional in Graph, and an explicit undefined is not the same
  // as an absent property — send it only when the caller named one.
  const withZone = Option(input.recurrence_time_zone).fold(
    () => range,
    (tz) => ({ ...range, recurrenceTimeZone: tz }),
  )

  return { pattern, range: withZone }
}

export const describeRecurrence = (recurrence: GraphPatternedRecurrence): string => {
  const { pattern, range } = recurrence
  const every = pattern.interval > 1 ? `every ${pattern.interval} ` : "every "

  const base = (() => {
    switch (pattern.type) {
      case "daily":
        return `${every}${pattern.interval > 1 ? "days" : "day"}`
      case "weekly":
        return `${every}${pattern.interval > 1 ? "weeks" : "week"} on ${(pattern.daysOfWeek ?? []).join(", ")}`
      case "absoluteMonthly":
        return `${every}${pattern.interval > 1 ? "months" : "month"} on day ${pattern.dayOfMonth}`
      case "relativeMonthly":
        return `${every}${pattern.interval > 1 ? "months" : "month"} on the ${pattern.index ?? "first"} ${(pattern.daysOfWeek ?? []).join(", ")}`
      case "absoluteYearly":
        return `${every}${pattern.interval > 1 ? "years" : "year"} on ${pattern.month}/${pattern.dayOfMonth}`
      case "relativeYearly":
        return `${every}${pattern.interval > 1 ? "years" : "year"} on the ${pattern.index ?? "first"} ${(pattern.daysOfWeek ?? []).join(", ")} of month ${pattern.month}`
      default:
        return pattern.type
    }
  })()

  const ends = (() => {
    switch (range.type) {
      case "endDate":
        return ` from ${range.startDate} until ${range.endDate}`
      case "numbered":
        return ` from ${range.startDate}, ${range.numberOfOccurrences} times`
      default:
        return ` from ${range.startDate}`
    }
  })()

  return `${base}${ends}`
}
