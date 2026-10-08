import { describe, expect, it } from "vitest"

import type { GraphPatternedRecurrence } from "../src/types"
import { buildRecurrence, describeRecurrence } from "../src/utils/recurrence"

const START = "2026-10-01"

const built = (result: GraphPatternedRecurrence | string): GraphPatternedRecurrence => {
  if (typeof result === "string") throw new Error(`expected a recurrence, got: ${result}`)
  return result
}

describe("buildRecurrence", () => {
  it("should build a daily pattern with a default interval of 1", () => {
    const r = built(buildRecurrence({ pattern: "daily" }, START))
    expect(r.pattern).toEqual({ type: "daily", interval: 1 })
    expect(r.range).toEqual({ type: "noEnd", startDate: START })
  })

  // Graph errors on any property the pattern type does not support, so a daily
  // pattern must not carry the day_of_month a caller passed by mistake.
  it("should drop properties that do not belong to the pattern type", () => {
    const r = built(buildRecurrence({ pattern: "daily", day_of_month: 15, month: 3 }, START))
    expect(r.pattern).toEqual({ type: "daily", interval: 1 })
    expect(r.pattern).not.toHaveProperty("dayOfMonth")
    expect(r.pattern).not.toHaveProperty("month")
  })

  it("should build quarterly as absoluteMonthly with an interval of 3", () => {
    const r = built(buildRecurrence({ pattern: "absoluteMonthly", interval: 3, day_of_month: 15 }, START))
    expect(r.pattern).toEqual({ type: "absoluteMonthly", interval: 3, dayOfMonth: 15 })
  })

  it("should send firstDayOfWeek on a weekly pattern", () => {
    const r = built(buildRecurrence({ pattern: "weekly", days_of_week: ["monday"] }, START))
    expect(r.pattern).toEqual({
      type: "weekly",
      interval: 1,
      daysOfWeek: ["monday"],
      firstDayOfWeek: "monday",
    })
  })

  it("should default the range start to the task's due date", () => {
    const r = built(buildRecurrence({ pattern: "daily" }, "2026-10-01T09:00:00"))
    expect(r.range.startDate).toBe("2026-10-01")
  })

  it("should prefer an explicit start_date over the due date", () => {
    const r = built(buildRecurrence({ pattern: "daily", start_date: "2027-01-01" }, START))
    expect(r.range.startDate).toBe("2027-01-01")
  })

  it("should build a numbered range", () => {
    const r = built(buildRecurrence({ pattern: "daily", range_type: "numbered", number_of_occurrences: 5 }, START))
    expect(r.range).toEqual({ type: "numbered", startDate: START, numberOfOccurrences: 5 })
  })

  it("should build an endDate range and trim a datetime to a date", () => {
    const r = built(
      buildRecurrence({ pattern: "daily", range_type: "endDate", end_date: "2027-06-15T00:00:00" }, START),
    )
    expect(r.range).toEqual({ type: "endDate", startDate: START, endDate: "2027-06-15" })
  })

  it("should include a time zone only when one was given", () => {
    const withZone = built(buildRecurrence({ pattern: "daily", recurrence_time_zone: "UTC" }, START))
    expect(withZone.range.recurrenceTimeZone).toBe("UTC")
    const without = built(buildRecurrence({ pattern: "daily" }, START))
    expect(without.range).not.toHaveProperty("recurrenceTimeZone")
  })

  describe("missing required properties", () => {
    it("should reject a weekly pattern with no days", () => {
      expect(buildRecurrence({ pattern: "weekly" }, START)).toContain("days_of_week")
    })

    it("should reject an absoluteMonthly pattern with no day of month", () => {
      expect(buildRecurrence({ pattern: "absoluteMonthly" }, START)).toContain("day_of_month")
    })

    it("should reject an absoluteYearly pattern missing the month", () => {
      expect(buildRecurrence({ pattern: "absoluteYearly", day_of_month: 15 }, START)).toContain("month")
    })

    it("should reject a numbered range with no count", () => {
      expect(buildRecurrence({ pattern: "daily", range_type: "numbered" }, START)).toContain("number_of_occurrences")
    })

    it("should reject an endDate range with no end date", () => {
      expect(buildRecurrence({ pattern: "daily", range_type: "endDate" }, START)).toContain("end_date")
    })

    it("should reject a non-positive interval", () => {
      expect(buildRecurrence({ pattern: "daily", interval: 0 }, START)).toContain("positive")
    })

    it("should reject a recurrence with no start date available", () => {
      expect(buildRecurrence({ pattern: "daily" })).toContain("start date")
    })
  })
})

describe("describeRecurrence", () => {
  it("should describe a quarterly pattern in words", () => {
    const r = built(buildRecurrence({ pattern: "absoluteMonthly", interval: 3, day_of_month: 15 }, START))
    expect(describeRecurrence(r)).toBe("every 3 months on day 15 from 2026-10-01")
  })

  it("should describe an annual pattern", () => {
    const r = built(buildRecurrence({ pattern: "absoluteYearly", day_of_month: 1, month: 10 }, START))
    expect(describeRecurrence(r)).toBe("every year on 10/1 from 2026-10-01")
  })

  it("should describe a numbered range", () => {
    const r = built(buildRecurrence({ pattern: "daily", range_type: "numbered", number_of_occurrences: 5 }, START))
    expect(describeRecurrence(r)).toBe("every day from 2026-10-01, 5 times")
  })
})
