import { Some } from "functype"
import { Left, Right } from "functype/either"
import { beforeEach, describe, expect, it, vi } from "vitest"

import type { GraphEvent } from "../src/types"

vi.mock("../src/client/graph-client", () => ({
  getGraphClient: vi.fn(),
}))

import { getGraphClient } from "../src/client/graph-client"
import {
  createEvent,
  deleteEvent,
  findMeetingAvailability,
  getEvent,
  listCalendarView,
  listEvents,
  updateEvent,
} from "../src/tools/calendar-tools"

const mockEvent: Partial<GraphEvent> = {
  id: "evt-123",
  subject: "Test Event",
  start: { dateTime: "2026-04-07T10:00:00", timeZone: "UTC" },
  end: { dateTime: "2026-04-07T11:00:00", timeZone: "UTC" },
}

const mockClient = {
  createEvent: vi.fn(),
  updateEvent: vi.fn(),
  listCalendarView: vi.fn(),
  findMeetingTimes: vi.fn(),
  getEvent: vi.fn(),
  getCalendar: vi.fn(),
  listEvents: vi.fn(),
  deleteEvent: vi.fn(),
}

beforeEach(() => {
  vi.clearAllMocks()
  vi.mocked(getGraphClient).mockReturnValue(Some(mockClient as never))
  mockClient.getCalendar.mockResolvedValue(Right({ id: "cal-1", name: "Team", canEdit: true }))
})

describe("calendar-tools", () => {
  describe("createEvent", () => {
    it("should create a basic event", async () => {
      mockClient.createEvent.mockResolvedValue(Right(mockEvent))
      const result = await createEvent({
        subject: "Meeting",
        start: "2026-04-07T10:00:00",
        end: "2026-04-07T11:00:00",
      })
      expect(result.isRight()).toBe(true)
      expect(result.value).toContain("Event created")
      expect(mockClient.createEvent).toHaveBeenCalledWith(
        {
          subject: "Meeting",
          start: { dateTime: "2026-04-07T10:00:00", timeZone: "UTC" },
          end: { dateTime: "2026-04-07T11:00:00", timeZone: "UTC" },
        },
        undefined,
        undefined,
      )
    })

    it("should create event with custom time zone", async () => {
      mockClient.createEvent.mockResolvedValue(Right(mockEvent))
      await createEvent({
        subject: "Meeting",
        start: "2026-04-07T10:00:00",
        end: "2026-04-07T11:00:00",
        time_zone: "America/New_York",
      })
      expect(mockClient.createEvent).toHaveBeenCalledWith(
        expect.objectContaining({
          start: { dateTime: "2026-04-07T10:00:00", timeZone: "America/New_York" },
          end: { dateTime: "2026-04-07T11:00:00", timeZone: "America/New_York" },
        }),
        undefined,
        "America/New_York",
      )
    })

    it("should create event with location", async () => {
      mockClient.createEvent.mockResolvedValue(Right(mockEvent))
      await createEvent({
        subject: "Meeting",
        start: "2026-04-07T10:00:00",
        end: "2026-04-07T11:00:00",
        location: "Conference Room A",
      })
      expect(mockClient.createEvent).toHaveBeenCalledWith(
        expect.objectContaining({
          location: { displayName: "Conference Room A" },
        }),
        undefined,
        undefined,
      )
    })

    it("should create event with Text body by default", async () => {
      mockClient.createEvent.mockResolvedValue(Right(mockEvent))
      await createEvent({
        subject: "Meeting",
        start: "2026-04-07T10:00:00",
        end: "2026-04-07T11:00:00",
        body: "Agenda items",
      })
      expect(mockClient.createEvent).toHaveBeenCalledWith(
        expect.objectContaining({
          body: { contentType: "Text", content: "Agenda items" },
        }),
        undefined,
        undefined,
      )
    })

    it("should create event with HTML body", async () => {
      mockClient.createEvent.mockResolvedValue(Right(mockEvent))
      await createEvent({
        subject: "Meeting",
        start: "2026-04-07T10:00:00",
        end: "2026-04-07T11:00:00",
        body: "<h1>Agenda</h1>",
        content_type: "HTML",
      })
      expect(mockClient.createEvent).toHaveBeenCalledWith(
        expect.objectContaining({
          body: { contentType: "HTML", content: "<h1>Agenda</h1>" },
        }),
        undefined,
        undefined,
      )
    })

    it("should never send isDraft to Graph (read-only per docs)", async () => {
      mockClient.createEvent.mockResolvedValue(Right(mockEvent))
      await createEvent({
        subject: "Draft Meeting",
        start: "2026-04-07T10:00:00",
        end: "2026-04-07T11:00:00",
        is_draft: true,
      })
      const callArg = mockClient.createEvent.mock.calls[0][0] as Record<string, unknown>
      expect(callArg).not.toHaveProperty("isDraft")
    })

    it("should create event with attendees", async () => {
      mockClient.createEvent.mockResolvedValue(Right(mockEvent))
      await createEvent({
        subject: "Meeting",
        start: "2026-04-07T10:00:00",
        end: "2026-04-07T11:00:00",
        attendees: "alice@example.com, bob@example.com",
      })
      expect(mockClient.createEvent).toHaveBeenCalledWith(
        expect.objectContaining({
          attendees: [
            { emailAddress: { address: "alice@example.com" }, type: "required" },
            { emailAddress: { address: "bob@example.com" }, type: "required" },
          ],
        }),
        undefined,
        undefined,
      )
    })

    it("should omit attendees from Graph payload when is_draft=true", async () => {
      mockClient.createEvent.mockResolvedValue(Right(mockEvent))
      const result = await createEvent({
        subject: "Draft Meeting",
        start: "2026-04-07T10:00:00",
        end: "2026-04-07T11:00:00",
        attendees: "alice@example.com",
        is_draft: true,
      })
      const callArg = mockClient.createEvent.mock.calls[0][0] as Record<string, unknown>
      expect(callArg).not.toHaveProperty("attendees")
      expect(callArg).not.toHaveProperty("isDraft")
      expect(result.value).toContain("Saved as draft")
    })
  })

  describe("updateEvent", () => {
    it("should update event subject", async () => {
      mockClient.updateEvent.mockResolvedValue(Right(mockEvent))
      const result = await updateEvent({ event_id: "evt-123", subject: "Updated Meeting" })
      expect(result.isRight()).toBe(true)
      expect(result.value).toContain("Event updated")
      expect(mockClient.updateEvent).toHaveBeenCalledWith(
        "evt-123",
        { subject: "Updated Meeting" },
        undefined,
        undefined,
      )
    })

    it("should update event with Text body by default", async () => {
      mockClient.updateEvent.mockResolvedValue(Right(mockEvent))
      await updateEvent({ event_id: "evt-123", body: "New agenda" })
      expect(mockClient.updateEvent).toHaveBeenCalledWith(
        "evt-123",
        {
          body: { contentType: "Text", content: "New agenda" },
        },
        undefined,
        undefined,
      )
    })

    it("should update event with HTML body", async () => {
      mockClient.updateEvent.mockResolvedValue(Right(mockEvent))
      await updateEvent({ event_id: "evt-123", body: "<p>New agenda</p>", content_type: "HTML" })
      expect(mockClient.updateEvent).toHaveBeenCalledWith(
        "evt-123",
        {
          body: { contentType: "HTML", content: "<p>New agenda</p>" },
        },
        undefined,
        undefined,
      )
    })

    it("should update event time with custom timezone", async () => {
      mockClient.updateEvent.mockResolvedValue(Right(mockEvent))
      await updateEvent({
        event_id: "evt-123",
        start: "2026-04-07T14:00:00",
        end: "2026-04-07T15:00:00",
        time_zone: "America/Chicago",
      })
      expect(mockClient.updateEvent).toHaveBeenCalledWith(
        "evt-123",
        {
          start: { dateTime: "2026-04-07T14:00:00", timeZone: "America/Chicago" },
          end: { dateTime: "2026-04-07T15:00:00", timeZone: "America/Chicago" },
        },
        undefined,
        "America/Chicago",
      )
    })

    it("should update attendees from comma-separated string", async () => {
      mockClient.updateEvent.mockResolvedValue(Right(mockEvent))
      await updateEvent({
        event_id: "evt-123",
        attendees: "alice@example.com, bob@example.com",
      })
      expect(mockClient.updateEvent).toHaveBeenCalledWith(
        "evt-123",
        {
          attendees: [
            { emailAddress: { address: "alice@example.com" }, type: "required" },
            { emailAddress: { address: "bob@example.com" }, type: "required" },
          ],
        },
        undefined,
        undefined,
      )
    })

    it("should enable Teams meeting when online_meeting=true", async () => {
      mockClient.updateEvent.mockResolvedValue(Right(mockEvent))
      await updateEvent({ event_id: "evt-123", online_meeting: true })
      expect(mockClient.updateEvent).toHaveBeenCalledWith(
        "evt-123",
        {
          isOnlineMeeting: true,
          onlineMeetingProvider: "teamsForBusiness",
        },
        undefined,
        undefined,
      )
    })

    it("should not send online meeting flags when online_meeting=false", async () => {
      mockClient.updateEvent.mockResolvedValue(Right(mockEvent))
      await updateEvent({ event_id: "evt-123", subject: "Renamed", online_meeting: false })
      const updates = mockClient.updateEvent.mock.calls[0][1] as Record<string, unknown>
      expect(updates).not.toHaveProperty("isOnlineMeeting")
      expect(updates).not.toHaveProperty("onlineMeetingProvider")
    })
  })

  describe("listCalendarView", () => {
    it("should call listCalendarView with start/end and ordering", async () => {
      mockClient.listCalendarView.mockResolvedValue(Right({ value: [mockEvent] }))
      const result = await listCalendarView({
        start_date_time: "2026-05-22T00:00:00Z",
        end_date_time: "2026-05-29T00:00:00Z",
      })
      expect(result.isRight()).toBe(true)
      expect(mockClient.listCalendarView).toHaveBeenCalledWith(
        "2026-05-22T00:00:00Z",
        "2026-05-29T00:00:00Z",
        {
          $top: 50,
          $orderby: "start/dateTime",
        },
        undefined,
        undefined,
      )
    })

    it("should honor custom top", async () => {
      mockClient.listCalendarView.mockResolvedValue(Right({ value: [] }))
      await listCalendarView({
        start_date_time: "2026-05-22T00:00:00Z",
        end_date_time: "2026-05-29T00:00:00Z",
        top: 100,
      })
      expect(mockClient.listCalendarView).toHaveBeenCalledWith(
        "2026-05-22T00:00:00Z",
        "2026-05-29T00:00:00Z",
        expect.objectContaining({ $top: 100 }),
        undefined,
        undefined,
      )
    })

    it("should reject missing start_date_time", async () => {
      const result = await listCalendarView({ start_date_time: "", end_date_time: "2026-05-29T00:00:00Z" })
      expect(result.isLeft()).toBe(true)
      expect(mockClient.listCalendarView).not.toHaveBeenCalled()
    })

    it("should reject missing end_date_time", async () => {
      const result = await listCalendarView({ start_date_time: "2026-05-22T00:00:00Z", end_date_time: "" })
      expect(result.isLeft()).toBe(true)
      expect(mockClient.listCalendarView).not.toHaveBeenCalled()
    })
  })

  describe("findMeetingAvailability", () => {
    const mockResult = {
      emptySuggestionsReason: "",
      meetingTimeSuggestions: [
        {
          confidence: 100,
          organizerAvailability: "free",
          attendeeAvailability: [{ availability: "free", attendee: { emailAddress: { address: "bob@example.com" } } }],
          meetingTimeSlot: {
            start: { dateTime: "2026-06-04T15:00:00.0000000", timeZone: "UTC" },
            end: { dateTime: "2026-06-04T15:30:00.0000000", timeZone: "UTC" },
          },
        },
      ],
    }

    it("should return ranked suggestions and send the correct Graph body", async () => {
      mockClient.findMeetingTimes.mockResolvedValue(Right(mockResult))
      const result = await findMeetingAvailability({
        participants: ["bob@example.com"],
        after_date_time: "2026-06-04T00:00:00Z",
        before_date_time: "2026-06-06T00:00:00Z",
      })
      expect(result.isRight()).toBe(true)
      expect(result.value).toContain("Meeting Time Suggestions")
      expect(result.value).toContain("100% confidence")
      expect(mockClient.findMeetingTimes).toHaveBeenCalledWith(
        expect.objectContaining({
          attendees: [{ type: "required", emailAddress: { address: "bob@example.com" } }],
          meetingDuration: "PT30M",
          maxCandidates: 3,
          isOrganizerOptional: false,
          minimumAttendeePercentage: 100,
          returnSuggestionReasons: true,
        }),
      )
    })

    it("should encode duration_minutes as an ISO 8601 duration", async () => {
      mockClient.findMeetingTimes.mockResolvedValue(Right(mockResult))
      await findMeetingAvailability({
        participants: ["bob@example.com"],
        after_date_time: "2026-06-04T00:00:00Z",
        before_date_time: "2026-06-06T00:00:00Z",
        duration_minutes: 45,
      })
      expect(mockClient.findMeetingTimes).toHaveBeenCalledWith(expect.objectContaining({ meetingDuration: "PT45M" }))
    })

    it("should surface emptySuggestionsReason when no slot is found", async () => {
      mockClient.findMeetingTimes.mockResolvedValue(
        Right({ emptySuggestionsReason: "AttendeesUnavailable", meetingTimeSuggestions: [] }),
      )
      const result = await findMeetingAvailability({
        participants: ["bob@example.com"],
        after_date_time: "2026-06-04T00:00:00Z",
        before_date_time: "2026-06-06T00:00:00Z",
      })
      expect(result.isRight()).toBe(true)
      expect(result.value).toContain("No common availability found")
      expect(result.value).toContain("AttendeesUnavailable")
    })

    it("should reject an empty participants list", async () => {
      const result = await findMeetingAvailability({
        participants: [],
        after_date_time: "2026-06-04T00:00:00Z",
        before_date_time: "2026-06-06T00:00:00Z",
      })
      expect(result.isLeft()).toBe(true)
      expect(mockClient.findMeetingTimes).not.toHaveBeenCalled()
    })

    it("should reject when before_date_time is not after after_date_time", async () => {
      const result = await findMeetingAvailability({
        participants: ["bob@example.com"],
        after_date_time: "2026-06-06T00:00:00Z",
        before_date_time: "2026-06-04T00:00:00Z",
      })
      expect(result.isLeft()).toBe(true)
      expect(mockClient.findMeetingTimes).not.toHaveBeenCalled()
    })
  })

  // A Teams invite's HTML is layout markup around a join link; the report that prompted this saw it raw.
  describe("getEvent body format", () => {
    it("asks Graph for a text body by default", async () => {
      mockClient.getEvent.mockResolvedValue(
        Right({ ...mockEvent, body: { contentType: "text", content: "Join here" } }),
      )

      const result = await getEvent({ event_id: "e1" })

      expect(mockClient.getEvent).toHaveBeenCalledWith("e1", "text", undefined, undefined)
      expect(result.value).toContain("Join here")
    })

    it("passes html through when asked", async () => {
      mockClient.getEvent.mockResolvedValue(Right(mockEvent))

      await getEvent({ event_id: "e1", body_format: "html" })

      expect(mockClient.getEvent).toHaveBeenCalledWith("e1", "html", undefined, undefined)
    })
  })

  // Graph stores an all-day event midnight to midnight with an exclusive end; callers name the
  // last day. Compass's staff days and carnivals are the motivating case.
  describe("all-day events", () => {
    const perth = "Australia/Perth"

    it("create_event sends isAllDay with midnight times and the day after the last day as end", async () => {
      mockClient.createEvent.mockResolvedValue(Right(mockEvent))

      const result = await createEvent({
        subject: "Swimming carnival",
        start: "2026-11-06",
        end: "2026-11-07",
        all_day: true,
        time_zone: perth,
      })

      expect(result.isRight()).toBe(true)
      expect(mockClient.createEvent).toHaveBeenCalledWith(
        {
          subject: "Swimming carnival",
          start: { dateTime: "2026-11-06T00:00:00", timeZone: perth },
          end: { dateTime: "2026-11-08T00:00:00", timeZone: perth },
          isAllDay: true,
          showAs: "free",
        },
        undefined,
        "Australia/Perth",
      )
    })

    it("defaults end to start, giving a one-day event", async () => {
      mockClient.createEvent.mockResolvedValue(Right(mockEvent))

      await createEvent({ subject: "Last day of term", start: "2026-12-18", all_day: true, time_zone: perth })

      const sent = mockClient.createEvent.mock.calls[0][0] as Record<string, { dateTime: string }>
      expect(sent.start!.dateTime).toBe("2026-12-18T00:00:00")
      expect(sent.end!.dateTime).toBe("2026-12-19T00:00:00")
    })

    it("rolls the end over month and year boundaries", async () => {
      mockClient.createEvent.mockResolvedValue(Right(mockEvent))

      await createEvent({ subject: "NYE", start: "2026-12-31", all_day: true })
      await createEvent({ subject: "Leap", start: "2028-02-28", end: "2028-02-29", all_day: true })

      expect((mockClient.createEvent.mock.calls[0][0] as { end: { dateTime: string } }).end.dateTime).toBe(
        "2027-01-01T00:00:00",
      )
      expect((mockClient.createEvent.mock.calls[1][0] as { end: { dateTime: string } }).end.dateTime).toBe(
        "2028-03-01T00:00:00",
      )
    })

    it("accepts a midnight datetime as a date", async () => {
      mockClient.createEvent.mockResolvedValue(Right(mockEvent))

      await createEvent({ subject: "PD day", start: "2026-11-13T00:00:00", all_day: true })

      expect((mockClient.createEvent.mock.calls[0][0] as { end: { dateTime: string } }).end.dateTime).toBe(
        "2026-11-14T00:00:00",
      )
    })

    it.each([
      ["a time of day", { start: "2026-11-06T09:00:00" }],
      ["a malformed date", { start: "6/11/2026" }],
      ["an impossible date", { start: "2026-02-30" }],
      ["an end before the start", { start: "2026-11-06", end: "2026-11-05" }],
    ])("refuses %s before calling Graph", async (_label, dates) => {
      const result = await createEvent({ subject: "x", all_day: true, ...dates })

      expect(result.isLeft()).toBe(true)
      expect(mockClient.createEvent).not.toHaveBeenCalled()
    })

    it("still requires end for a timed event", async () => {
      const result = await createEvent({ subject: "x", start: "2026-11-06T09:00:00" })

      expect(result.isLeft()).toBe(true)
      expect(mockClient.createEvent).not.toHaveBeenCalled()
    })

    it("does not send isAllDay for a timed event", async () => {
      mockClient.createEvent.mockResolvedValue(Right(mockEvent))

      await createEvent({ subject: "x", start: "2026-11-06T09:00:00", end: "2026-11-06T10:00:00" })

      expect(mockClient.createEvent.mock.calls[0][0]).not.toHaveProperty("isAllDay")
    })

    it("update_event all_day: true rewrites the span and sets isAllDay", async () => {
      mockClient.getEvent.mockResolvedValue(Right({ ...mockEvent, isAllDay: false }))
      mockClient.updateEvent.mockResolvedValue(Right(mockEvent))

      await updateEvent({ event_id: "evt-123", all_day: true, start: "2026-11-06", time_zone: perth })

      expect(mockClient.updateEvent).toHaveBeenCalledWith(
        "evt-123",
        {
          isAllDay: true,
          start: { dateTime: "2026-11-06T00:00:00", timeZone: perth },
          end: { dateTime: "2026-11-07T00:00:00", timeZone: perth },
          showAs: "free",
        },
        undefined,
        "Australia/Perth",
      )
    })

    it("update_event all_day: true without start is refused", async () => {
      const result = await updateEvent({ event_id: "evt-123", all_day: true })

      expect(result.isLeft()).toBe(true)
      expect(mockClient.updateEvent).not.toHaveBeenCalled()
    })

    it("update_event all_day: false needs both times, then clears isAllDay", async () => {
      mockClient.getEvent.mockResolvedValue(Right({ ...mockEvent, isAllDay: true }))
      mockClient.updateEvent.mockResolvedValue(Right(mockEvent))

      const missing = await updateEvent({ event_id: "evt-123", all_day: false, start: "2026-11-06T09:00:00" })
      expect(missing.isLeft()).toBe(true)

      await updateEvent({
        event_id: "evt-123",
        all_day: false,
        start: "2026-11-06T09:00:00",
        end: "2026-11-06T10:00:00",
      })
      expect(mockClient.updateEvent).toHaveBeenCalledOnce()
      expect(mockClient.updateEvent.mock.calls[0][1]).toMatchObject({ isAllDay: false, showAs: "busy" })
    })
  })
  // Outlook marks an all-day event free; Graph would mark it busy and block the day for everyone.
  describe("show_as", () => {
    it("create_event leaves a timed event to Graph's default (busy)", async () => {
      mockClient.createEvent.mockResolvedValue(Right(mockEvent))

      await createEvent({ subject: "x", start: "2026-11-06T09:00:00", end: "2026-11-06T10:00:00" })

      expect(mockClient.createEvent.mock.calls[0][0]).not.toHaveProperty("showAs")
    })

    it("an explicit show_as wins over the all-day default", async () => {
      mockClient.createEvent.mockResolvedValue(Right(mockEvent))

      await createEvent({ subject: "Camp", start: "2026-11-06", all_day: true, show_as: "oof" })

      expect(mockClient.createEvent.mock.calls[0][0]).toMatchObject({ isAllDay: true, showAs: "oof" })
    })

    it("create_event passes show_as through on a timed event", async () => {
      mockClient.createEvent.mockResolvedValue(Right(mockEvent))

      await createEvent({
        subject: "x",
        start: "2026-11-06T09:00:00",
        end: "2026-11-06T10:00:00",
        show_as: "tentative",
      })

      expect(mockClient.createEvent.mock.calls[0][0]).toMatchObject({ showAs: "tentative" })
    })

    it("update_event without all_day or show_as leaves showAs alone", async () => {
      mockClient.updateEvent.mockResolvedValue(Right(mockEvent))

      await updateEvent({ event_id: "evt-123", subject: "Renamed" })

      expect(mockClient.updateEvent.mock.calls[0][1]).not.toHaveProperty("showAs")
    })

    it("update_event moving an all-day event keeps the status someone chose", async () => {
      mockClient.getEvent.mockResolvedValue(Right({ ...mockEvent, isAllDay: true, showAs: "oof" }))
      mockClient.updateEvent.mockResolvedValue(Right(mockEvent))

      await updateEvent({ event_id: "evt-123", all_day: true, start: "2026-11-07", calendar_id: "cal-1" })

      expect(mockClient.getEvent).toHaveBeenCalledWith("evt-123", "text", "cal-1")
      expect(mockClient.updateEvent.mock.calls[0][1]).not.toHaveProperty("showAs")
    })

    it("update_event all_day: false on a timed event keeps its status", async () => {
      mockClient.getEvent.mockResolvedValue(Right({ ...mockEvent, isAllDay: false, showAs: "tentative" }))
      mockClient.updateEvent.mockResolvedValue(Right(mockEvent))

      await updateEvent({
        event_id: "evt-123",
        all_day: false,
        start: "2026-11-06T09:00:00",
        end: "2026-11-06T10:00:00",
      })

      expect(mockClient.updateEvent.mock.calls[0][1]).not.toHaveProperty("showAs")
    })

    it("update_event reads nothing when show_as is given", async () => {
      mockClient.updateEvent.mockResolvedValue(Right(mockEvent))

      await updateEvent({ event_id: "evt-123", all_day: true, start: "2026-11-07", show_as: "oof" })

      expect(mockClient.getEvent).not.toHaveBeenCalled()
      expect(mockClient.updateEvent.mock.calls[0][1]).toMatchObject({ showAs: "oof" })
    })

    it("update_event changes nothing when the read fails", async () => {
      mockClient.getEvent.mockResolvedValue(Left(new Error("Graph is down")))

      const result = await updateEvent({ event_id: "evt-123", all_day: true, start: "2026-11-07" })

      expect(result.isLeft()).toBe(true)
      expect(String((result.value as Error).message)).toContain("Graph is down")
      expect(mockClient.updateEvent).not.toHaveBeenCalled()
    })

    it("update_event sets show_as on its own", async () => {
      mockClient.updateEvent.mockResolvedValue(Right(mockEvent))

      await updateEvent({ event_id: "evt-123", show_as: "free" })

      expect(mockClient.updateEvent).toHaveBeenCalledWith("evt-123", { showAs: "free" }, undefined, undefined)
    })
  })
  // Graph's default reminder for an all-day event fires at 23:45 the night before.
  describe("reminders", () => {
    it("create_event leaves the reminder to Graph when neither option is given", async () => {
      mockClient.createEvent.mockResolvedValue(Right(mockEvent))

      await createEvent({ subject: "x", start: "2026-11-06", all_day: true })

      expect(mockClient.createEvent.mock.calls[0][0]).not.toHaveProperty("isReminderOn")
      expect(mockClient.createEvent.mock.calls[0][0]).not.toHaveProperty("reminderMinutesBeforeStart")
    })

    it("create_event with reminder_minutes turns the reminder on with that lead time", async () => {
      mockClient.createEvent.mockResolvedValue(Right(mockEvent))

      await createEvent({ subject: "Swimming", start: "2026-11-23", all_day: true, reminder_minutes: 1080 })

      expect(mockClient.createEvent.mock.calls[0][0]).toMatchObject({
        isReminderOn: true,
        reminderMinutesBeforeStart: 1080,
      })
    })

    it("create_event with reminder: false turns the reminder off", async () => {
      mockClient.createEvent.mockResolvedValue(Right(mockEvent))

      await createEvent({ subject: "x", start: "2026-11-06", all_day: true, reminder: false })

      expect(mockClient.createEvent.mock.calls[0][0]).toMatchObject({ isReminderOn: false })
      expect(mockClient.createEvent.mock.calls[0][0]).not.toHaveProperty("reminderMinutesBeforeStart")
    })

    it("refuses reminder: false together with reminder_minutes", async () => {
      const result = await createEvent({
        subject: "x",
        start: "2026-11-06",
        all_day: true,
        reminder: false,
        reminder_minutes: 30,
      })

      expect(result.isLeft()).toBe(true)
      expect(mockClient.createEvent).not.toHaveBeenCalled()
    })

    it("update_event changes only the reminder", async () => {
      mockClient.updateEvent.mockResolvedValue(Right(mockEvent))

      await updateEvent({ event_id: "evt-123", reminder: false })

      expect(mockClient.updateEvent).toHaveBeenCalledWith("evt-123", { isReminderOn: false }, undefined, undefined)
    })

    it("update_event sets a new lead time", async () => {
      mockClient.updateEvent.mockResolvedValue(Right(mockEvent))

      await updateEvent({ event_id: "evt-123", reminder_minutes: 60 })

      expect(mockClient.updateEvent).toHaveBeenCalledWith(
        "evt-123",
        { isReminderOn: true, reminderMinutesBeforeStart: 60 },
        undefined,
        undefined,
      )
    })

    it("update_event without reminder options leaves the reminder alone", async () => {
      mockClient.updateEvent.mockResolvedValue(Right(mockEvent))

      await updateEvent({ event_id: "evt-123", subject: "Renamed" })

      expect(mockClient.updateEvent.mock.calls[0][1]).not.toHaveProperty("isReminderOn")
    })

    it("event details show the reminder", async () => {
      mockClient.updateEvent.mockResolvedValue(
        Right({ ...mockEvent, isReminderOn: true, reminderMinutesBeforeStart: 1080 }),
      )

      const result = await updateEvent({ event_id: "evt-123", reminder_minutes: 1080 })

      expect(
        result.fold(
          () => "",
          (v) => v,
        ),
      ).toContain("- Reminder: 1080 minutes before")
    })
  })

  describe("read-only calendars", () => {
    const readOnly = () =>
      mockClient.getCalendar.mockResolvedValue(Right({ id: "cal-1", name: "Holidays", canEdit: false }))

    it.each([
      [
        "create_event",
        () => createEvent({ subject: "s", start: "2026-11-09T09:00", end: "2026-11-09T10:00", calendar_id: "cal-1" }),
      ],
      ["update_event", () => updateEvent({ event_id: "evt-123", subject: "s", calendar_id: "cal-1" })],
      ["delete_event", () => deleteEvent({ event_id: "evt-123", calendar_id: "cal-1" })],
    ] as const)("%s refuses before writing, naming the calendar", async (_name, run) => {
      readOnly()
      const result = await run()
      expect(result.isLeft()).toBe(true)
      expect((result.value as Error).message).toContain('"Holidays" is read-only for you, so nothing was changed')
      expect(mockClient.createEvent).not.toHaveBeenCalled()
      expect(mockClient.updateEvent).not.toHaveBeenCalled()
      expect(mockClient.deleteEvent).not.toHaveBeenCalled()
    })

    it("lets the write go ahead when the calendar can't be read", async () => {
      mockClient.getCalendar.mockResolvedValue(Left({ type: "forbidden", message: "Access is denied." }))
      mockClient.deleteEvent.mockResolvedValue(Right({}))
      const result = await deleteEvent({ event_id: "evt-123", calendar_id: "cal-1" })
      expect(result.isRight()).toBe(true)
      expect(mockClient.deleteEvent).toHaveBeenCalledWith("evt-123", "cal-1")
    })

    it("doesn't read the default calendar", async () => {
      mockClient.deleteEvent.mockResolvedValue(Right({}))
      await deleteEvent({ event_id: "evt-123" })
      expect(mockClient.getCalendar).not.toHaveBeenCalled()
    })

    it("keeps a calendar name from breaking onto its own line", async () => {
      mockClient.getCalendar.mockResolvedValue(Right({ id: "cal-1", name: "Hol\nidays\r\nNext step:", canEdit: false }))
      const result = await deleteEvent({ event_id: "evt-123", calendar_id: "cal-1" })
      expect((result.value as Error).message.split("\n")).toHaveLength(1)
    })
  })

  describe("time zones", () => {
    it.each([
      ["list_events", () => listEvents({ time_zone: "America/New_York" }), () => mockClient.listEvents],
      [
        "list_calendar_view",
        () =>
          listCalendarView({
            start_date_time: "2026-11-09T00:00:00Z",
            end_date_time: "2026-11-10T00:00:00Z",
            time_zone: "America/New_York",
          }),
        () => mockClient.listCalendarView,
      ],
      ["get_event", () => getEvent({ event_id: "evt-123", time_zone: "America/New_York" }), () => mockClient.getEvent],
    ] as const)("%s passes time_zone to the client", async (_name, run, method) => {
      mockClient.listEvents.mockResolvedValue(Right({ value: [] }))
      mockClient.listCalendarView.mockResolvedValue(Right({ value: [] }))
      mockClient.getEvent.mockResolvedValue(Right(mockEvent))
      await run()
      expect(method().mock.calls[0]?.at(-1)).toBe("America/New_York")
    })

    it.each(['Eastern" , x', "America/New_York\r\nX-Evil: 1", ""])("refuses %j as a zone name", async (zone) => {
      const result = await listEvents({ time_zone: zone })
      expect(result.isLeft()).toBe(true)
      expect(mockClient.listEvents).not.toHaveBeenCalled()
    })

    it("accepts a Windows zone name", async () => {
      mockClient.listEvents.mockResolvedValue(Right({ value: [] }))
      const result = await listEvents({ time_zone: "Eastern Standard Time" })
      expect(result.isRight()).toBe(true)
    })

    it("prints every time with its zone", async () => {
      mockClient.listEvents.mockResolvedValue(
        Right({
          value: [
            {
              id: "e1",
              subject: "Standup",
              start: { dateTime: "2026-11-09T09:30:00.0000000", timeZone: "America/New_York" },
              end: { dateTime: "2026-11-09T10:00:00.0000000", timeZone: "America/New_York" },
              originalStartTimeZone: "GMT Standard Time",
            },
          ],
        }),
      )
      const result = await listEvents({ time_zone: "America/New_York" })
      expect(result.value).toContain(
        "**Standup** (2026-11-09 09:30–10:00 America/New_York; set in GMT Standard Time) (ID: e1)",
      )
    })
  })

  describe("write replies", () => {
    it("create_event leaves out the body Graph echoes back as HTML", async () => {
      mockClient.createEvent.mockResolvedValue(
        Right({ ...mockEvent, body: { contentType: "html", content: "<html><body>Agenda</body></html>" } }),
      )
      const result = await createEvent({ subject: "s", start: "2026-11-09T09:00", end: "2026-11-09T10:00" })
      expect(result.value).not.toContain("<html>")
      expect(result.value).not.toContain("## Body")
    })

    it("update_event leaves it out too", async () => {
      mockClient.updateEvent.mockResolvedValue(
        Right({ ...mockEvent, body: { contentType: "html", content: "<html><body>Agenda</body></html>" } }),
      )
      const result = await updateEvent({ event_id: "evt-123", subject: "s" })
      expect(result.value).not.toContain("<html>")
    })

    it("get_event still shows the body", async () => {
      mockClient.getEvent.mockResolvedValue(Right({ ...mockEvent, body: { contentType: "text", content: "Agenda" } }))
      const result = await getEvent({ event_id: "evt-123" })
      expect(result.value).toContain("## Body\nAgenda")
    })
  })
})
