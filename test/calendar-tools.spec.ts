import { Some } from "functype"
import { Right } from "functype/either"
import { beforeEach, describe, expect, it, vi } from "vitest"

import type { GraphEvent } from "../src/types"

vi.mock("../src/client/graph-client", () => ({
  getGraphClient: vi.fn(),
}))

import { getGraphClient } from "../src/client/graph-client"
import { createEvent, findMeetingAvailability, listCalendarView, updateEvent } from "../src/tools/calendar-tools"

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
}

beforeEach(() => {
  vi.clearAllMocks()
  vi.mocked(getGraphClient).mockReturnValue(Some(mockClient as never))
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
      expect(mockClient.createEvent).toHaveBeenCalledWith({
        subject: "Meeting",
        start: { dateTime: "2026-04-07T10:00:00", timeZone: "UTC" },
        end: { dateTime: "2026-04-07T11:00:00", timeZone: "UTC" },
      })
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
      expect(mockClient.updateEvent).toHaveBeenCalledWith("evt-123", { subject: "Updated Meeting" })
    })

    it("should update event with Text body by default", async () => {
      mockClient.updateEvent.mockResolvedValue(Right(mockEvent))
      await updateEvent({ event_id: "evt-123", body: "New agenda" })
      expect(mockClient.updateEvent).toHaveBeenCalledWith("evt-123", {
        body: { contentType: "Text", content: "New agenda" },
      })
    })

    it("should update event with HTML body", async () => {
      mockClient.updateEvent.mockResolvedValue(Right(mockEvent))
      await updateEvent({ event_id: "evt-123", body: "<p>New agenda</p>", content_type: "HTML" })
      expect(mockClient.updateEvent).toHaveBeenCalledWith("evt-123", {
        body: { contentType: "HTML", content: "<p>New agenda</p>" },
      })
    })

    it("should update event time with custom timezone", async () => {
      mockClient.updateEvent.mockResolvedValue(Right(mockEvent))
      await updateEvent({
        event_id: "evt-123",
        start: "2026-04-07T14:00:00",
        end: "2026-04-07T15:00:00",
        time_zone: "America/Chicago",
      })
      expect(mockClient.updateEvent).toHaveBeenCalledWith("evt-123", {
        start: { dateTime: "2026-04-07T14:00:00", timeZone: "America/Chicago" },
        end: { dateTime: "2026-04-07T15:00:00", timeZone: "America/Chicago" },
      })
    })

    it("should update attendees from comma-separated string", async () => {
      mockClient.updateEvent.mockResolvedValue(Right(mockEvent))
      await updateEvent({
        event_id: "evt-123",
        attendees: "alice@example.com, bob@example.com",
      })
      expect(mockClient.updateEvent).toHaveBeenCalledWith("evt-123", {
        attendees: [
          { emailAddress: { address: "alice@example.com" }, type: "required" },
          { emailAddress: { address: "bob@example.com" }, type: "required" },
        ],
      })
    })

    it("should enable Teams meeting when online_meeting=true", async () => {
      mockClient.updateEvent.mockResolvedValue(Right(mockEvent))
      await updateEvent({ event_id: "evt-123", online_meeting: true })
      expect(mockClient.updateEvent).toHaveBeenCalledWith("evt-123", {
        isOnlineMeeting: true,
        onlineMeetingProvider: "teamsForBusiness",
      })
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
      expect(mockClient.listCalendarView).toHaveBeenCalledWith("2026-05-22T00:00:00Z", "2026-05-29T00:00:00Z", {
        $top: 50,
        $orderby: "start/dateTime",
      })
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
})
