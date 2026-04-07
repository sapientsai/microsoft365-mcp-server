import { Some } from "functype"
import { Right } from "functype/either"
import { beforeEach, describe, expect, it, vi } from "vitest"

import type { GraphEvent } from "../src/types"

vi.mock("../src/client/graph-client", () => ({
  getGraphClient: vi.fn(),
}))

import { getGraphClient } from "../src/client/graph-client"
import { createEvent, updateEvent } from "../src/tools/calendar-tools"

const mockEvent: Partial<GraphEvent> = {
  id: "evt-123",
  subject: "Test Event",
  start: { dateTime: "2026-04-07T10:00:00", timeZone: "UTC" },
  end: { dateTime: "2026-04-07T11:00:00", timeZone: "UTC" },
}

const mockClient = {
  createEvent: vi.fn(),
  updateEvent: vi.fn(),
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

    it("should create draft event with isDraft flag", async () => {
      mockClient.createEvent.mockResolvedValue(Right(mockEvent))
      await createEvent({
        subject: "Draft Meeting",
        start: "2026-04-07T10:00:00",
        end: "2026-04-07T11:00:00",
        is_draft: true,
      })
      expect(mockClient.createEvent).toHaveBeenCalledWith(
        expect.objectContaining({
          isDraft: true,
        }),
      )
    })

    it("should not set isDraft when is_draft is false", async () => {
      mockClient.createEvent.mockResolvedValue(Right(mockEvent))
      await createEvent({
        subject: "Meeting",
        start: "2026-04-07T10:00:00",
        end: "2026-04-07T11:00:00",
        is_draft: false,
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

    it("should create draft event with attendees without sending invites", async () => {
      mockClient.createEvent.mockResolvedValue(Right(mockEvent))
      await createEvent({
        subject: "Draft Meeting",
        start: "2026-04-07T10:00:00",
        end: "2026-04-07T11:00:00",
        attendees: "alice@example.com",
        is_draft: true,
      })
      const callArg = mockClient.createEvent.mock.calls[0][0] as Record<string, unknown>
      expect(callArg).toHaveProperty("isDraft", true)
      expect(callArg).toHaveProperty("attendees")
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
  })
})
