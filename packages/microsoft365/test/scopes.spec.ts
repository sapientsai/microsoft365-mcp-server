import { afterEach, beforeEach, describe, expect, it } from "vitest"

import {
  DEFAULT_INTERACTIVE_SCOPES,
  GRAPH_SCOPES,
  resolveExtraScopes,
  resolveInteractiveScopes,
} from "../src/auth/scopes"

describe("scopes", () => {
  // resolveExtraScopes defaults to reading process.env, so a developer with MS365_EXTRA_SCOPES
  // exported would otherwise fail the "unset" cases against their own shell.
  const saved = process.env.MS365_EXTRA_SCOPES
  beforeEach(() => delete process.env.MS365_EXTRA_SCOPES)
  afterEach(() => {
    if (saved === undefined) delete process.env.MS365_EXTRA_SCOPES
    else process.env.MS365_EXTRA_SCOPES = saved
  })

  describe("GRAPH_SCOPES", () => {
    it("should define Mail.ReadWrite scope", () => {
      expect(GRAPH_SCOPES.MAIL_READWRITE).toBe("Mail.ReadWrite")
    })

    it("should define all mail scopes", () => {
      expect(GRAPH_SCOPES.MAIL_READ).toBe("Mail.Read")
      expect(GRAPH_SCOPES.MAIL_READWRITE).toBe("Mail.ReadWrite")
      expect(GRAPH_SCOPES.MAIL_SEND).toBe("Mail.Send")
    })
  })

  describe("DEFAULT_INTERACTIVE_SCOPES", () => {
    it("should include Mail.ReadWrite for draft support", () => {
      expect(DEFAULT_INTERACTIVE_SCOPES).toContain("Mail.ReadWrite")
    })

    it("should include Mail.Read and Mail.Send", () => {
      expect(DEFAULT_INTERACTIVE_SCOPES).toContain("Mail.Read")
      expect(DEFAULT_INTERACTIVE_SCOPES).toContain("Mail.Send")
    })

    it("should include calendar write scope", () => {
      expect(DEFAULT_INTERACTIVE_SCOPES).toContain("Calendars.ReadWrite")
    })

    it("should include files write scope", () => {
      expect(DEFAULT_INTERACTIVE_SCOPES).toContain("Files.ReadWrite")
    })

    it("should include tasks write scope for Planner and To Do", () => {
      expect(DEFAULT_INTERACTIVE_SCOPES).toContain("Tasks.ReadWrite")
    })

    it("should include Teams and Chat scopes", () => {
      expect(DEFAULT_INTERACTIVE_SCOPES).toContain("Chat.ReadWrite")
      expect(DEFAULT_INTERACTIVE_SCOPES).toContain("ChatMessage.Read")
      expect(DEFAULT_INTERACTIVE_SCOPES).toContain("ChatMessage.Send")
      expect(DEFAULT_INTERACTIVE_SCOPES).toContain("ChannelMessage.Send")
    })

    it("should include SharePoint scopes", () => {
      expect(DEFAULT_INTERACTIVE_SCOPES).toContain("Sites.Read.All")
      expect(DEFAULT_INTERACTIVE_SCOPES).toContain("Sites.ReadWrite.All")
    })

    it("should have no duplicate scopes", () => {
      expect(new Set(DEFAULT_INTERACTIVE_SCOPES).size).toBe(DEFAULT_INTERACTIVE_SCOPES.length)
    })

    // This is the guard, not a style check. OnlineMeetingTranscript.Read.All and OnlineMeetings.Read
    // are admin-consent permissions: a non-admin user cannot consent past them, so putting either in
    // the defaults would fail sign-in for every tenant that has not granted it — including tenants
    // that never wanted transcripts. They stay opt-in via MS365_EXTRA_SCOPES.
    it("must not request admin-consent meeting scopes by default", () => {
      expect(DEFAULT_INTERACTIVE_SCOPES).not.toContain(GRAPH_SCOPES.ONLINE_MEETING_TRANSCRIPT_READ_ALL)
      expect(DEFAULT_INTERACTIVE_SCOPES).not.toContain(GRAPH_SCOPES.ONLINE_MEETINGS_READ)
    })
  })

  describe("meeting scopes", () => {
    it("names the transcript permissions so callers do not hand-type them", () => {
      expect(GRAPH_SCOPES.ONLINE_MEETING_TRANSCRIPT_READ_ALL).toBe("OnlineMeetingTranscript.Read.All")
      expect(GRAPH_SCOPES.ONLINE_MEETINGS_READ).toBe("OnlineMeetings.Read")
    })
  })

  describe("resolveExtraScopes", () => {
    it("returns nothing when MS365_EXTRA_SCOPES is unset or empty", () => {
      expect(resolveExtraScopes(undefined)).toEqual([])
      expect(resolveExtraScopes("")).toEqual([])
      expect(resolveExtraScopes("  ,  ,")).toEqual([])
    })

    it("splits on commas and trims", () => {
      expect(resolveExtraScopes(" OnlineMeetings.Read , OnlineMeetingTranscript.Read.All ")).toEqual([
        "OnlineMeetings.Read",
        "OnlineMeetingTranscript.Read.All",
      ])
    })

    it("drops scopes already in the defaults and repeats within the value", () => {
      expect(resolveExtraScopes("Mail.Read,OnlineMeetings.Read,OnlineMeetings.Read")).toEqual(["OnlineMeetings.Read"])
    })
  })

  describe("resolveInteractiveScopes", () => {
    it("is the defaults when nothing extra is configured", () => {
      expect(resolveInteractiveScopes("")).toEqual([...DEFAULT_INTERACTIVE_SCOPES])
    })

    it("appends the extras without disturbing the defaults", () => {
      const scopes = resolveInteractiveScopes("OnlineMeetingTranscript.Read.All")

      expect(scopes.slice(0, DEFAULT_INTERACTIVE_SCOPES.length)).toEqual([...DEFAULT_INTERACTIVE_SCOPES])
      expect(scopes).toContain("OnlineMeetingTranscript.Read.All")
      expect(new Set(scopes).size).toBe(scopes.length)
    })
  })
})
