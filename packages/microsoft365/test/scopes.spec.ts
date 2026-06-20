import { describe, expect, it } from "vitest"

import { DEFAULT_INTERACTIVE_SCOPES, GRAPH_SCOPES } from "../src/auth/scopes"

describe("scopes", () => {
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
  })
})
