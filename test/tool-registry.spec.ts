import { describe, expect, it } from "vitest"

import { filterTools, PRESETS, TOOL_METADATA } from "../src/tools/tool-registry"

describe("tool-registry", () => {
  describe("PRESETS", () => {
    it("should define expected preset names", () => {
      expect(Object.keys(PRESETS)).toEqual(["personal", "collaboration", "productivity", "all"])
    })

    it("personal preset should include mail, calendar, contacts, todo, files, onenote", () => {
      expect(PRESETS.personal).toEqual(["mail", "calendar", "contacts", "todo", "files", "onenote"])
    })

    it("collaboration preset should include chats, teams, planner, groups", () => {
      expect(PRESETS.collaboration).toEqual(["chats", "teams", "planner", "groups"])
    })
  })

  describe("TOOL_METADATA", () => {
    it("should have unique tool names", () => {
      const names = TOOL_METADATA.map((m) => m.name)
      expect(new Set(names).size).toBe(names.length)
    })

    it("should include draft tools", () => {
      const names = TOOL_METADATA.map((m) => m.name)
      expect(names).toContain("create_draft")
      expect(names).toContain("send_draft")
    })

    it("should mark draft tools as write operations", () => {
      const createDraft = TOOL_METADATA.find((m) => m.name === "create_draft")
      const sendDraft = TOOL_METADATA.find((m) => m.name === "send_draft")
      expect(createDraft?.readOnly).toBe(false)
      expect(sendDraft?.readOnly).toBe(false)
    })

    it("should not include confirm_action", () => {
      const names = TOOL_METADATA.map((m) => m.name)
      expect(names).not.toContain("confirm_action")
    })
  })

  describe("filterTools", () => {
    it("should return all tools when no filters are set", () => {
      const result = filterTools({})
      const nonOrgTools = TOOL_METADATA.filter((m) => !m.orgOnly)
      expect(result.size).toBe(nonOrgTools.length)
    })

    it("should return all tools including org-only when orgMode is enabled", () => {
      const result = filterTools({ orgMode: true })
      expect(result.size).toBe(TOOL_METADATA.length)
    })

    it("should filter by preset", () => {
      const result = filterTools({ presets: ["productivity"], orgMode: true })
      // productivity = mail + calendar + todo + auth (always included)
      for (const name of result) {
        const meta = TOOL_METADATA.find((m) => m.name === name)
        expect(["mail", "calendar", "todo", "auth"]).toContain(meta?.domain)
      }
    })

    it("should include auth tools even with preset filter", () => {
      const result = filterTools({ presets: ["personal"] })
      expect(result.has("get_auth_status")).toBe(true)
      expect(result.has("list_accounts")).toBe(true)
    })

    it("should filter to read-only tools", () => {
      const result = filterTools({ readOnly: true, orgMode: true })
      for (const name of result) {
        const meta = TOOL_METADATA.find((m) => m.name === name)
        expect(meta?.readOnly).toBe(true)
      }
    })

    it("should exclude org-only tools when orgMode is false", () => {
      const result = filterTools({ orgMode: false })
      for (const name of result) {
        const meta = TOOL_METADATA.find((m) => m.name === name)
        expect(meta?.orgOnly).toBe(false)
      }
    })

    it("should filter by regex pattern", () => {
      const result = filterTools({ enabledPattern: "^list_", orgMode: true })
      for (const name of result) {
        expect(name).toMatch(/^list_/)
      }
      expect(result.size).toBeGreaterThan(0)
    })

    it("should combine preset and readOnly filters", () => {
      const result = filterTools({ presets: ["personal"], readOnly: true })
      for (const name of result) {
        const meta = TOOL_METADATA.find((m) => m.name === name)
        expect(meta?.readOnly).toBe(true)
        expect(["mail", "calendar", "contacts", "todo", "files", "onenote", "auth"]).toContain(meta?.domain)
      }
    })

    it("should combine preset and regex filters", () => {
      const result = filterTools({ presets: ["personal"], enabledPattern: "mail|calendar" })
      for (const name of result) {
        expect(name).toMatch(/mail|calendar/)
      }
    })

    it("should exclude all send_* mail tools but keep create_*_draft when requireDraft is true", () => {
      const result = filterTools({ requireDraft: true })
      expect(result.has("send_message")).toBe(false)
      expect(result.has("send_reply")).toBe(false)
      expect(result.has("send_reply_all")).toBe(false)
      expect(result.has("send_forward")).toBe(false)
      expect(result.has("create_draft")).toBe(true)
      expect(result.has("create_reply_draft")).toBe(true)
      expect(result.has("create_reply_all_draft")).toBe(true)
      expect(result.has("create_forward_draft")).toBe(true)
      expect(result.has("send_draft")).toBe(true)
    })

    it("should include all send_* mail tools when requireDraft is false", () => {
      const result = filterTools({ requireDraft: false })
      expect(result.has("send_message")).toBe(true)
      expect(result.has("send_reply")).toBe(true)
      expect(result.has("send_reply_all")).toBe(true)
      expect(result.has("send_forward")).toBe(true)
    })

    it("should include all send_* mail tools when requireDraft is omitted", () => {
      const result = filterTools({})
      expect(result.has("send_message")).toBe(true)
      expect(result.has("send_reply")).toBe(true)
      expect(result.has("send_reply_all")).toBe(true)
      expect(result.has("send_forward")).toBe(true)
    })

    it("should leave non-mail tools untouched when requireDraft is true", () => {
      const result = filterTools({ requireDraft: true })
      expect(result.has("create_event")).toBe(true)
      expect(result.has("list_messages")).toBe(true)
      expect(result.has("get_message")).toBe(true)
    })
  })
})
