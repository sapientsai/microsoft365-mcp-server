import { describe, expect, it } from "vitest"

import { filterTools, PRESETS, TOOL_METADATA } from "../src/tools/tool-registry"

describe("tool-registry", () => {
  describe("PRESETS", () => {
    it("should define expected preset names", () => {
      expect(Object.keys(PRESETS)).toEqual(["personal", "collaboration", "productivity", "rag", "all"])
    })

    it("personal preset should include mail, calendar, contacts, todo, files, onenote", () => {
      expect(PRESETS.personal).toEqual(["mail", "calendar", "contacts", "todo", "files", "onenote"])
    })

    it("collaboration preset should include chats, teams, meetings, planner, groups", () => {
      expect(PRESETS.collaboration).toEqual(["chats", "teams", "meetings", "planner", "groups"])
    })

    it("all preset should list every domain, so a new domain is never silently excluded", () => {
      const domains = new Set(TOOL_METADATA.map((m) => m.domain))
      for (const domain of domains) expect(PRESETS.all).toContain(domain)
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

    // read_document carries the three heavy parsers, so which deployments expose it matters.
    it("exposes read_document by default, hides it under an unrelated preset, includes it under rag", () => {
      expect(filterTools({}).has("read_document")).toBe(true)
      expect(filterTools({ presets: ["personal"] }).has("read_document")).toBe(false)
      expect(filterTools({ presets: ["rag"] }).has("read_document")).toBe(true)
    })

    // MS365_PRESETS="" splits to [""], which is length 1 and therefore engages the preset filter
    // against a preset name that matches nothing — collapsing the surface to auth alone. Pre-existing
    // behavior, pinned here because the unset and named-preset cases both miss it.
    it("collapses to auth-only when a preset string is empty", () => {
      const result = filterTools({ presets: [""] })
      for (const name of result) {
        expect(TOOL_METADATA.find((m) => m.name === name)?.domain).toBe("auth")
      }
      expect(result.has("read_document")).toBe(false)
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
