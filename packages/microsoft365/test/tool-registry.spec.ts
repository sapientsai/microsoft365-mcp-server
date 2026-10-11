import { readFileSync } from "node:fs"

import { describe, expect, it } from "vitest"

import { DOMAIN_DESCRIPTIONS, filterTools, PRESETS, TOOL_METADATA } from "../src/tools/tool-registry"

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

    // registerTools only registers a definition whose name filterTools returns, and filterTools reads
    // TOOL_METADATA. A tool defined in index.ts but missing here never registers, with no error. The
    // definitions can't be imported (loading index.ts starts the server), so their names are read from
    // the source: each definition's name sits at four spaces of indent.
    it("should list every tool defined in index.ts, and nothing else", () => {
      const source = readFileSync(new URL("../src/index.ts", import.meta.url), "utf8")
      const defined = [...source.matchAll(/^ {4}name: "([a-z_]+)",$/gm)].map((match) => match[1]).sort()
      const registered = TOOL_METADATA.map((m) => m.name).sort()
      expect(defined.length).toBeGreaterThan(0)
      expect(defined).toEqual(registered)
    })

    it("should keep the domain and readOnly flag of each tool in step with its definition", () => {
      const source = readFileSync(new URL("../src/index.ts", import.meta.url), "utf8")
      const blocks = source.split(/^ {4}name: "/m).slice(1)
      const fromSource = blocks.map((block) => ({
        name: block.slice(0, block.indexOf('"')),
        domain: /^ {4}domain: "([a-z]+)"/m.exec(block)?.[1],
        readOnly: /^ {4}readOnly: (true|false),$/m.exec(block)?.[1] === "true",
      }))
      const fromMetadata = new Map(TOOL_METADATA.map((m) => [m.name, m]))
      const mismatched = fromSource.filter((tool) => {
        const meta = fromMetadata.get(tool.name)
        return meta === undefined || meta.domain !== tool.domain || meta.readOnly !== tool.readOnly
      })
      expect(mismatched).toEqual([])
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

  // Issue #57. The capability summary the model reads is built from these descriptions, and the
  // `rag` domain never had one — so the deployed server listed every capability except reading
  // documents, which is the one it exists for. The old map was Record<string, string> consumed
  // with .filter(Boolean), so the gap could not fail: it just produced a shorter list.
  //
  // Record<ToolDomain, string> now makes a missing description a typecheck error, which is the
  // real guard. These cover what types cannot: that a description is not blank, and that every
  // domain actually carrying tools is represented.
  describe("domain descriptions", () => {
    it("describes every domain that has tools registered", () => {
      const domainsWithTools = [...new Set(TOOL_METADATA.map((t) => t.domain))]
      const undescribed = domainsWithTools.filter((d) => (DOMAIN_DESCRIPTIONS[d] ?? "").trim().length === 0)
      expect(
        undescribed,
        `These domains have tools but no usable description, so their tools are absent from the ` +
          `server's capability summary and a model never learns they exist: ${undescribed.join(", ")}`,
      ).toEqual([])
    })

    it("describes the rag domain, which read_document lives in", () => {
      const ragTools = TOOL_METADATA.filter((t) => t.domain === "rag").map((t) => t.name)
      expect(ragTools).toContain("read_document")
      expect(DOMAIN_DESCRIPTIONS.rag.toLowerCase()).toContain("text")
    })
  })
})

// A tool absent from TOOL_METADATA can never be exposed: filterTools derives the
// allowed set from the registry alone, so a tool defined in index.ts but not
// registered here is silently dropped at startup with no error. That is exactly how
// scan_messages shipped invisible — defined, built, and filtered out.
describe("registry covers every defined tool", () => {
  it("registers every tool declared in index.ts", async () => {
    const { readFileSync } = await import("node:fs")
    const { fileURLToPath } = await import("node:url")
    const { dirname, join } = await import("node:path")

    const here = dirname(fileURLToPath(import.meta.url))
    const source = readFileSync(join(here, "../src/index.ts"), "utf8")

    // Tool definitions are object literals whose first key is `name:`.
    const defined = [...source.matchAll(/^\s{4}name: "([a-z0-9_]+)",$/gm)].map((m) => m[1]!)
    expect(defined.length, "found no tool definitions to check").toBeGreaterThan(20)

    const registered = new Set(TOOL_METADATA.map((m) => m.name))
    const missing = defined.filter((name) => !registered.has(name))

    expect(missing, `defined in index.ts but missing from TOOL_METADATA: ${missing.join(", ")}`).toEqual([])
  })
})
