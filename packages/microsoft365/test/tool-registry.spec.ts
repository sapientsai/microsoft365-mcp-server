import { readFileSync } from "node:fs"

import { describe, expect, it } from "vitest"

import {
  annotationsFor,
  DOMAIN_DESCRIPTIONS,
  filterTools,
  PRESETS,
  TOOL_METADATA,
  toolAnnotations,
} from "../src/tools/tool-registry"

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
    // TOOL_METADATA. A tool defined in index.ts but missing here never registers, with no error; that
    // is how scan_messages first shipped invisible, defined and built but filtered out. The
    // definitions can't be imported (loading index.ts starts the server), so their names are read from
    // the source: each definition's name sits at four spaces of indent.
    it("should list every tool defined in index.ts, and nothing else", () => {
      const source = readFileSync(new URL("../src/index.ts", import.meta.url), "utf8")
      const defined = [...source.matchAll(/^ {4}name: "([a-z0-9_]+)",$/gm)].map((match) => match[1]).sort()
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
      expect(fromSource.length).toBe(TOOL_METADATA.length)
      const fromMetadata = new Map(TOOL_METADATA.map((m) => [m.name, m]))
      const mismatched = fromSource.filter((tool) => {
        const meta = fromMetadata.get(tool.name)
        return meta === undefined || meta.domain !== tool.domain || meta.readOnly !== tool.readOnly
      })
      expect(mismatched).toEqual([])
    })
  })

  // Anthropic's connector review asks every tool for a title and the applicable hint; Claude uses the
  // hints to decide when to confirm a call. They are built from TOOL_METADATA alone.
  describe("annotations", () => {
    it("gives every tool a distinct, non-empty title", () => {
      const titles = TOOL_METADATA.map((m) => m.title.trim())
      expect(titles.filter((t) => t.length === 0)).toEqual([])
      expect(new Set(titles).size).toBe(titles.length)
    })

    it("marks every read-only tool readOnlyHint: true, with no destructive hint", () => {
      const wrong = TOOL_METADATA.filter((m) => m.readOnly).filter((m) => {
        const a = toolAnnotations(m)
        return a.readOnlyHint !== true || a.destructiveHint !== undefined
      })
      expect(wrong.map((m) => m.name)).toEqual([])
    })

    it("gives every write an explicit destructive hint, true unless it is additive", () => {
      const writes = TOOL_METADATA.filter((m) => !m.readOnly)
      const wrong = writes.filter((m) => {
        const a = toolAnnotations(m)
        return a.readOnlyHint !== false || a.destructiveHint !== (m.additive !== true)
      })
      expect(wrong.map((m) => m.name)).toEqual([])
    })

    it("never marks a read-only tool additive", () => {
      expect(TOOL_METADATA.filter((m) => m.readOnly && m.additive).map((m) => m.name)).toEqual([])
    })

    // destructiveHint: false may let a client run a tool without asking, so the set is kept small
    // and changes to it should be deliberate.
    it("limits additive writes to ones that only create something private", () => {
      expect(
        TOOL_METADATA.filter((m) => m.additive)
          .map((m) => m.name)
          .sort(),
      ).toEqual([
        "copy_onenote_page",
        "create_contact",
        "create_draft",
        "create_folder",
        "create_forward_draft",
        "create_onenote_notebook",
        "create_onenote_page",
        "create_onenote_section",
        "create_reply_all_draft",
        "create_reply_draft",
        "create_todo_task",
        "get_upload_config",
      ])
    })

    it("keeps every send, update, delete and move destructive", () => {
      const risky = TOOL_METADATA.filter((m) => /^(send|update|delete|move|batch_move|respond)_/.test(m.name))
      expect(risky.length).toBeGreaterThan(10)
      expect(risky.filter((m) => toolAnnotations(m).destructiveHint !== true).map((m) => m.name)).toEqual([])
    })

    it("flags graph_query as open-world and destructive", () => {
      expect(annotationsFor("graph_query")).toEqual({
        title: "Raw Microsoft Graph Request",
        readOnlyHint: false,
        destructiveHint: true,
        openWorldHint: true,
      })
    })

    // A second source of hints is how they drift apart; registerTools reads only the registry.
    it("leaves annotations out of the tool definitions in index.ts", () => {
      const source = readFileSync(new URL("../src/index.ts", import.meta.url), "utf8")
      expect(source.match(/^ {4}annotations:/gm) ?? []).toEqual([])
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
