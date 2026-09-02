import { describe, expect, it } from "vitest"

import { toolDefinitions } from "../src/tools/definitions"
import { TOOL_METADATA } from "../src/tools/tool-registry"

// toolDefinitions (what gets registered with FastMCP) and TOOL_METADATA (what preset
// and read-only filtering consults) are two hand-maintained lists describing the same
// tools. A tool in one but not the other fails quietly: filterTools derives the allowed
// set from the registry alone, so a tool that is defined but unregistered is silently
// dropped at startup with no error — that is exactly how scan_messages shipped
// invisible, defined and built but filtered out. The reverse is just as quiet: a tool
// registered but undefined is advertised by a preset and never registered.
//
// This supersedes an earlier check that regex-matched `name:` out of index.ts source
// text. That could only see definitions in one file, and went blind the moment they
// moved. Importing the assembled array checks what actually ships.
describe("toolDefinitions ↔ TOOL_METADATA", () => {
  const definitionNames = toolDefinitions.map((t) => t.name)
  const metadataNames = TOOL_METADATA.map((m) => m.name)

  it("describe exactly the same tools", () => {
    expect([...definitionNames].sort()).toEqual([...metadataNames].sort())
  })

  it("agree on each tool's domain", () => {
    for (const tool of toolDefinitions) {
      const meta = TOOL_METADATA.find((m) => m.name === tool.name)
      expect(meta, `no TOOL_METADATA entry for ${tool.name}`).toBeDefined()
      expect(meta!.domain, `domain mismatch for ${tool.name}`).toBe(tool.domain)
    }
  })

  it("agree on which tools are read-only", () => {
    for (const tool of toolDefinitions) {
      const meta = TOOL_METADATA.find((m) => m.name === tool.name)
      expect(meta, `no TOOL_METADATA entry for ${tool.name}`).toBeDefined()
      expect(meta!.readOnly, `readOnly mismatch for ${tool.name}`).toBe(tool.readOnly)
    }
  })

  it("declares no tool twice", () => {
    expect(new Set(definitionNames).size).toBe(definitionNames.length)
    expect(new Set(metadataNames).size).toBe(metadataNames.length)
  })

  // A read-only tool that forgets readOnlyHint is advertised to clients as if it
  // mutates, which suppresses auto-approval for a tool that only reads.
  it("marks every read-only tool with readOnlyHint", () => {
    for (const tool of toolDefinitions.filter((t) => t.readOnly)) {
      expect(tool.annotations?.readOnlyHint, `${tool.name} is readOnly but lacks readOnlyHint`).toBe(true)
    }
  })

  it("never marks a write tool as readOnlyHint", () => {
    for (const tool of toolDefinitions.filter((t) => !t.readOnly)) {
      expect(tool.annotations?.readOnlyHint, `${tool.name} is a write tool but claims readOnlyHint`).not.toBe(true)
    }
  })

  it("gives every tool a non-empty description", () => {
    for (const tool of toolDefinitions) {
      expect(tool.description.length, `${tool.name} has no description`).toBeGreaterThan(0)
    }
  })
})
