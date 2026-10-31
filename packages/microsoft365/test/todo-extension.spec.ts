import { describe, expect, it } from "vitest"

import {
  extensionExpand,
  extensionFields,
  extensionProblem,
  findExtension,
  toExtensionBody,
  toExtensionPatch,
} from "../src/utils/todo-extension"

describe("extensionProblem", () => {
  it("accepts a reverse-DNS name with flat fields", () => {
    expect(extensionProblem({ name: "com.civala.assistant", fields: { rank: 1, tier: "top", pinned: true } })).toBe(
      undefined,
    )
  })

  it.each([
    ["no dot", "assistant"],
    ["a quote, which would break out of the $filter", "com.civala'x"],
    ["a space", "com.civala assistant"],
    ["a leading digit", "1com.civala"],
    ["too many characters", `com.${"a".repeat(100)}`],
  ])("refuses a name with %s", (_label, name) => {
    expect(extensionProblem({ name, fields: {} })).toContain("reverse DNS")
  })

  it.each([
    ["a nested object", { a: { b: 1 } }, "string, number or boolean"],
    ["null, which Graph refuses", { a: null }, "not null"],
    ["an array", { a: [1] }, "string, number or boolean"],
    ["a non-finite number", { a: Number.POSITIVE_INFINITY }, "finite"],
    ["a key Graph reads as an annotation", { "rank@odata.type": "x" }, "field name"],
    ["a reserved key", { id: "x" }, "field name"],
  ])("refuses %s", (_label, fields, message) => {
    expect(extensionProblem({ name: "com.civala.assistant", fields: fields as never })).toContain(message)
  })

  it("refuses more than 20 fields", () => {
    const fields = Object.fromEntries(Array.from({ length: 21 }, (_, i) => [`f${i}`, i]))
    expect(extensionProblem({ name: "com.civala.assistant", fields })).toContain("at most 20 fields")
  })

  it("refuses more than 2 KB as JSON", () => {
    expect(extensionProblem({ name: "com.civala.assistant", fields: { note: "x".repeat(2100) } })).toContain(
      "2048 bytes",
    )
  })
})

describe("extension shapes", () => {
  // As Graph returned it live on 2026-10-30.
  const read = {
    extensionName: "com.claude.test",
    id: "microsoft.graph.openTypeExtension.com.claude.test",
    "rank@odata.type": "#Int64",
    rank: 1,
    tier: "top",
    pinned: true,
  }

  it("keeps only the caller's fields on read", () => {
    expect(extensionFields(read)).toEqual({ rank: 1, tier: "top", pinned: true })
  })

  // Another writer could store a key this server would refuse; it must not reach the line.
  it("drops a key it wouldn't accept on input", () => {
    expect(extensionFields({ ...read, "bad\u2028key": "x", "with.dot": 1 })).toEqual({
      rank: 1,
      tier: "top",
      pinned: true,
    })
  })

  it("finds an extension by extensionName, case-insensitively, never by id", () => {
    expect(findExtension([read], "COM.claude.test")).toBe(read)
    expect(
      findExtension([{ id: "microsoft.graph.openTypeExtension.com.claude.test", rank: 1 }], "com.claude.test"),
    ).toBe(undefined)
    expect(findExtension(undefined, "com.claude.test")).toBe(undefined)
  })

  it("builds the create and patch bodies Graph takes", () => {
    expect(toExtensionBody("com.claude.test", { rank: 1 })).toEqual({
      "@odata.type": "microsoft.graph.openTypeExtension",
      extensionName: "com.claude.test",
      rank: 1,
    })
    expect(toExtensionPatch({ rank: 2 })).toEqual({ "@odata.type": "microsoft.graph.openTypeExtension", rank: 2 })
  })

  it("filters the expand by the bare name", () => {
    expect(extensionExpand("com.claude.test")).toBe("extensions($filter=id eq 'com.claude.test')")
  })
})
