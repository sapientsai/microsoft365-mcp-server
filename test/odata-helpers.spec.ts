import { describe, expect, it } from "vitest"

import { buildODataQuery } from "../src/utils/odata-helpers"

describe("buildODataQuery", () => {
  it("should return empty string for no params", () => {
    expect(buildODataQuery()).toBe("")
    expect(buildODataQuery({})).toBe("")
  })

  it("should build $select query", () => {
    const result = buildODataQuery({ $select: ["id", "displayName", "mail"] })
    expect(result).toBe("?$select=id,displayName,mail")
  })

  it("should build $filter query", () => {
    const result = buildODataQuery({ $filter: "displayName eq 'Test'" })
    expect(result).toContain("$filter=")
    expect(result).toContain("displayName")
  })

  it("should build $top and $skip query", () => {
    const result = buildODataQuery({ $top: 10, $skip: 20 })
    expect(result).toContain("$top=10")
    expect(result).toContain("$skip=20")
  })

  it("should build $orderby query", () => {
    const result = buildODataQuery({ $orderby: "displayName asc" })
    expect(result).toContain("$orderby=")
  })

  it("should build $search query", () => {
    const result = buildODataQuery({ $search: "test query" })
    expect(result).toContain('$search="')
  })

  it("should build $count query", () => {
    const result = buildODataQuery({ $count: true })
    expect(result).toBe("?$count=true")
  })

  it("should combine multiple params", () => {
    const result = buildODataQuery({ $select: ["id"], $top: 5, $orderby: "id" })
    expect(result).toContain("$select=id")
    expect(result).toContain("$top=5")
    expect(result).toContain("$orderby=")
    expect(result.startsWith("?")).toBe(true)
  })
})
