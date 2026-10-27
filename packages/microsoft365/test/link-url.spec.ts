import { describe, expect, it } from "vitest"

import { displayableLinkUrl, linkLocation } from "../src/utils/link-url"

describe("displayableLinkUrl", () => {
  it.each([
    ["https://contoso.sharepoint.com/:f:/s/team/abc", "https://contoso.sharepoint.com/:f:/s/team/abc"],
    ["http://example.com/a b", "http://example.com/a%20b"],
    ["https://example.com/a\nb\r\nc", "https://example.com/abc"],
    ["https://example.com/p\u0001q?x=\u0002y", "https://example.com/p%01q?x=%02y"],
    ["https://example.com/\u2028x", "https://example.com/%E2%80%A8x"],
  ])("%j -> %s", (raw, expected) => {
    expect(displayableLinkUrl(raw)).toBe(expected)
  })

  it.each([
    "javascript:alert(1)",
    "file:///etc/passwd",
    "data:text/html,x",
    "mailto:a@example.com",
    "",
    "relative/path",
  ])("rejects %j", (raw) => {
    expect(displayableLinkUrl(raw)).toBeUndefined()
  })
})

describe("linkLocation", () => {
  it("separates a missing link from a withheld one", () => {
    expect(linkLocation({ id: "a" })).toBe("link not returned by Graph")
    expect(linkLocation({ id: "a", sourceUrl: "javascript:x" })).toBe("link withheld: not an http(s) URL")
    expect(linkLocation({ id: "a", sourceUrl: "https://example.com/x" })).toBe("https://example.com/x")
  })
})
