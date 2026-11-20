import { Left, Right } from "functype/either"
import { describe, expect, it } from "vitest"

import { requestProblem } from "../src/request-check"
import { uploadPathProblem } from "../src/upload/upload"
import { drivePath, encodedId, graphPath, graphPathProblem, idSegment, raw } from "../src/utils/graph-path"

const value = (built: { isLeft: () => boolean; value: unknown }) => built.value
const message = (built: { value: unknown }) => (built.value as { message: string }).message

describe("idSegment", () => {
  it.each([
    ["an empty ID", ""],
    ["a dot", "."],
    ["two dots", ".."],
    ["a slash", "AAA/forward"],
    ["a backslash", "AAA\\forward"],
    ["a lone surrogate", "A\ud800B"],
  ])("refuses %s", (_, id) => {
    const result = idSegment(id)
    expect(result.isLeft()).toBe(true)
    expect(value(result)).toMatchObject({ type: "api", status: 400 })
  })

  it.each([
    ["a question mark", "A?b=1", "A%3Fb=1"],
    ["a hash", "A#b", "A%23b"],
    ["a percent sign", "%2e%2e", "%252e%252e"],
    ["a tab", ".\t.", ".%09."],
    ["a space", "a b", "a%20b"],
    ["a guest UPN", "x_y.com#EXT#@t.onmicrosoft.com", "x_y.com%23EXT%23@t.onmicrosoft.com"],
  ])("encodes %s so it stays inside the segment", (_, id, encoded) => {
    expect(value(idSegment(id))).toBe(encoded)
  })

  it.each([
    ["a message ID", "AAMkAGI2TG93AAA-_x="],
    ["a base64 ID with a plus", "AQAA+dZqA="],
    ["a chat ID", "19:meeting_abc@thread.v2"],
    ["a site ID", "contoso.sharepoint.com,1c2a-11,9f3b-22"],
    ["a drive ID", "b!x-Yz_09"],
    ["a OneNote ID", "1-abc!123"],
    ["a UPN", "ada@contoso.com"],
  ])("leaves %s as it is", (_, id) => {
    expect(value(idSegment(id))).toBe(id)
  })

  it("quotes a refused ID on one line", () => {
    const text = message(idSegment("A/\nForged: line"))
    expect(text.split("\n")).toHaveLength(1)
    expect(text).toContain('"A/\\nForged: line"')
  })
})

describe("graphPath", () => {
  it("puts each ID in as one segment and leaves raw text as written", () => {
    expect(value(graphPath`${raw("/users/a%40b.com")}/messages/${"A#b"}/createForward`)).toBe(
      "/users/a%40b.com/messages/A%23b/createForward",
    )
  })

  it("refuses the ID that used to turn a forward draft into a send", () => {
    const built = graphPath`${raw("/me")}/messages/${"AAA/forward#"}/createForward`
    expect(built.isLeft()).toBe(true)
  })

  it("nests another built path as written, and passes its Left on", () => {
    const inner = graphPath`/me/calendars/${"cal-1"}/events`
    expect(value(graphPath`${inner}/${"ev-1"}`)).toBe("/me/calendars/cal-1/events/ev-1")
    const refused = Left({ type: "api" as const, status: 400, message: "no" })
    expect(value(graphPath`${refused}/${"ev-1"}`)).toEqual({ type: "api", status: 400, message: "no" })
    expect(value(graphPath`${Right("/x")}/y`)).toBe("/x/y")
  })

  it("encodes an encodedId exactly as encodeURIComponent does, slash included", () => {
    expect(value(graphPath`/rules/${encodedId("AQAA/AJ5+dZqA=")}`)).toBe("/rules/AQAA%2FAJ5%2BdZqA%3D")
    expect(graphPath`/rules/${encodedId("a/../b")}`.isLeft()).toBe(true)
    expect(graphPath`/rules/${encodedId("..")}`.isLeft()).toBe(true)
  })

  it("splits a drive path into encoded pieces and refuses a step up", () => {
    expect(value(graphPath`/me/drive/root:/${drivePath("/Docs//Q1 #1/")}:/children`)).toBe(
      "/me/drive/root:/Docs/Q1%20%231:/children",
    )
    expect(value(graphPath`/root:/${drivePath("a:b")}:/children`)).toBe("/root:/a%3Ab:/children")
    expect(graphPath`/root:/${drivePath("Docs/../../users/x")}:/children`.isLeft()).toBe(true)
    expect(graphPath`/root:/${drivePath("/")}:/children`.isLeft()).toBe(true)
  })
})

describe("graphPathProblem", () => {
  it.each([
    ["a dot segment", "/me/../users/x/messages"],
    ["an encoded dot segment", "/me/%2e%2E/users/x"],
    ["a trailing single dot", "/me/messages/."],
    ["a backslash", "/me\\..\\users"],
    ["a hash", "/me/messages/A/forward#"],
    ["a tab", "/me/messages/.\t."],
    ["a line break in the query", "/me/messages?$filter=a\nb"],
  ])("refuses %s", (_, path) => {
    expect(graphPathProblem(path)).toBeDefined()
  })

  it.each([
    "/me/messages/A/$value",
    "/me/drive/root:/a b/c.txt:/content",
    "/me/drive/root/search(q='..')",
    '/me/messages?$search="a b"',
    "/me/drive/root:/notes..txt:/content",
    "/users/a@b.com/messages?$filter=x eq '..'",
    "/users?$filter=onPremisesSamAccountName eq 'DOMAIN\\ada'",
  ])("allows %s", (path) => {
    expect(graphPathProblem(path)).toBeUndefined()
  })
})

describe("the path check on every resilient fetch", () => {
  it("refuses a backslash in the authority, which fetch would read as the end of the host", () => {
    expect(requestProblem("https://graph.microsoft.com\\@evil.example/x", undefined)).toBeDefined()
  })

  it("refuses a full URL that fetch would resolve elsewhere", () => {
    expect(requestProblem("https://graph.microsoft.com/v1.0/me/../../beta/users/x", undefined)).toMatch(/"\.\."/)
    expect(requestProblem("https://graph.microsoft.com/v1.0/me/messages/A#x", undefined)).toMatch(/"#"/)
  })

  it("allows Graph's own continuation and upload-session URLs", () => {
    const nextLink = "https://graph.microsoft.com/v1.0/me/messages?$skiptoken=abc%3D%3D&$top=50"
    const session = "https://contoso.sharepoint.com/_api/v2.0/drive/items/X/uploadSession?guid=a&path=~tmpA_b.txt"
    expect(requestProblem(nextLink, undefined)).toBeUndefined()
    expect(requestProblem(session, undefined)).toBeUndefined()
  })
})

describe("uploadPathProblem", () => {
  it.each([
    ["a non-content path", "/me/sendMail"],
    ["a query string", "/me/drive/root:/a.txt:/content?x=1"],
    ["a dot segment", "/me/drive/root:/../../users/x:/content"],
    ["a hash", "/me/drive/root:/a.txt:/content#"],
  ])("refuses %s", (_, path) => {
    expect(uploadPathProblem(path)).toBeDefined()
  })

  it.each(["/me/drive/root:/a b.txt:/content", "/drives/b!x/items/ITEM/content", "/me/drive/root:/A.TXT:/CONTENT"])(
    "allows %s",
    (path) => {
      expect(uploadPathProblem(path)).toBeUndefined()
    },
  )
})
