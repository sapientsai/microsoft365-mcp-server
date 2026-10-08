import { Some } from "functype"
import { Right } from "functype/either"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

vi.mock("../src/client/graph-client", () => ({
  getGraphClient: vi.fn(),
}))

import { getGraphClient } from "../src/client/graph-client"
import { graphQuery, isDirectMailSend } from "../src/tools/graph-query-tools"

const mockClient = {
  graphQuery: vi.fn(),
}

beforeEach(() => {
  vi.clearAllMocks()
  vi.mocked(getGraphClient).mockReturnValue(Some(mockClient as never))
  mockClient.graphQuery.mockResolvedValue(Right({ ok: true }))
})

afterEach(() => {
  vi.unstubAllEnvs()
})

describe("isDirectMailSend", () => {
  it.each([
    "/me/sendMail",
    "/users/it@civala.com/sendMail",
    "/users('it@civala.com')/sendMail",
    "/me/messages/m1/reply",
    "/me/messages/m1/replyAll",
    "/me/messages/m1/forward",
    "/me/mailFolders/inbox/messages/m1/forward",
    "/me/messages('m1')/microsoft.graph.reply",
    "/ME/SENDMAIL",
    "/me/sendMail/",
    "/me/sendMail?foo=bar",
    "/me/send%4Dail",
    "/me/sendMail()",
    "/me%2FsendMail",
    "/groups/g1/threads/t1/reply",
    "/groups/g1/conversations/c1/threads/t1/reply",
    "/groups/g1/threads/t1/posts/p1/reply",
    "/groups/g1/threads/t1/posts/p1/forward",
    "/me/events/e1/forward",
    "https://graph.microsoft.com/beta/me/sendMail",
  ])("should treat %s as a direct send", (path) => {
    expect(isDirectMailSend("POST", path)).toBe(true)
  })

  // fetch sends the WHATWG-normalized URL, not the raw string, so each of these goes out as a send.
  it.each([
    ["a fragment", "/me/sendMail#x"],
    ["a tab", "/me/send\tMail"],
    ["a newline", "/me/send\nMail"],
    ["a trailing space", "/me/sendMail "],
    ["a backslash", "/me\\sendMail"],
    ["a dot-dot segment", "/me/sendMail/x/.."],
    ["a dot segment", "/me/sendMail/."],
    ["an encoded dot segment", "/me/sendMail/%2e"],
    ["a dot-dot after reply", "/me/messages/m1/reply/x/.."],
  ])("should see through %s", (_label, path) => {
    expect(isDirectMailSend("POST", `https://graph.microsoft.com/v1.0${path}`)).toBe(true)
  })

  it.each([
    "/me/messages",
    "/me/messages/m1",
    "/me/messages/m1/createReply",
    "/me/messages/m1/createReplyAll",
    "/me/messages/m1/createForward",
    "/teams/t1/channels/c1/messages/m1/replies",
    "/me/events/e1/accept",
  ])("should not treat %s as a direct send", (path) => {
    expect(isDirectMailSend("POST", path)).toBe(false)
  })

  it("should fold Unicode lookalikes the way .NET uppercasing does", () => {
    expect(isDirectMailSend("POST", "/me/\u017FendMail")).toBe(true)
  })

  // A drive item addressed by path can be named like a send action; reading it must still work.
  it.each(["GET", "get", "HEAD"])("should never treat a %s as a send", (method) => {
    expect(isDirectMailSend(method, "/me/drive/root:/Projects/Forward")).toBe(false)
    expect(isDirectMailSend(method, "/me/sendMail")).toBe(false)
  })

  it.each([undefined, "PATCH", "PUT", "DELETE", "MERGE"])(
    "should treat method %s on a send path as a send",
    (method) => {
      expect(isDirectMailSend(method, "/me/sendMail")).toBe(true)
    },
  )

  it("should only let GET requests inside a $batch through", () => {
    expect(isDirectMailSend("POST", "/$batch", { requests: [{ method: "GET", url: "/me/drive/root:/Reply" }] })).toBe(
      false,
    )
    expect(isDirectMailSend("POST", "/$batch", { requests: [{ url: "/me/sendMail" }] })).toBe(true)
  })

  // send_draft stays available under MS365_REQUIRE_DRAFT, so sending a draft through
  // graph_query has to stay allowed too. Blocking it would only push callers to the tool.
  it("should allow sending an existing draft", () => {
    expect(isDirectMailSend("POST", "/me/messages/m1/send")).toBe(false)
  })

  it("should catch a send hidden inside a $batch", () => {
    const body = {
      requests: [
        { id: "1", method: "GET", url: "/me/messages" },
        { id: "2", method: "POST", url: "me/sendMail", body: {} },
      ],
    }
    expect(isDirectMailSend("POST", "/$batch", body)).toBe(true)
  })

  it("should read $batch property names case-insensitively", () => {
    const body = { Requests: [{ Id: "1", Method: "POST", Url: "/me/sendMail#x" }] }
    expect(isDirectMailSend("POST", "/$batch", body)).toBe(true)
  })

  it("should allow a $batch with no sends", () => {
    const body = { requests: [{ id: "1", method: "POST", url: "/me/messages/m1/createReply" }] }
    expect(isDirectMailSend("POST", "/$batch", body)).toBe(false)
  })

  it("should ignore a $batch body without a requests array", () => {
    expect(isDirectMailSend("POST", "/$batch", { requests: "nope" })).toBe(false)
    expect(isDirectMailSend("POST", "/$batch")).toBe(false)
  })
})

// graphQuery reads MS365_REQUIRE_DRAFT itself rather than taking it from the tool definition, so
// these tests drive the env var: dropping the check anywhere on the path fails them.
describe("graphQuery", () => {
  it("should refuse a direct send when MS365_REQUIRE_DRAFT is on, without calling Graph", async () => {
    vi.stubEnv("MS365_REQUIRE_DRAFT", "true")
    const result = await graphQuery({ method: "POST", path: "/me/sendMail", body: JSON.stringify({ message: {} }) })

    expect(result.isLeft()).toBe(true)
    expect(
      result.fold(
        (e) => e.message,
        () => "",
      ),
    ).toContain("send_draft")
    expect(mockClient.graphQuery).not.toHaveBeenCalled()
  })

  it("should pass a direct send through when MS365_REQUIRE_DRAFT is off", async () => {
    vi.stubEnv("MS365_REQUIRE_DRAFT", "false")
    const result = await graphQuery({ method: "POST", path: "/me/sendMail", body: "{}" })

    expect(result.isRight()).toBe(true)
    expect(mockClient.graphQuery).toHaveBeenCalledWith("POST", "/me/sendMail", {}, undefined, undefined)
  })

  it("should allow other writes when MS365_REQUIRE_DRAFT is on", async () => {
    vi.stubEnv("MS365_REQUIRE_DRAFT", "true")
    const result = await graphQuery({ method: "POST", path: "/me/messages/m1/createReply", body: "{}" })

    expect(result.isRight()).toBe(true)
    expect(mockClient.graphQuery).toHaveBeenCalledTimes(1)
  })

  // The version is spliced into the URL ahead of the path, so it could otherwise smuggle in a
  // send that the path check never sees.
  it("should refuse an unknown version before building the URL", async () => {
    vi.stubEnv("MS365_REQUIRE_DRAFT", "true")
    const result = await graphQuery({ method: "POST", path: "/me", version: "v1.0/me/sendMail#" })

    expect(result.isLeft()).toBe(true)
    expect(
      result.fold(
        (e) => e.message,
        () => "",
      ),
    ).toContain("version must be one of")
    expect(mockClient.graphQuery).not.toHaveBeenCalled()
  })

  it("should refuse a normalized send when MS365_REQUIRE_DRAFT is on", async () => {
    vi.stubEnv("MS365_REQUIRE_DRAFT", "true")
    const result = await graphQuery({ method: "POST", path: "/me/sendMail/x/..", version: "beta" })

    expect(result.isLeft()).toBe(true)
    expect(mockClient.graphQuery).not.toHaveBeenCalled()
  })

  it("should return an error, not throw, when body is not JSON", async () => {
    const result = await graphQuery({ method: "POST", path: "/me/messages", body: "{not json" })

    expect(result.isLeft()).toBe(true)
    expect(
      result.fold(
        (e) => e.message,
        () => "",
      ),
    ).toContain("not valid JSON")
    expect(mockClient.graphQuery).not.toHaveBeenCalled()
  })
})
