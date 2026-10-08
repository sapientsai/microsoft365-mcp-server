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
    ["POST", "/me/sendMail"],
    ["POST", "/users/it@civala.com/sendMail"],
    ["POST", "/me/messages/m1/reply"],
    ["POST", "/me/messages/m1/replyAll"],
    ["POST", "/me/messages/m1/forward"],
    ["POST", "/me/mailFolders/inbox/messages/m1/forward"],
    ["POST", "/me/messages('m1')/microsoft.graph.reply"],
    ["post", "/ME/SENDMAIL"],
    ["POST", "me/sendMail/"],
    ["POST", "/me/sendMail?foo=bar"],
    ["POST", "/me/send%4Dail"],
  ])("should treat %s %s as a direct send", (method, path) => {
    expect(isDirectMailSend(method, path)).toBe(true)
  })

  it.each([
    ["GET", "/me/messages"],
    ["POST", "/me/messages"],
    ["POST", "/me/messages/m1/createReply"],
    ["POST", "/me/messages/m1/createReplyAll"],
    ["POST", "/me/messages/m1/createForward"],
    ["PATCH", "/me/messages/m1"],
    ["DELETE", "/me/messages/m1"],
    ["POST", "/me/events/e1/forward"],
  ])("should not treat %s %s as a direct send", (method, path) => {
    expect(isDirectMailSend(method, path)).toBe(false)
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
        { id: "2", method: "POST", url: "/me/sendMail", body: {} },
      ],
    }
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
