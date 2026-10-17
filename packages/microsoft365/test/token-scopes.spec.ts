import { describe, expect, it } from "vitest"

import { parseJwtScopes } from "../src/auth/token-scopes"

const segment = (value: unknown): string => Buffer.from(JSON.stringify(value)).toString("base64url")
const token = (claims: unknown, signature = "sig"): string =>
  `${segment({ alg: "RS256", typ: "JWT" })}.${segment(claims)}.${signature}`

describe("parseJwtScopes", () => {
  it("reads delegated scopes from scp", () => {
    expect(parseJwtScopes(token({ scp: "Mail.Read  User.Read openid" }))).toEqual(["Mail.Read", "User.Read", "openid"])
  })

  it("reads app roles when there is no scp", () => {
    expect(parseJwtScopes(token({ roles: ["Mail.Read", 7, "Sites.Read.All"] }))).toEqual([
      "Mail.Read",
      "Sites.Read.All",
    ])
  })

  it("prefers scp over roles, as Entra ID issues one or the other", () => {
    expect(parseJwtScopes(token({ scp: "Mail.Read", roles: ["X"] }))).toEqual(["Mail.Read"])
  })

  it("accepts a token with an empty signature segment", () => {
    expect(parseJwtScopes(token({ scp: "User.Read" }, ""))).toEqual(["User.Read"])
  })

  // Personal-account tokens are opaque, and a malformed token must never fail a status call.
  it.each([
    ["an opaque token", "EwBwA8l6BAAU7p9QDpi"],
    ["two segments", "abc.def"],
    ["four segments", "a.b.c.d"],
    ["a payload that isn't JSON", `${segment({})}.${Buffer.from("not json").toString("base64url")}.sig`],
    ["a payload that is a string", token("scp")],
    ["no scope claims", token({ aud: "graph" })],
    ["an empty string", ""],
  ])("returns no scopes for %s", (_label, value) => {
    expect(parseJwtScopes(value)).toEqual([])
  })
})
