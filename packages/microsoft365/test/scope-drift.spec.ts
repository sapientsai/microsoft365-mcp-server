import jwt from "jsonwebtoken"
import { describe, expect, it } from "vitest"

import { describeScopeDrift, missingScopes, parseGrantedScopes, resolveRequiredScopes } from "../src/auth/scope-drift"

const tokenWith = (claims: Record<string, unknown>): string => jwt.sign(claims, "test-secret")

describe("parseGrantedScopes", () => {
  it("reads space-separated delegated scopes from scp", () => {
    expect(parseGrantedScopes(tokenWith({ scp: "Mail.ReadWrite User.Read" }))).toEqual(["Mail.ReadWrite", "User.Read"])
  })

  it("reads app-only permissions from roles", () => {
    expect(parseGrantedScopes(tokenWith({ roles: ["Mail.ReadWrite"] }))).toEqual(["Mail.ReadWrite"])
  })

  // An unreadable token must not look like "nothing granted" — see missingScopes.
  it("returns nothing for a token it cannot read", () => {
    expect(parseGrantedScopes("not-a-jwt")).toEqual([])
    expect(parseGrantedScopes(tokenWith({}))).toEqual([])
  })
})

describe("missingScopes", () => {
  it("finds a required scope that was not granted", () => {
    expect(missingScopes(["Mail.ReadWrite"], ["Mail.ReadWrite.Shared"])).toEqual(["Mail.ReadWrite.Shared"])
  })

  it("is satisfied when everything required is present", () => {
    expect(missingScopes(["Mail.ReadWrite", "Mail.ReadWrite.Shared"], ["Mail.ReadWrite.Shared"])).toEqual([])
  })

  it("ignores casing differences between Azure and the token", () => {
    expect(missingScopes(["mail.readwrite.shared"], ["Mail.ReadWrite.Shared"])).toEqual([])
  })

  // Extra permissions are legitimate — an app registration may serve other clients — so
  // they must never trigger a re-acquisition.
  it("does not treat extra granted scopes as drift", () => {
    expect(missingScopes(["Mail.ReadWrite.Shared", "Files.Read"], ["Mail.ReadWrite.Shared"])).toEqual([])
  })

  // Failing open matters: treating an unparseable token as "everything missing" would
  // re-authenticate every deployment whose token shape this does not recognise.
  it("reports no drift when the granted set could not be read", () => {
    expect(missingScopes([], ["Mail.ReadWrite.Shared"])).toEqual([])
  })

  it("requires nothing when nothing is required", () => {
    expect(missingScopes(["Mail.Read"], [])).toEqual([])
  })
})

describe("resolveRequiredScopes", () => {
  it("requires no shared scope when no other mailbox is configured", () => {
    expect(resolveRequiredScopes({} as NodeJS.ProcessEnv)).toEqual([])
    expect(resolveRequiredScopes({ MS365_ALLOWED_MAILBOXES: "  " } as NodeJS.ProcessEnv)).toEqual([])
  })

  it("requires the shared write scope when other mailboxes are allowed", () => {
    expect(resolveRequiredScopes({ MS365_ALLOWED_MAILBOXES: "bel@example.com" } as NodeJS.ProcessEnv)).toEqual([
      "Mail.ReadWrite.Shared",
    ])
  })

  // App-only tokens carry roles, not delegated scopes: Mail.ReadWrite as an application
  // permission already reaches whatever the tenant policy allows, and no .Shared variant
  // exists to grant. Requiring one would be an unfixable false positive.
  it.each(["client-secret", "certificate"])("requires no shared scope in %s mode", (mode) => {
    expect(
      resolveRequiredScopes({
        MS365_ALLOWED_MAILBOXES: "bel@example.com",
        MS365_AUTH_MODE: mode,
      } as NodeJS.ProcessEnv),
    ).toEqual([])
  })

  it("still requires the shared scope in the delegated modes", () => {
    for (const mode of ["interactive", "oauth-proxy", "client-token"]) {
      expect(
        resolveRequiredScopes({
          MS365_ALLOWED_MAILBOXES: "bel@example.com",
          MS365_AUTH_MODE: mode,
        } as NodeJS.ProcessEnv),
      ).toEqual(["Mail.ReadWrite.Shared"])
    }
  })

  // A read-only deployment cannot write anywhere, so demanding the write scope would
  // report drift that no permission change could ever satisfy.
  it("requires only the shared read scope in read-only mode", () => {
    expect(
      resolveRequiredScopes({
        MS365_ALLOWED_MAILBOXES: "bel@example.com",
        MS365_READ_ONLY: "true",
      } as NodeJS.ProcessEnv),
    ).toEqual(["Mail.Read.Shared"])
  })
})

describe("describeScopeDrift", () => {
  // The fix is in Azure, not here: re-running the server cannot grant a permission, so
  // the message has to say where to go.
  it("points at the app registration rather than at a retry", () => {
    const message = describeScopeDrift(["Mail.ReadWrite.Shared"])
    expect(message).toContain("Mail.ReadWrite.Shared")
    expect(message).toContain("app registration")
    expect(message).toContain("consent")
  })

  // Granting the permission looks like it should be enough, and is not: `.default` keeps
  // matching the cached token until it expires. Both steps have to be stated.
  it("says the cached token must be cleared as well", () => {
    const message = describeScopeDrift(["Mail.ReadWrite.Shared"])
    expect(message).toContain("token cache")
    expect(message).toContain("will not take effect")
  })

  it("names the cache directory when it is known", () => {
    expect(describeScopeDrift(["Mail.ReadWrite.Shared"], "/tmp/token-cache")).toContain("/tmp/token-cache")
  })
})
