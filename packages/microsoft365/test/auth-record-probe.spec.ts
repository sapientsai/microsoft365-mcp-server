import { describe, expect, it } from "vitest"

import { InteractiveBrowserCredential } from "@azure/identity"

// readActiveAccount reaches through InteractiveBrowserCredential to its internal
// msalClient. That is private API, so this asserts the shape it depends on still
// exists — if a future @azure/identity changes it, the silent degradation would
// otherwise only show up as "it keeps asking me to sign in".
describe("InteractiveBrowserCredential internals", () => {
  it("exposes an msalClient with getActiveAccount", () => {
    const credential = new InteractiveBrowserCredential({
      tenantId: "test-tenant",
      clientId: "test-client",
    }) as unknown as { msalClient?: { getActiveAccount?: () => unknown } }

    expect(credential.msalClient, "msalClient no longer present on the credential").toBeDefined()
    expect(typeof credential.msalClient?.getActiveAccount, "getActiveAccount is no longer a function").toBe("function")
  })

  it("returns no account before any sign-in", () => {
    const credential = new InteractiveBrowserCredential({
      tenantId: "test-tenant",
      clientId: "test-client",
    }) as unknown as { msalClient: { getActiveAccount: () => unknown } }

    expect(credential.msalClient.getActiveAccount()).toBeUndefined()
  })
})
