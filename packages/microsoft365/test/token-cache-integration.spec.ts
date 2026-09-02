import { createRequire } from "node:module"

import { mkdtempSync, readdirSync, rmSync, statSync } from "node:fs"
import { tmpdir } from "node:os"
import { dirname, join } from "node:path"

import { useIdentityPlugin } from "@azure/identity"
import { afterEach, beforeEach, describe, expect, it } from "vitest"

import { fileCachePersistencePlugin } from "../src/auth/token-cache"

// Unit tests exercise our own plugin object. This drives the real @azure/identity
// plugin machinery instead — the path that silently did nothing under the keytar
// implementation, where the plugin failed to load and persistence quietly degraded.
describe("token cache integration with @azure/identity", () => {
  let directory: string

  beforeEach(() => {
    directory = mkdtempSync(join(tmpdir(), "cache-integration-"))
    process.env.MS365_TOKEN_CACHE_PATH = directory
  })

  afterEach(() => {
    delete process.env.MS365_TOKEN_CACHE_PATH
    rmSync(directory, { recursive: true, force: true })
  })

  it("registers with useIdentityPlugin and persists through the real config path", async () => {
    // msalClient reaches the plugin through this control object, which
    // useIdentityPlugin is responsible for populating. Asserting on it proves
    // registration actually took effect — the step that silently no-opped under the
    // keytar implementation.
    const require_ = createRequire(import.meta.url)
    const identityRoot = dirname(require_.resolve("@azure/identity/package.json"))
    const { msalNodeFlowCacheControl } = (await import(
      join(identityRoot, "dist/esm/msal/nodeFlows/msalPlugins.js")
    )) as { msalNodeFlowCacheControl: { setPersistence: (p: unknown) => void } }

    // Capture the provider identity registered, then exercise it exactly as
    // generatePluginConfiguration does internally.
    const captured: Array<(o: { name?: string }) => Promise<unknown>> = []
    msalNodeFlowCacheControl.setPersistence = (provider) =>
      captured.push(provider as (o: { name?: string }) => Promise<unknown>)

    useIdentityPlugin(fileCachePersistencePlugin)
    expect(captured, "useIdentityPlugin did not register a persistence provider").toHaveLength(1)

    const config = { cache: { cachePlugin: captured[0]!({ name: "integration" }) } }

    const plugin = (await config.cache.cachePlugin) as {
      beforeCacheAccess: (c: unknown) => Promise<void>
      afterCacheAccess: (c: unknown) => Promise<void>
    }
    expect(plugin, "identity produced no cache plugin").toBeDefined()

    await plugin.afterCacheAccess({
      cacheHasChanged: true,
      tokenCache: { serialize: () => '{"secret":"xyz"}', deserialize: () => undefined },
    })

    const files = readdirSync(directory)
    expect(files, "nothing was persisted to disk").toHaveLength(1)
    expect(statSync(join(directory, files[0]!)).mode & 0o777).toBe(0o600)

    const seen: string[] = []
    await plugin.beforeCacheAccess({
      cacheHasChanged: false,
      tokenCache: { serialize: () => "", deserialize: (d: string) => seen.push(d) },
    })
    expect(seen).toEqual(['{"secret":"xyz"}'])
  })
})
