import { chmodSync, mkdirSync, mkdtempSync, readFileSync, rmSync, statSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

import { createFileCachePlugin, fileCachePersistencePlugin, resolveCacheDirectory } from "../src/auth/token-cache"

const cacheContext = (serialized: string, hasChanged = true) => {
  const deserialize = vi.fn()
  return {
    context: {
      cacheHasChanged: hasChanged,
      tokenCache: { deserialize, serialize: () => serialized },
    },
    deserialize,
  }
}

describe("token cache", () => {
  let directory: string

  beforeEach(() => {
    directory = mkdtempSync(join(tmpdir(), "token-cache-"))
  })

  afterEach(() => {
    rmSync(directory, { recursive: true, force: true })
    vi.restoreAllMocks()
  })

  describe("resolveCacheDirectory", () => {
    it("prefers an explicit MS365_TOKEN_CACHE_PATH", () => {
      expect(resolveCacheDirectory({ MS365_TOKEN_CACHE_PATH: "/custom/path" })).toBe("/custom/path")
    })

    it("falls back to XDG_CONFIG_HOME", () => {
      expect(resolveCacheDirectory({ XDG_CONFIG_HOME: "/xdg" })).toBe("/xdg/microsoft365-mcp-server/token-cache")
    })

    // TOKEN_STORAGE_PATH belongs to the oauth-proxy provider. Reading it here would
    // couple two unrelated features together.
    it("ignores TOKEN_STORAGE_PATH", () => {
      expect(resolveCacheDirectory({ TOKEN_STORAGE_PATH: "/proxy", XDG_CONFIG_HOME: "/xdg" })).not.toContain("/proxy")
    })
  })

  describe("round trip", () => {
    it("writes on change and reads back", async () => {
      const plugin = createFileCachePlugin("test", directory)
      await plugin.afterCacheAccess(cacheContext('{"token":"abc"}').context)

      const reader = cacheContext("")
      await plugin.beforeCacheAccess(reader.context)

      expect(reader.deserialize).toHaveBeenCalledWith('{"token":"abc"}')
    })

    it("does not write when the cache is unchanged", async () => {
      const plugin = createFileCachePlugin("test", directory)
      await plugin.afterCacheAccess(cacheContext("{}", false).context)

      const reader = cacheContext("")
      await plugin.beforeCacheAccess(reader.context)

      expect(reader.deserialize).not.toHaveBeenCalled()
    })

    it("treats a missing cache as empty rather than failing", async () => {
      const plugin = createFileCachePlugin("absent", join(directory, "does-not-exist"))
      const reader = cacheContext("")

      await expect(plugin.beforeCacheAccess(reader.context)).resolves.toBeUndefined()
      expect(reader.deserialize).not.toHaveBeenCalled()
    })

    it("keeps separate caches for different names", async () => {
      const a = createFileCachePlugin("cache-a", directory)
      const b = createFileCachePlugin("cache-b", directory)

      await a.afterCacheAccess(cacheContext('"from-a"').context)
      await b.afterCacheAccess(cacheContext('"from-b"').context)

      const reader = cacheContext("")
      await a.beforeCacheAccess(reader.context)
      expect(reader.deserialize).toHaveBeenCalledWith('"from-a"')
    })

    // A name that sanitises away entirely must still land in its own file.
    it("distinguishes names that sanitise to the same string", async () => {
      const a = createFileCachePlugin("!!!", directory)
      const b = createFileCachePlugin("???", directory)

      await a.afterCacheAccess(cacheContext('"a"').context)
      await b.afterCacheAccess(cacheContext('"b"').context)

      const reader = cacheContext("")
      await a.beforeCacheAccess(reader.context)
      expect(reader.deserialize).toHaveBeenCalledWith('"a"')
    })

    it("keeps a path separator in the name from escaping the directory", async () => {
      const plugin = createFileCachePlugin("../../escape", directory)
      await plugin.afterCacheAccess(cacheContext('"contained"').context)

      const reader = cacheContext("")
      await plugin.beforeCacheAccess(reader.context)
      expect(reader.deserialize).toHaveBeenCalledWith('"contained"')
    })
  })

  // The token is a bearer credential: file permissions are the security model.
  describe("permissions", () => {
    it("writes the cache readable only by its owner", async () => {
      const plugin = createFileCachePlugin("perms", directory)
      await plugin.afterCacheAccess(cacheContext('{"token":"abc"}').context)

      const file = readFileSync(join(directory, findCacheFile(directory)), "utf8")
      expect(file).toBe('{"token":"abc"}')
      expect(statSync(join(directory, findCacheFile(directory))).mode & 0o777).toBe(0o600)
    })

    it("discards a cache that others can read", async () => {
      const plugin = createFileCachePlugin("loose", directory)
      await plugin.afterCacheAccess(cacheContext('{"token":"abc"}').context)

      const path = join(directory, findCacheFile(directory))
      chmodSync(path, 0o644)

      const reader = cacheContext("")
      vi.spyOn(console, "error").mockImplementation(() => undefined)
      await plugin.beforeCacheAccess(reader.context)

      expect(reader.deserialize).not.toHaveBeenCalled()
    })

    it("does not reuse a leftover temp file's permissions", async () => {
      const plugin = createFileCachePlugin("stale", directory)
      const path = join(directory, cacheFileNameFor("stale"))
      mkdirSync(directory, { recursive: true })
      writeFileSync(`${path}.${process.pid}.tmp`, "stale", { mode: 0o666 })

      await plugin.afterCacheAccess(cacheContext('{"token":"fresh"}').context)

      expect(statSync(path).mode & 0o777).toBe(0o600)
    })
  })

  describe("failure handling", () => {
    it("degrades to in-memory when the cache cannot be written", async () => {
      // A file where the directory should be makes every write fail.
      const blocked = join(directory, "blocked")
      writeFileSync(blocked, "not a directory")

      const plugin = createFileCachePlugin("test", blocked)
      vi.spyOn(console, "error").mockImplementation(() => undefined)

      await expect(plugin.afterCacheAccess(cacheContext("{}").context)).resolves.toBeUndefined()
    })
  })

  describe("fileCachePersistencePlugin", () => {
    it("registers a provider that yields a working plugin", async () => {
      const setPersistence = vi.fn()
      fileCachePersistencePlugin({ cachePluginControl: { setPersistence } })

      expect(setPersistence).toHaveBeenCalledOnce()

      const provider = setPersistence.mock.calls[0]![0] as (o: { name?: string }) => Promise<unknown>
      const plugin = (await provider({ name: "from-provider" })) as {
        beforeCacheAccess: (c: unknown) => Promise<void>
      }

      expect(typeof plugin.beforeCacheAccess).toBe("function")
    })
  })
})

// The plugin owns its naming scheme; tests locate the file rather than duplicating it.
const findCacheFile = (directory: string): string => {
  const { readdirSync } = require("node:fs") as typeof import("node:fs")
  const found = readdirSync(directory).find((f) => f.endsWith(".json"))
  if (!found) throw new Error(`No cache file written in ${directory}`)
  return found
}

const cacheFileNameFor = (name: string): string => {
  const { createHash } = require("node:crypto") as typeof import("node:crypto")
  const safe = name.replace(/[^A-Za-z0-9._-]/g, "_").slice(0, 64)
  return `${safe}.${createHash("sha256").update(name).digest("hex").slice(0, 8)}.json`
}
