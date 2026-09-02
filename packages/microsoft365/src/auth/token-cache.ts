import { createHash } from "node:crypto"
import { chmodSync, mkdirSync, readFileSync, renameSync, statSync, unlinkSync, writeFileSync } from "node:fs"
import { homedir } from "node:os"
import { dirname, join } from "node:path"

// Azure Identity keeps tokens in memory unless a persistence plugin is registered, so
// every restart forces a fresh interactive sign-in. That is merely annoying for a CLI
// and fatal for anything running unattended.
//
// The official plugin (@azure/identity-cache-persistence) solves this by way of keytar,
// an archived native keychain binding that has to compile at install time. It does not
// build everywhere — and under a pnpm build-script policy it is skipped outright, which
// silently leaves persistence disabled. So this implements the same plugin contract
// against a mode-0600 file instead: no native dependency, nothing to compile, and it
// behaves identically on every platform.
//
// The token cache is a bearer credential. Anyone who can read the file can act as the
// signed-in user until the refresh token expires, so the file permissions are the whole
// of the security model here and are enforced on both write and read.

const OWNER_ONLY = 0o600
const DIR_OWNER_ONLY = 0o700

type CacheContext = {
  readonly cacheHasChanged: boolean
  readonly tokenCache: {
    readonly deserialize: (data: string) => void
    readonly serialize: () => string
  }
}

export type CachePlugin = {
  readonly beforeCacheAccess: (context: CacheContext) => Promise<void>
  readonly afterCacheAccess: (context: CacheContext) => Promise<void>
}

/**
 * Where the cache lives. MS365_TOKEN_CACHE_PATH wins; otherwise it sits under the XDG
 * config directory, which is where a user would look for it.
 *
 * TOKEN_STORAGE_PATH is deliberately not consulted: it is the oauth-proxy provider's
 * own setting, and quietly sharing it would make two unrelated features move together.
 */
export const resolveCacheDirectory = (env: NodeJS.ProcessEnv = process.env): string =>
  env.MS365_TOKEN_CACHE_PATH ??
  join(env.XDG_CONFIG_HOME ?? join(homedir(), ".config"), "microsoft365-mcp-server", "token-cache")

// The cache name arrives from Azure Identity and reaches the filesystem, so it is
// reduced to a safe basename rather than trusted. A name that sanitises to nothing
// still has to produce a distinct file, hence the hash suffix.
const cacheFileName = (name: string): string => {
  const safe = name.replace(/[^A-Za-z0-9._-]/g, "_").slice(0, 64)
  const digest = createHash("sha256").update(name).digest("hex").slice(0, 8)
  return `${safe}.${digest}.json`
}

const isMissing = (error: unknown): boolean =>
  typeof error === "object" && error !== null && (error as { code?: string }).code === "ENOENT"

/**
 * Reads the cache, refusing any file that is readable by more than its owner.
 *
 * Loosened permissions mean the token may already have been exposed, so the file is
 * discarded rather than used: re-authenticating costs one sign-in, whereas trusting a
 * possibly-leaked token costs an account. Returns undefined for anything unreadable,
 * which Azure Identity treats as an empty cache.
 */
const readCache = (path: string): string | undefined => {
  try {
    const mode = statSync(path).mode & 0o777
    if (mode !== OWNER_ONLY) {
      console.error(
        `[Auth] Ignoring token cache at ${path}: permissions are ${mode.toString(8)}, expected 600. ` +
          "The file may have been exposed, so it will be discarded and a fresh sign-in requested.",
      )
      try {
        unlinkSync(path)
      } catch {
        // Best effort. If it cannot be removed, not trusting it is what matters.
      }
      return undefined
    }
    return readFileSync(path, "utf8")
  } catch (error) {
    if (!isMissing(error)) {
      console.error(`[Auth] Could not read token cache at ${path}: ${(error as Error).message}`)
    }
    return undefined
  }
}

/**
 * Writes the cache via a temporary file and a rename.
 *
 * The rename is atomic, so a crash or a concurrent reader never observes a
 * half-written cache — an outcome that would look like corruption and force a
 * re-authentication. The temporary file is created 0600 up front rather than being
 * relaxed afterwards, so the token is never briefly world-readable.
 */
const writeCache = (path: string, contents: string): void => {
  try {
    mkdirSync(dirname(path), { recursive: true, mode: DIR_OWNER_ONLY })
    const temporary = `${path}.${process.pid}.tmp`
    writeFileSync(temporary, contents, { encoding: "utf8", mode: OWNER_ONLY })
    // writeFileSync honours mode only when it creates the file; a leftover temp from a
    // previous crash would keep its old permissions, so set them explicitly.
    chmodSync(temporary, OWNER_ONLY)
    renameSync(temporary, path)
  } catch (error) {
    // A cache that cannot be written is a lost convenience, not a failed sign-in, so
    // this degrades to in-memory rather than propagating.
    console.error(`[Auth] Could not persist token cache to ${path}: ${(error as Error).message}`)
  }
}

/**
 * Builds the plugin Azure Identity asks for when tokenCachePersistenceOptions is set.
 * Matches the shape of the official keytar-backed plugin, minus the native dependency.
 */
export const createFileCachePlugin = (name: string, directory = resolveCacheDirectory()): CachePlugin => {
  const path = join(directory, cacheFileName(name))

  return {
    beforeCacheAccess: (context) => {
      const contents = readCache(path)
      if (contents) context.tokenCache.deserialize(contents)
      return Promise.resolve()
    },
    afterCacheAccess: (context) => {
      if (context.cacheHasChanged) writeCache(path, context.tokenCache.serialize())
      return Promise.resolve()
    },
  }
}

type PluginContext = {
  readonly cachePluginControl: {
    readonly setPersistence: (provider: (options?: { name?: string }) => Promise<CachePlugin>) => void
  }
}

/**
 * An identity plugin, in the shape `useIdentityPlugin` expects: a function handed a
 * plugin context. Identity types that context as `unknown` and documents that plugin
 * authors are responsible for casting it, which is what happens here.
 */
export const fileCachePersistencePlugin = (context: unknown): void => {
  const { cachePluginControl } = context as PluginContext
  cachePluginControl.setPersistence((options) => Promise.resolve(createFileCachePlugin(options?.name ?? "msal.cache")))
}

// The token cache alone is not enough to skip a sign-in.
//
// InteractiveBrowserCredential only attempts silent auth when it already knows which
// account to look for: with no cachedAccount it raises AuthenticationRequiredError
// before consulting the cache at all. That account identity comes from an
// AuthenticationRecord, which lives in memory and dies with the process — so without
// persisting it too, a perfectly good refresh token sits on disk unreachable and every
// restart prompts again.
//
// The record is not a credential: it identifies the account (home account id, tenant,
// username, authority). It is stored beside the cache at the same 0600 anyway, since
// the username is personal information.

export type AuthenticationRecordLike = {
  readonly authority: string
  readonly homeAccountId: string
  readonly tenantId: string
  readonly username: string
  readonly clientId: string
}

const RECORD_FILE = "authentication-record.json"

export const readAuthenticationRecord = (directory = resolveCacheDirectory()): AuthenticationRecordLike | undefined => {
  const contents = readCache(join(directory, RECORD_FILE))
  if (!contents) return undefined

  try {
    const parsed = JSON.parse(contents) as Partial<AuthenticationRecordLike>
    // A truncated or hand-edited record would otherwise fail deep inside MSAL with a
    // far less obvious error than simply signing in again.
    if (!parsed.homeAccountId || !parsed.username || !parsed.clientId) return undefined
    return parsed as AuthenticationRecordLike
  } catch {
    console.error("[Auth] Stored authentication record is unreadable; a fresh sign-in will be required.")
    return undefined
  }
}

export const writeAuthenticationRecord = (
  record: AuthenticationRecordLike,
  directory = resolveCacheDirectory(),
): void => writeCache(join(directory, RECORD_FILE), JSON.stringify(record))
