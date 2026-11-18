#!/usr/bin/env node
// Keeps server.json and manifest.json in step with package.json.
//
// The MCP Registry rejects a submission whose server.json version does not name
// a version that exists on npm, and it verifies ownership by matching
// package.json `mcpName` against server.json `name`. Both are easy to forget on
// a release, and the failure only shows up in the publish job.
//
// `npm version` bumps package.json alone, so server.json is generated from it
// rather than hand-edited.
//
// manifest.json describes the .mcpb desktop extension, which publish.yml attaches to
// each release. Only its top-level "version" line is rewritten, so the file's own
// formatting stays as Prettier leaves it.
//
//   node scripts/sync-server-json.mjs           write server.json
//   node scripts/sync-server-json.mjs --check   fail if either is stale
//
// There is deliberately no `version` lifecycle hook: `npm version` does not see a
// repository from packages/microsoft365, which is not a git root, so the hook would
// never fire. `pnpm sync:server-json` is a documented release step instead, and
// --check runs in CI, in prepublishOnly, and again before the registry publish.

import { readFile, writeFile } from "node:fs/promises"

const TARGET = "server.json"

const pkg = JSON.parse(await readFile("package.json", "utf8"))
const server = JSON.parse(await readFile(TARGET, "utf8"))

if (server.name !== pkg.mcpName) {
  console.error(`✗ package.json mcpName "${pkg.mcpName}" does not match ${TARGET} name "${server.name}".`)
  console.error("  The registry uses that match to verify npm package ownership.")
  process.exit(1)
}

const synced = {
  ...server,
  version: pkg.version,
  packages: server.packages.map((p) => (p.registryType === "npm" ? { ...p, version: pkg.version } : p)),
}

const MANIFEST = "manifest.json"
// Two-space indent: the top-level key, not a nested "version".
const MANIFEST_VERSION = /^( {2}"version": ")[^"]*(")/m

const manifest = await readFile(MANIFEST, "utf8")
if (!MANIFEST_VERSION.test(manifest)) {
  console.error(`✗ ${MANIFEST} has no top-level "version" line to keep in step.`)
  process.exit(1)
}

const files = [
  { path: TARGET, current: await readFile(TARGET, "utf8"), next: `${JSON.stringify(synced, null, 2)}\n` },
  { path: MANIFEST, current: manifest, next: manifest.replace(MANIFEST_VERSION, `$1${pkg.version}$2`) },
]

if (process.argv.includes("--check")) {
  const stale = files.filter((f) => f.current !== f.next)
  for (const f of stale) console.error(`✗ ${f.path} is out of sync with package.json (${pkg.version}).`)
  if (stale.length > 0) {
    console.error("  Run: pnpm sync:server-json")
    process.exit(1)
  }
  console.log(`✓ ${files.map((f) => f.path).join(" and ")} match package.json (${pkg.version})`)
} else {
  for (const f of files) await writeFile(f.path, f.next)
  console.log(`✓ wrote ${files.map((f) => f.path).join(" and ")} (${pkg.version})`)
}
