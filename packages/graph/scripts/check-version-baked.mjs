#!/usr/bin/env node
// Post-build guard: the built bundle must carry a real release version.
//
// src/index.ts reads a build-time `__VERSION__` and falls back to "0.0.0-dev" when it is not
// substituted. That fallback is correct for `pnpm dev`, and silent everywhere else — for months
// every deployed container logged "microsoft-mcp-server v0.0.0-dev" because tsdown.config.ts had
// no `define` block at all. It cost real time during an incident: the logs could name the server
// but not which build was running, so identifying it meant a round trip through the registry to
// compare image digests.
//
// This runs after `build` in the validate chain rather than as a test, because tests run BEFORE
// the build that bakes the version in — a test could only ever assert the fallback.

import { readFileSync } from "node:fs"
import { dirname, join } from "node:path"
import { fileURLToPath } from "node:url"

const here = dirname(fileURLToPath(import.meta.url))
const dist = join(here, "..", "dist")
const releasePkgPath = join(here, "..", "..", "microsoft365", "package.json")

const expected = JSON.parse(readFileSync(releasePkgPath, "utf-8")).version

// The entry re-exports from a hashed chunk, so search every emitted .js rather than guessing.
const { readdirSync } = await import("node:fs")
const bundles = readdirSync(dist).filter((f) => f.endsWith(".js"))
const sources = bundles.map((f) => readFileSync(join(dist, f), "utf-8"))

const carriesVersion = sources.some((code) => code.includes(`"${expected}"`))
const carriesFallback = sources.some((code) => code.includes("0.0.0-dev"))

const fail = (message) => {
  console.error(`✖ ${message}`)
  process.exit(1)
}

if (!carriesVersion) {
  fail(
    `No bundle in dist/ contains the release version "${expected}".\n` +
      `  tsdown.config.ts must define __VERSION__ from packages/microsoft365/package.json — that is\n` +
      `  where this repo's release tag takes its version, and both images are built from one tag.\n` +
      `  Without it every container reports v0.0.0-dev and cannot say which release it is running.`,
  )
}

// The literal survives in the ternary's else-branch even when substitution works, so its presence
// alone is not a failure. Only its presence WITHOUT the real version means substitution was lost.
if (carriesFallback && !carriesVersion) {
  fail("dist/ carries the 0.0.0-dev fallback and no real version — __VERSION__ was not substituted.")
}

console.log(`✔ version check passed (dist/ reports ${expected})`)
