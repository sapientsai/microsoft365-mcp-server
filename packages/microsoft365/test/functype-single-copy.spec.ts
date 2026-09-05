import { readFileSync } from "node:fs"
import { dirname, join } from "node:path"
import { fileURLToPath } from "node:url"

import { describe, expect, it } from "vitest"

// The workspace carried `functype: "1.4.4"` as a root override for months. Its stated reason was
// real: two copies of functype's structurally-recursive Either in one tree make tsc blow the
// instantiation budget (TS2589). Pinning one version fixed that.
//
// The pin then went stale in silence. somamcp moved to ^1.9.0 and three of our four packages moved
// to ^1.9.0, and 1.4.4 does not satisfy ^1.9.0 — so the override stopped preventing a duplicate and
// started forcing a five-minor downgrade instead. Nothing caught it: an override masks
// `pnpm outdated`, dependabot only watches direct dependencies, and the suite passes either way.
//
// So this guards the HAZARD rather than the fix. A pin is one way to keep a single copy and not the
// only one; what must stay true is that there is exactly one. If a future dependency drags in a
// second functype, this fails with the versions named, instead of tsc failing later with an
// instantiation-depth error that says nothing about why.

const HERE = dirname(fileURLToPath(import.meta.url))
const LOCKFILE = join(HERE, "..", "..", "..", "pnpm-lock.yaml")

/**
 * Versions of `functype` the lockfile resolves, as package keys rather than every reference.
 *
 * Anchored to two-space indentation so it matches entries in the `packages:`/`snapshots:` maps and
 * not the `functype: ^1.9.0` specifier lines under an importer. The negative lookbehind keeps
 * `eslint-plugin-functype` and `functype-log` out — only the bare package counts.
 */
export const resolvedFunctypeVersions = (lockfile: string): ReadonlyArray<string> => {
  const matches = lockfile.matchAll(/^ {2}functype@(\d+\.\d+\.\d+)/gm)
  return [...new Set([...matches].map((m) => m[1] as string))].sort()
}

describe("resolvedFunctypeVersions", () => {
  it("reads bare functype and ignores same-prefixed packages", () => {
    const sample = [
      "packages:",
      "  functype@1.9.0:",
      "  functype-log@1.9.0:",
      "  eslint-plugin-functype@2.109.0:",
      "  functype@1.4.4:",
    ].join("\n")
    expect(resolvedFunctypeVersions(sample)).toEqual(["1.4.4", "1.9.0"])
  })
})

describe("the workspace resolves exactly one functype", () => {
  it("has no duplicate copy to blow the tsc instantiation budget", () => {
    const versions = resolvedFunctypeVersions(readFileSync(LOCKFILE, "utf-8"))

    expect(versions.length, "no functype found in the lockfile at all — has it been dropped?").toBeGreaterThan(0)
    expect(
      versions,
      `The workspace resolves ${versions.length} copies of functype (${versions.join(", ")}). Two copies of ` +
        `its recursive Either type make tsc fail with TS2589, and the error will point at your code rather ` +
        `than at this. Find who wants the second one with \`pnpm why functype -r\`, then either align the ` +
        `ranges or add a root override in pnpm-workspace.yaml — and if you add an override, say in a comment ` +
        `which version it is tracking, because an override that outlives its reason silently downgrades ` +
        `every package instead.`,
    ).toHaveLength(1)
  })
})
