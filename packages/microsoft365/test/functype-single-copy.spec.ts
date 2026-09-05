import { readdirSync, readFileSync } from "node:fs"
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
const ROOT = join(HERE, "..", "..", "..")
const LOCKFILE = join(ROOT, "pnpm-lock.yaml")
const PACKAGES = join(ROOT, "packages")

/**
 * Versions of `functype` the lockfile resolves, as package keys rather than every reference.
 *
 * Anchored to two-space indentation so it matches entries in the `packages:`/`snapshots:` maps and
 * not the deeper-indented `functype: ^1.9.0` specifier lines under an importer. That same anchor is
 * what excludes `functype-log` and `eslint-plugin-functype`: the key has to start with `functype@`.
 *
 * The version group deliberately runs past the patch digits to the peer suffix or colon, so two
 * prereleases of one version stay distinct rather than collapsing into a single entry.
 */
export const resolvedFunctypeVersions = (lockfile: string): ReadonlyArray<string> => {
  const matches = lockfile.matchAll(/^ {2}functype@(\d+\.\d+\.\d+[^(:\s]*)/gm)
  return [...new Set([...matches].map((m) => m[1] as string))].sort()
}

describe("resolvedFunctypeVersions", () => {
  it("reads bare functype and ignores same-prefixed packages", () => {
    const sample = [
      "packages:",
      "  functype@1.9.0:",
      "  functype-log@1.9.0:",
      "  eslint-plugin-functype@2.109.0:",
      "  functype@1.4.4(zod@4.5.4):",
      "  functype@2.0.0-beta.1:",
      "  functype@2.0.0-beta.2:",
    ].join("\n")
    // Peer suffixes are stripped; prereleases of one version stay distinct.
    expect(resolvedFunctypeVersions(sample)).toEqual(["1.4.4", "1.9.0", "2.0.0-beta.1", "2.0.0-beta.2"])
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

/**
 * Whether `version` falls inside `range`, for the caret / tilde / exact forms that appear in this
 * workspace. Not a semver library on purpose — a spec asserting something about dependencies should
 * not add one.
 */
export const satisfies = (version: string, range: string): boolean => {
  const parse = (raw: string): readonly number[] => (raw.match(/(\d+)\.(\d+)\.(\d+)/) ?? []).slice(1, 4).map(Number)
  const v = parse(version)
  const floor = parse(range)
  if (v.length !== 3 || floor.length !== 3) return false

  const compare = (a: readonly number[], b: readonly number[]): number =>
    (a[0] ?? 0) - (b[0] ?? 0) || (a[1] ?? 0) - (b[1] ?? 0) || (a[2] ?? 0) - (b[2] ?? 0)

  if (compare(v, floor) < 0) return false
  if (range.startsWith("^")) return v[0] === floor[0]
  if (range.startsWith("~")) return v[0] === floor[0] && v[1] === floor[1]
  return compare(v, floor) === 0
}

describe("satisfies", () => {
  it.each([
    ["1.9.0", "^1.9.0", true],
    ["1.9.1", "^1.9.0", true],
    ["1.4.4", "^1.9.0", false], // the #69 bug: an override forcing a version below the declared range
    ["1.9.0", "^2.0.0", false],
    ["1.9.0", "~1.9.0", true],
    ["1.10.0", "~1.9.0", false],
    ["1.4.4", "1.4.4", true],
  ])("%s vs %s -> %s", (version, range, expected) => {
    expect(satisfies(version, range)).toBe(expected)
  })
})

// The duplicate check above would have passed on main: the override held everything at a single
// 1.4.4, so there was exactly one copy and no signal at all. What it could not see is that 1.4.4
// does not satisfy the ^1.9.0 those packages declare — a root override rewrites the resolution
// without touching the range, so the two drift apart silently and every other tool is masked.
// This is the assertion that actually catches issue #69.
describe("every package gets the functype it asks for", () => {
  const declared = readdirSync(PACKAGES, { withFileTypes: true })
    .filter((entry) => entry.isDirectory())
    .flatMap((entry) => {
      const manifest = JSON.parse(readFileSync(join(PACKAGES, entry.name, "package.json"), "utf-8")) as {
        dependencies?: Record<string, string>
      }
      const range = manifest.dependencies?.functype
      return range ? [{ name: entry.name, range }] : []
    })

  it("finds the workspace packages that depend on functype", () => {
    expect(declared.length).toBeGreaterThan(0)
  })

  it.each(declared)("$name declares $range and gets it", ({ name, range }) => {
    const [resolved] = resolvedFunctypeVersions(readFileSync(LOCKFILE, "utf-8"))
    expect(
      satisfies(resolved ?? "", range),
      `packages/${name} declares functype ${range} but the lockfile resolves ${resolved}. Something is ` +
        `overriding the resolution below the declared range — check overrides in pnpm-workspace.yaml. ` +
        `An override that outlives its reason downgrades every package instead of protecting them, and ` +
        `nothing else reports it: it masks pnpm outdated, and the suite passes either way.`,
    ).toBe(true)
  })
})
