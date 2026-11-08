import { execFileSync } from "node:child_process"
import { readFileSync } from "node:fs"
import { join } from "node:path"

import { describe, expect, it } from "vitest"

// This repository is public and general-purpose, so it names no client or client's customer: not in
// code, tests, fixtures or docs. Fixtures use contoso.com and example names instead. The names are
// built from parts so this file doesn't match itself.
const CLIENT_NAME = new RegExp(
  [
    ["ci", "vala"],
    ["on", "cala"],
  ]
    .map((parts) => parts.join(""))
    .join("|"),
  "i",
)
const REPO_ROOT = join(__dirname, "..", "..", "..")
const SKIPPED = new Set(["pnpm-lock.yaml"])

const trackedFiles = (): ReadonlyArray<string> =>
  execFileSync("git", ["ls-files", "-z"], { cwd: REPO_ROOT, encoding: "utf8" })
    .split("\0")
    .filter((path) => path !== "" && !SKIPPED.has(path))

describe("no client names in the repository", () => {
  it("finds none in any tracked file", () => {
    const hits = trackedFiles().flatMap((path) => {
      const text = readFileSync(join(REPO_ROOT, path), "utf8")
      return text
        .split("\n")
        .flatMap((line, k) => (CLIENT_NAME.test(line) ? [`${path}:${k + 1}`] : []))
        .concat(CLIENT_NAME.test(path) ? [path] : [])
    })
    expect(hits).toEqual([])
  })
})
