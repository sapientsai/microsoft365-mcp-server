import { readdirSync, readFileSync } from "node:fs"
import { dirname, join } from "node:path"
import { fileURLToPath } from "node:url"

import { describe, expect, it } from "vitest"

// Guards the invariant behind read_document's dynamic import: the three document parsers
// (mammoth/unpdf/exceljs, reached via @sapientsai/document-extract) must stay off the delegated
// server's startup path.
//
// This is NOT bought by the package split. A single top-level `import { extractTextFromBuffer }
// from "@sapientsai/document-extract"` anywhere under src/ defeats it just as thoroughly as
// bundling extraction into core would have, and nothing else in the suite would notice — the
// server would still start, every tool would still work, and startup would just quietly get
// heavier. That is precisely the kind of regression a test has to catch, because a human review
// diffing a one-line import change will not.
//
// It also cannot be caught by a preset: MS365_PRESETS is unset in the default deployment, so
// filterTools admits every tool and the `rag` domain gates nothing. Only the dynamic import does.

const HERE = dirname(fileURLToPath(import.meta.url))
const SRC = join(HERE, "..", "src")

const EXTRACT_PACKAGE = "@sapientsai/document-extract"
const PARSERS = ["mammoth", "unpdf", "exceljs"]

const sourceFiles = (dir: string): ReadonlyArray<string> =>
  readdirSync(dir, { withFileTypes: true }).flatMap((entry) => {
    const path = join(dir, entry.name)
    if (entry.isDirectory()) return sourceFiles(path)
    return entry.isFile() && path.endsWith(".ts") ? [path] : []
  })

// Comments are stripped so that merely *naming* a parser in an explanatory comment — which several
// files legitimately do — is not mistaken for an import.
const stripComments = (code: string): string => code.replace(/\/\*[\s\S]*?\*\//g, "").replace(/(^|[^:])\/\/.*$/gm, "$1")

const countAll = (code: string, specifier: string): number =>
  code.match(new RegExp(`["']${specifier.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}["']`, "g"))?.length ?? 0

const countDynamic = (code: string, specifier: string): number =>
  code.match(new RegExp(`\\bimport\\s*\\(\\s*["']${specifier.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}["']`, "g"))
    ?.length ?? 0

/**
 * Every reference to `specifier` that is not a dynamic `import(...)`. Counting rather than
 * pattern-matching import syntax keeps this robust to multi-line import statements, `export ...
 * from`, and `require(...)` alike.
 */
const staticReferences = (files: ReadonlyArray<string>, specifier: string): ReadonlyArray<string> =>
  files.filter((file) => {
    const code = stripComments(readFileSync(file, "utf-8"))
    return countAll(code, specifier) > countDynamic(code, specifier)
  })

describe("extraction stays off the startup path", () => {
  const files = sourceFiles(SRC)

  it("finds source files to check (guards against a silently empty scan)", () => {
    expect(files.length).toBeGreaterThan(20)
  })

  it("never imports the extract package statically", () => {
    expect(staticReferences(files, EXTRACT_PACKAGE)).toEqual([])
  })

  it.each(PARSERS)("never imports %s directly", (parser) => {
    expect(staticReferences(files, parser)).toEqual([])
  })

  it("reaches extraction through exactly one dynamic import", () => {
    const dynamic = files.filter((file) => countDynamic(readFileSync(file, "utf-8"), EXTRACT_PACKAGE) > 0)

    expect(dynamic).toHaveLength(1)
    expect(dynamic[0]).toMatch(/read-document-tools\.ts$/)
  })

  // The extract package is a devDependency on purpose, like core. tsdown inlines devDependencies, and
  // for a dynamic import it emits a separate chunk loaded through `import("./<chunk>.js")`, so the
  // deferral survives inlining. It used to be a runtime dependency on the belief that inlining
  // would turn the dynamic import into a no-op. That left `"@sapientsai/document-extract":
  // "workspace:*"` in the published manifest, and since the package is private, `npm install
  // microsoft365-mcp-server` failed with EUNSUPPORTEDPROTOCOL for every published version.
  //
  // The parsers are the opposite: runtime dependencies, so tsdown keeps them external and npm
  // installs them. As devDependencies they would be inlined into the chunk.
  it("bundles the extract package and leaves the parsers to npm", () => {
    const pkg = JSON.parse(readFileSync(join(HERE, "..", "package.json"), "utf-8")) as {
      dependencies?: Record<string, string>
      devDependencies?: Record<string, string>
    }

    expect(pkg.devDependencies?.[EXTRACT_PACKAGE]).toBeDefined()
    expect(pkg.dependencies?.[EXTRACT_PACKAGE]).toBeUndefined()
    for (const parser of PARSERS) expect(pkg.dependencies?.[parser]).toBeDefined()
  })

  // A workspace: specifier in dependencies is what npm cannot install. Any of them, not just extract.
  it("declares no workspace package as a runtime dependency", () => {
    const pkg = JSON.parse(readFileSync(join(HERE, "..", "package.json"), "utf-8")) as {
      dependencies?: Record<string, string>
    }

    const workspace = Object.entries(pkg.dependencies ?? {}).filter(([, range]) => range.startsWith("workspace:"))
    expect(workspace).toEqual([])
  })
})

// The bundle half of this guard — asserting the emitted dist/ reaches extraction only through a
// dynamic `import(...)` of a local chunk, and did not inline the parsers — lives in
// scripts/check-bundle-deferral.mjs and runs as part of `pnpm build`. It cannot live here: `ts-builds validate` runs test before
// build, so a vitest assertion about dist/ either skips on a clean checkout or reads a stale
// artifact locally. Neither state tests anything.
