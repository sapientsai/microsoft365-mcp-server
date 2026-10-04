#!/usr/bin/env node
// Post-build half of the guard in test/lazy-extraction.spec.ts.
//
// That spec asserts the *sources* keep document extraction behind a dynamic import. This script
// asserts the *bundler honored it*. tsdown inlines the extract package (a devDependency) into its
// own chunk and turns `import("@sapientsai/document-extract")` into `import("./<chunk>.js")`. The
// parsers (mammoth/unpdf/exceljs) stay external, because they are runtime dependencies. So the
// invariants on dist/ are:
//
//   1. No file reachable by *static* import from an entry point references a parser. That is the
//      startup path; a parser there loads on every server start.
//   2. index.js dynamically imports a local chunk whose static graph reaches all three parsers, so
//      extraction is still reachable — lazily.
//   3. Every parser appears as an import specifier somewhere. If none does, the bundler inlined it,
//      which ships megabytes of parser code in the bundle.
//   4. No file references a workspace package. They are private and never published, so a
//      specifier left in dist/ cannot resolve after `npm install` — the bug that made every
//      published version uninstallable while it was still a runtime dependency.
//
// It lives in the build step rather than the test suite on purpose. `ts-builds validate` runs
// test before build, so a vitest assertion about dist/ can only ever read a bundle from some
// earlier build — it skips on a clean checkout (dead in CI) and asserts against a stale artifact
// locally. Running right after `tsdown --clean` is the only point where the bundle on disk is
// known to correspond to the sources just checked.
//
// Wired two ways, both of which must stay in place:
//   - `ts-builds.config.json` appends `check:bundle` to the validate chain, so CI covers it.
//     That chain restates ts-builds' default steps because there is no append mechanism; if the
//     toolchain's default chain gains a step, add it there too.
//   - the `build` script in package.json, so an explicit `pnpm build` verifies what it emitted.

import { readdirSync, readFileSync } from "node:fs"
import { dirname, join, relative } from "node:path"
import { fileURLToPath } from "node:url"

const PACKAGE_ROOT = join(dirname(fileURLToPath(import.meta.url)), "..")
const DIST = join(PACKAGE_ROOT, "dist")
const ENTRIES = ["index.js", "bin.js"]

const PARSERS = ["mammoth", "unpdf", "exceljs"]
const WORKSPACE_PACKAGES = ["@sapientsai/document-extract", "@sapientsai/ms-graph-core"]

// Comments are stripped so that a specifier named in a preserved comment or a sourcemap URL is not
// mistaken for an import. The `[^:]` guard keeps `https://` from being read as a line comment.
const stripComments = (code) => code.replace(/\/\*[\s\S]*?\*\//g, "").replace(/(^|[^:])\/\/.*$/gm, "$1")

const escape = (specifier) => specifier.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")

// A specifier or any of its subpaths: "unpdf/pdfjs" or "exceljs/dist/es5" loads the parser too.
const specifierPattern = (specifier) => `["']${escape(specifier)}(?:/[^"']*)?["']`

const countAll = (code, specifier) => code.match(new RegExp(specifierPattern(specifier), "g"))?.length ?? 0

const countDynamic = (code, specifier) =>
  code.match(new RegExp(`\\bimport\\s*\\(\\s*${specifierPattern(specifier)}`, "g"))?.length ?? 0

// Local chunk specifiers. Static: `from "./x.js"` and bare `import "./x.js"`. Dynamic: `import("./x.js")`.
const localStatic = (code) =>
  [...code.matchAll(/(?:\bfrom|\bimport)\s*["']\.\/([^"']+\.js)["']/g)].map((match) => match[1])

const localDynamic = (code) => [...code.matchAll(/\bimport\s*\(\s*["']\.\/([^"']+\.js)["']/g)].map((match) => match[1])

const failures = []

let files
try {
  files = readdirSync(DIST).filter((name) => name.endsWith(".js"))
} catch {
  failures.push("dist/ is missing — the build did not emit it, so the deferral is unverifiable.")
  files = []
}

const code = new Map(files.map((name) => [name, stripComments(readFileSync(join(DIST, name), "utf-8"))]))

/** Every file reachable from `start` through static imports, including `start` itself. */
const staticClosure = (start) => {
  const seen = new Set()
  const visit = (name) => {
    if (seen.has(name) || !code.has(name)) return
    seen.add(name)
    localStatic(code.get(name)).forEach(visit)
  }
  visit(start)
  return seen
}

for (const entry of ENTRIES) {
  if (!code.has(entry)) {
    failures.push(`dist/${entry} is missing — the build did not emit it, so the deferral is unverifiable.`)
    continue
  }
  for (const name of staticClosure(entry)) {
    for (const parser of PARSERS) {
      const statics = countAll(code.get(name), parser) - countDynamic(code.get(name), parser)
      if (statics > 0) {
        const via = name === entry ? "" : ` (statically imported from dist/${entry})`
        failures.push(`dist/${name}${via} references "${parser}" — the parser is on the startup path.`)
      }
    }
  }
}

if (code.has("index.js")) {
  const lazyGraph = new Set(localDynamic(code.get("index.js")).flatMap((chunk) => [...staticClosure(chunk)]))
  const reachesAll = PARSERS.every((parser) => [...lazyGraph].some((name) => countAll(code.get(name), parser) > 0))
  if (!reachesAll) {
    failures.push(
      "dist/index.js dynamically imports no chunk that reaches all of " +
        `${PARSERS.join("/")} — read_document's extraction is missing or no longer deferred.`,
    )
  }
}

for (const parser of PARSERS) {
  if (![...code.values()].some((source) => countAll(source, parser) > 0)) {
    failures.push(
      `No file in dist/ imports "${parser}" — either the bundler inlined it (keep it in dependencies, not ` +
        `devDependencies) or extraction is no longer bundled at all.`,
    )
  }
}

for (const [name, source] of code) {
  for (const pkg of WORKSPACE_PACKAGES) {
    if (countAll(source, pkg) > 0) {
      failures.push(
        `dist/${name} references "${pkg}", a private workspace package that npm cannot install. ` +
          `Keep it in devDependencies so the bundler inlines it.`,
      )
    }
  }
}

if (failures.length > 0) {
  console.error(`\n✗ bundle deferral check failed (${relative(process.cwd(), PACKAGE_ROOT) || "."}):\n`)
  for (const failure of failures) console.error(`  - ${failure}`)
  console.error("\nSee test/lazy-extraction.spec.ts for the source-level half of this guard.\n")
  process.exit(1)
}

console.log(`✔ bundle deferral check passed (${files.map((name) => `dist/${name}`).join(", ")})`)
