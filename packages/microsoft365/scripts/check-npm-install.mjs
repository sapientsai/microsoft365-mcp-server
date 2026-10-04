#!/usr/bin/env node
// Packs this package and installs the tarball with plain npm, the way a user's `npx` would, then
// proves the installed copy can load what it needs.
//
// Every version published from this monorepo up to 1.2.7 failed `npm install` with
// EUNSUPPORTEDPROTOCOL: the manifest carried "@sapientsai/document-extract": "workspace:*", a
// private package npm can neither rewrite nor fetch. CI was green throughout, because nothing ever
// installed the artifact outside the workspace — the Docker image is built with `pnpm deploy`,
// which resolves workspace links itself. This check is the missing install.
//
// It needs the network and a built dist/, so it runs as its own CI step after validate rather than
// inside the validate chain:
//   - node.js.yml, so a pull request that reintroduces the bug fails;
//   - publish.yml, before the GitHub release and `npm publish`, so a bad artifact is never released.
//
// Set KEEP_INSTALL_CHECK=1 to keep the scratch directory for inspection.

import { execFileSync, spawnSync } from "node:child_process"
import { mkdirSync, mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import { dirname, join } from "node:path"
import { fileURLToPath } from "node:url"

const PACKAGE_ROOT = join(dirname(fileURLToPath(import.meta.url)), "..")
const pkg = JSON.parse(readFileSync(join(PACKAGE_ROOT, "package.json"), "utf-8"))

// Protocols only a workspace-aware package manager understands. npm fails or misresolves each one.
const LOCAL_PROTOCOL = /^(workspace|link|file|catalog|portal):/
const DEPENDENCY_FIELDS = ["dependencies", "optionalDependencies", "peerDependencies"]

const work = mkdtempSync(join(tmpdir(), "m365-install-check-"))

const fail = (message) => {
  console.error(`\n✗ npm install check failed:\n  ${message.replaceAll("\n", "\n  ")}\n`)
  console.error(`  Scratch directory kept for inspection: ${work}\n`)
  process.exit(1)
}

// Five minutes per step: a hung `npm install` would otherwise hold the CI job to its six-hour cap.
const STEP_TIMEOUT_MS = 300_000

const run = (command, args, cwd) => {
  const result = spawnSync(command, args, { cwd, encoding: "utf-8", timeout: STEP_TIMEOUT_MS })
  if (result.error) fail(`\`${command} ${args.join(" ")}\` did not finish: ${result.error.message}`)
  if (result.status !== 0) {
    fail(`\`${command} ${args.join(" ")}\` exited ${result.status}:\n${(result.stderr || result.stdout).trim()}`)
  }
  return result
}

// 1. Pack exactly what `npm publish` would upload.
const packed = JSON.parse(run("npm", ["pack", "--json", "--pack-destination", work], PACKAGE_ROOT).stdout)
const tarball = join(work, packed[0].filename)

// 2. The published manifest must name only registry-resolvable dependencies.
execFileSync("tar", ["-xzf", tarball, "-C", work, "package/package.json"])
const manifest = JSON.parse(readFileSync(join(work, "package", "package.json"), "utf-8"))
const local = DEPENDENCY_FIELDS.flatMap((field) =>
  Object.entries(manifest[field] ?? {})
    .filter(([, range]) => LOCAL_PROTOCOL.test(range))
    .map(([name, range]) => `${field}: "${name}": "${range}"`),
)
if (local.length > 0) {
  fail(
    `The packed package.json names dependencies npm cannot install:\n${local.join("\n")}\n` +
      `Bundle workspace packages instead: list them in devDependencies so tsdown inlines them.`,
  )
}

// 3. Install the tarball into an empty project. Lifecycle scripts are skipped: production never
// runs them either (the Docker image is assembled by pnpm with builds disallowed).
const app = join(work, "app")
mkdirSync(app)
writeFileSync(join(app, "package.json"), JSON.stringify({ name: "install-check", private: true, type: "module" }))
run("npm", ["install", tarball, "--ignore-scripts", "--no-audit", "--no-fund", "--loglevel=error"], app)

// 4. From inside the installed package: every bare specifier in dist/ must resolve, and the lazily
// loaded extraction chunk must import (which loads mammoth/unpdf/exceljs) and extract text.
// index.js itself cannot be imported here — it starts the server at module scope.
const installedDist = join(app, "node_modules", pkg.name, "dist")
const chunks = readdirSync(installedDist).filter((name) => name.endsWith(".js"))
const specifiers = new Set(
  chunks.flatMap((name) =>
    [
      ...readFileSync(join(installedDist, name), "utf-8").matchAll(
        /(?:\bfrom\s*|\bimport\s*\(?\s*)["']((?:@[\w.-]+\/)?\w[\w.-]*(?:\/[\w./-]+)?|node:[\w/]+)["']/g,
      ),
    ]
      .map((match) => match[1])
      .filter((specifier) => !specifier.startsWith("node:")),
  ),
)
const lazyChunks = chunks.filter((name) => name !== "index.js" && name !== "bin.js")

const probe = join(installedDist, "__install_check__.mjs")
writeFileSync(
  probe,
  `const unresolved = []
for (const specifier of ${JSON.stringify([...specifiers])}) {
  try { import.meta.resolve(specifier) } catch (error) { unresolved.push(specifier + ": " + error.message) }
}
if (unresolved.length > 0) throw new Error("unresolvable imports:\\n" + unresolved.join("\\n"))

let extracted = 0
for (const chunk of ${JSON.stringify(lazyChunks)}) {
  const mod = await import("./" + chunk)
  if (typeof mod.extractTextFromBuffer !== "function") continue
  const result = await mod.extractTextFromBuffer(Buffer.from("hello from npm"), "text/plain", "check.txt")
  if (!result.isRight() || result.value !== "hello from npm") throw new Error("extraction returned " + String(result.value))
  extracted++
}
if (extracted === 0) throw new Error("no chunk exports extractTextFromBuffer — read_document has nothing to load")
console.log(\`resolved \${${specifiers.size}} specifiers, extracted text through \${extracted} chunk(s)\`)
`,
)
const probed = run(process.execPath, [probe], installedDist)

// 5. The bin entry runs. It prints through console.error (stdout is reserved for JSON-RPC).
const versionRun = run(join(app, "node_modules", ".bin", pkg.name), ["--version"], app)
const reported = `${versionRun.stdout}${versionRun.stderr}`.trim()
if (reported !== pkg.version) fail(`\`${pkg.name} --version\` printed "${reported}", expected "${pkg.version}".`)

if (!process.env.KEEP_INSTALL_CHECK) rmSync(work, { recursive: true, force: true })
console.log(`✔ npm install check passed (${packed[0].filename}: ${probed.stdout.trim()}; --version ${reported})`)
