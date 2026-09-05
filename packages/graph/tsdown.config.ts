import { readFileSync } from "node:fs"
import { dirname, join } from "node:path"
import { fileURLToPath } from "node:url"

import { defineConfig } from "tsdown"

const __dirname = dirname(fileURLToPath(import.meta.url))

// This package is private and pinned at 0.0.0, so its own version says nothing. Both images in
// this repo are built and tagged from one release tag, and that tag's version lives in
// packages/microsoft365. Reading it here is the coupling that already exists in the release
// process, written down rather than left implicit — without it a tagged build reports
// "v0.0.0-dev" and cannot say which release it is.
const releasePkg = JSON.parse(readFileSync(join(__dirname, "..", "microsoft365", "package.json"), "utf-8")) as {
  version: string
}

const isProduction = process.env.NODE_ENV === "production"

export default defineConfig({
  entry: {
    index: "src/index.ts",
    bin: "src/bin.ts",
  },
  format: ["esm"],
  dts: true,
  sourcemap: isProduction,
  clean: true,
  target: "node16",
  outDir: "dist",
  platform: "node",
  treeshake: true,
  define: {
    __VERSION__: JSON.stringify(releasePkg.version),
  },
  outExtensions: () => ({
    js: ".js",
    dts: ".d.ts",
  }),
})
