import { mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"

import { afterEach, describe, expect, it, vi } from "vitest"

import { loadEnvFileIfPresent } from "../src/utils/env-file"

const dir = mkdtempSync(join(tmpdir(), "env-file-"))

afterEach(() => vi.unstubAllEnvs())

describe("loadEnvFileIfPresent", () => {
  // Docker and Claude Desktop installs have no .env; that must not stop the server starting.
  it("ignores a missing file", () => {
    expect(() => loadEnvFileIfPresent(join(dir, "missing.env"))).not.toThrow()
  })

  it("loads variables from the file", () => {
    const file = join(dir, "load.env")
    writeFileSync(file, 'ENV_FILE_TEST_A=from_file\nENV_FILE_TEST_B="quoted # kept"\n')
    vi.stubEnv("ENV_FILE_TEST_A", undefined)
    vi.stubEnv("ENV_FILE_TEST_B", undefined)

    loadEnvFileIfPresent(file)

    expect(process.env.ENV_FILE_TEST_A).toBe("from_file")
    expect(process.env.ENV_FILE_TEST_B).toBe("quoted # kept")
  })

  // dotenv's behaviour, which deployments rely on: a real environment variable beats the file.
  it("keeps a variable already set in the environment", () => {
    const file = join(dir, "override.env")
    writeFileSync(file, "ENV_FILE_TEST_C=from_file\n")
    vi.stubEnv("ENV_FILE_TEST_C", "from_shell")

    loadEnvFileIfPresent(file)

    expect(process.env.ENV_FILE_TEST_C).toBe("from_shell")
  })

  it("still reports a .env path it can't read for another reason", () => {
    expect(() => loadEnvFileIfPresent(dir)).toThrow()
  })
})

process.on("exit", () => rmSync(dir, { recursive: true, force: true }))
