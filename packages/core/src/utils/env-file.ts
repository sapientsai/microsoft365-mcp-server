import { readFileSync } from "node:fs"
import { parseEnv } from "node:util"

// Loads ./.env into process.env, for local runs. Docker and an npm install under Claude Desktop
// pass configuration as real environment variables and have no .env, which is the normal case,
// so a missing file is ignored. Any other failure to read it is reported.
//
// The file is read here rather than through process.loadEnvFile, which reports every failure to
// open the file as ENOENT and keeps a UTF-8 byte-order mark as part of the first key. As with
// dotenv, a variable already set in the environment, even to "", wins over the file.
export const loadEnvFileIfPresent = (path: string = ".env"): void => {
  const content = readEnvFile(path)
  if (content === undefined) return
  Object.entries(parseEnv(content.replace(/^\uFEFF/, "")))
    .filter(([key]) => process.env[key] === undefined)
    .forEach(([key, value]) => {
      process.env[key] = value
    })
}

const readEnvFile = (path: string): string | undefined => {
  try {
    return readFileSync(path, "utf8")
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT") return undefined
    throw error
  }
}
