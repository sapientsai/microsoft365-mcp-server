import { loadEnvFile } from "node:process"

// Loads ./.env into process.env, for local runs. Docker and an npm install under Claude Desktop
// pass configuration as real environment variables and have no .env, which is the normal case,
// so a missing file is ignored. Variables already set in the environment win over the file.
export const loadEnvFileIfPresent = (path: string = ".env"): void => {
  try {
    loadEnvFile(path)
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error
  }
}
