#!/usr/bin/env node
import { main } from "./index"

main().catch((error: unknown) => {
  console.error("[Fatal]", error)
  process.exit(1)
})
