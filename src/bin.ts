#!/usr/bin/env node
import "dotenv/config"

import { runServer } from "./index.js"

runServer().catch((error) => {
  console.error("Failed to start server:", error)
  process.exit(1)
})
