import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

import { createConfig, createServer, DEFAULT_CLIENT_ID } from "../src/index.js"

describe("Server Configuration", () => {
  const originalEnv = process.env

  beforeEach(() => {
    vi.resetModules()
    process.env = { ...originalEnv }
  })

  afterEach(() => {
    process.env = originalEnv
  })

  describe("createConfig", () => {
    it("should use default client ID when not provided", () => {
      delete process.env.AZURE_CLIENT_ID

      const config = createConfig()
      expect(config.clientId).toBe(DEFAULT_CLIENT_ID)
    })

    it("should use provided client ID", () => {
      process.env.AZURE_CLIENT_ID = "test-client-id"

      const config = createConfig()
      expect(config.clientId).toBe("test-client-id")
    })

    it("should use default tenant ID", () => {
      delete process.env.AZURE_TENANT_ID

      const config = createConfig()
      expect(config.tenantId).toBe("common")
    })

    it("should use default port", () => {
      delete process.env.PORT

      const config = createConfig()
      expect(config.port).toBe(8080)
    })

    it("should parse custom port", () => {
      process.env.PORT = "3000"

      const config = createConfig()
      expect(config.port).toBe(3000)
    })

    it("should parse custom scopes", () => {
      process.env.GRAPH_SCOPES = "User.Read,Mail.Send,Calendars.ReadWrite"

      const config = createConfig()
      expect(config.scopes).toEqual(["User.Read", "Mail.Send", "Calendars.ReadWrite"])
    })

    it("should use default base URL", () => {
      delete process.env.BASE_URL

      const config = createConfig()
      expect(config.baseUrl).toBe("http://localhost:8080")
    })
  })

  describe("createServer", () => {
    it("should create server with valid config", () => {
      const config = createConfig()
      const { server } = createServer(config)
      expect(server).toBeDefined()
    })
  })
})
