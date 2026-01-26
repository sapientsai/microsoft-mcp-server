import { describe, expect, it, vi, beforeEach, afterEach } from "vitest"
import { createServerConfig, createServer } from "../src/index.js"

describe("Server Configuration", () => {
  const originalEnv = process.env

  beforeEach(() => {
    vi.resetModules()
    process.env = { ...originalEnv }
  })

  afterEach(() => {
    process.env = originalEnv
  })

  describe("createServerConfig", () => {
    it("should throw if no client ID or access token", () => {
      delete process.env.AZURE_CLIENT_ID
      delete process.env.ACCESS_TOKEN

      expect(() => createServerConfig()).toThrow("AZURE_CLIENT_ID or ACCESS_TOKEN environment variable is required")
    })

    it("should default to device_code mode", () => {
      process.env.AZURE_CLIENT_ID = "test-client-id"
      delete process.env.AUTH_MODE
      delete process.env.ACCESS_TOKEN

      const config = createServerConfig()
      expect(config.authMode).toBe("device_code")
    })

    it("should use client_token mode when ACCESS_TOKEN is set", () => {
      process.env.ACCESS_TOKEN = "test-token"
      delete process.env.AZURE_CLIENT_ID

      const config = createServerConfig()
      expect(config.authMode).toBe("client_token")
    })

    it("should throw if client_credentials mode without secret", () => {
      process.env.AZURE_CLIENT_ID = "test-client-id"
      process.env.AUTH_MODE = "client_credentials"
      delete process.env.AZURE_CLIENT_SECRET

      expect(() => createServerConfig()).toThrow("AZURE_CLIENT_SECRET is required for client_credentials mode")
    })

    it("should parse custom scopes", () => {
      process.env.AZURE_CLIENT_ID = "test-client-id"
      process.env.GRAPH_SCOPES = "User.Read,Mail.Send,Calendars.ReadWrite"

      const config = createServerConfig()
      expect(config.scopes).toEqual(["User.Read", "Mail.Send", "Calendars.ReadWrite"])
    })

    it("should use default tenant ID", () => {
      process.env.AZURE_CLIENT_ID = "test-client-id"
      delete process.env.AZURE_TENANT_ID

      const config = createServerConfig()
      expect(config.tenantId).toBe("common")
    })
  })

  describe("createServer", () => {
    it("should create server with valid config", () => {
      const config = {
        clientId: "test-client-id",
        tenantId: "test-tenant-id",
        authMode: "client_token" as const,
        graphApiVersion: "v1.0" as const,
        scopes: ["User.Read"],
      }

      const server = createServer(config)
      expect(server).toBeDefined()
    })
  })
})
