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

    describe("auth mode", () => {
      it("should default to interactive mode", () => {
        delete process.env.AZURE_AUTH_MODE

        const config = createConfig()
        expect(config.authMode).toBe("interactive")
      })

      it("should accept interactive mode explicitly", () => {
        process.env.AZURE_AUTH_MODE = "interactive"

        const config = createConfig()
        expect(config.authMode).toBe("interactive")
      })

      it("should accept clientCredentials mode", () => {
        process.env.AZURE_AUTH_MODE = "clientCredentials"
        process.env.AZURE_TENANT_ID = "specific-tenant-id"
        process.env.AZURE_CLIENT_SECRET = "test-secret"

        const config = createConfig()
        expect(config.authMode).toBe("clientCredentials")
      })

      it("should reject invalid auth mode", () => {
        process.env.AZURE_AUTH_MODE = "invalid"

        expect(() => createConfig()).toThrow('Invalid AZURE_AUTH_MODE: "invalid"')
      })

      it("should require specific tenant for clientCredentials", () => {
        process.env.AZURE_AUTH_MODE = "clientCredentials"
        process.env.AZURE_TENANT_ID = "common"
        process.env.AZURE_CLIENT_SECRET = "test-secret"

        expect(() => createConfig()).toThrow('Client credentials auth requires a specific tenant ID, not "common"')
      })

      it("should require client secret for clientCredentials", () => {
        process.env.AZURE_AUTH_MODE = "clientCredentials"
        process.env.AZURE_TENANT_ID = "specific-tenant-id"
        delete process.env.AZURE_CLIENT_SECRET

        expect(() => createConfig()).toThrow("Client credentials auth requires AZURE_CLIENT_SECRET")
      })
    })

    describe("app scopes", () => {
      it("should use default app scopes", () => {
        delete process.env.GRAPH_APP_SCOPES

        const config = createConfig()
        expect(config.appScopes).toEqual(["https://graph.microsoft.com/.default"])
      })

      it("should parse custom app scopes", () => {
        process.env.GRAPH_APP_SCOPES = "https://graph.microsoft.com/.default,https://management.azure.com/.default"

        const config = createConfig()
        expect(config.appScopes).toEqual([
          "https://graph.microsoft.com/.default",
          "https://management.azure.com/.default",
        ])
      })
    })

    describe("API key", () => {
      it("should be undefined when not set", () => {
        delete process.env.MCP_API_KEY

        const config = createConfig()
        expect(config.apiKey).toBeUndefined()
      })

      it("should parse API key when provided", () => {
        process.env.MCP_API_KEY = "my-secret-api-key"

        const config = createConfig()
        expect(config.apiKey).toBe("my-secret-api-key")
      })
    })
  })

  describe("createServer", () => {
    it("should create server with valid config", () => {
      const config = createConfig()
      const { server } = createServer(config)
      expect(server).toBeDefined()
    })

    it("should create auth provider for interactive mode", () => {
      const config = createConfig()
      const { authProvider, tokenManager } = createServer(config)

      expect(authProvider).toBeDefined()
      expect(tokenManager).toBeUndefined()
    })

    it("should create token manager for clientCredentials mode", () => {
      process.env.AZURE_AUTH_MODE = "clientCredentials"
      process.env.AZURE_TENANT_ID = "specific-tenant-id"
      process.env.AZURE_CLIENT_SECRET = "test-secret"

      const config = createConfig()
      const { authProvider, tokenManager } = createServer(config)

      expect(authProvider).toBeUndefined()
      expect(tokenManager).toBeDefined()
    })
  })
})
