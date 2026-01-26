import { describe, expect, it, beforeEach } from "vitest"
import { AuthManager } from "../src/auth/auth-manager.js"
import type { ServerConfig } from "../src/types.js"

describe("AuthManager", () => {
  const baseConfig: ServerConfig = {
    clientId: "test-client-id",
    tenantId: "test-tenant-id",
    authMode: "client_token",
    graphApiVersion: "v1.0",
    scopes: ["User.Read"],
  }

  describe("client_token mode", () => {
    let authManager: AuthManager

    beforeEach(() => {
      authManager = new AuthManager(baseConfig)
    })

    it("should start with no token", async () => {
      const status = await authManager.getAuthStatus()
      expect(status.authenticated).toBe(false)
      expect(status.mode).toBe("client_token")
    })

    it("should set access token", async () => {
      const token = "test-access-token"
      const expiresOn = new Date(Date.now() + 3600 * 1000)

      authManager.setAccessToken(token, expiresOn)

      const accessToken = await authManager.getAccessToken()
      expect(accessToken).toBe(token)
    })

    it("should report authenticated after setting token", async () => {
      authManager.setAccessToken("test-token")

      const status = await authManager.getAuthStatus()
      expect(status.authenticated).toBe(true)
    })

    it("should clear token on sign out", async () => {
      authManager.setAccessToken("test-token")
      await authManager.signOut()

      const status = await authManager.getAuthStatus()
      expect(status.authenticated).toBe(false)
    })

    it("should return null for expired token", async () => {
      const expiredDate = new Date(Date.now() - 1000)
      authManager.setAccessToken("expired-token", expiredDate)

      const token = await authManager.getAccessToken()
      expect(token).toBeNull()
    })
  })

  describe("mode switching", () => {
    it("should allow setting mode", () => {
      const authManager = new AuthManager(baseConfig)

      authManager.setMode("device_code")
      expect(authManager.getMode()).toBe("device_code")

      authManager.setMode("client_credentials")
      expect(authManager.getMode()).toBe("client_credentials")
    })

    it("should switch to client_token mode when setting token", () => {
      const authManager = new AuthManager({
        ...baseConfig,
        authMode: "device_code",
      })

      authManager.setAccessToken("test-token")
      expect(authManager.getMode()).toBe("client_token")
    })
  })
})
