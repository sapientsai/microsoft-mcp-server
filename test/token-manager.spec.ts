import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

import type { ServerConfig } from "../src/auth/types.js"
import { createTokenManager } from "../src/auth/token-manager.js"

describe("Token Manager", () => {
  const mockConfig: ServerConfig = {
    clientId: "test-client-id",
    clientSecret: "test-client-secret",
    tenantId: "test-tenant-id",
    baseUrl: "http://localhost:8080",
    port: 8080,
    scopes: ["User.Read"],
    authMode: "clientCredentials",
    appScopes: ["https://graph.microsoft.com/.default"],
  }

  beforeEach(() => {
    vi.useFakeTimers()
    vi.stubGlobal("fetch", vi.fn())
  })

  afterEach(() => {
    vi.useRealTimers()
    vi.unstubAllGlobals()
  })

  it("should fetch token on first call", async () => {
    const mockResponse = {
      access_token: "test-token-123",
      expires_in: 3600,
      token_type: "Bearer",
    }

    vi.mocked(fetch).mockResolvedValueOnce({
      ok: true,
      json: () => Promise.resolve(mockResponse),
    } as Response)

    const tokenManager = createTokenManager(mockConfig)
    const token = await tokenManager.getToken()

    expect(token).toBe("test-token-123")
    expect(fetch).toHaveBeenCalledTimes(1)
    expect(fetch).toHaveBeenCalledWith(
      `https://login.microsoftonline.com/${mockConfig.tenantId}/oauth2/v2.0/token`,
      expect.objectContaining({
        method: "POST",
        headers: {
          "Content-Type": "application/x-www-form-urlencoded",
        },
      }),
    )
  })

  it("should return cached token when valid", async () => {
    const mockResponse = {
      access_token: "cached-token",
      expires_in: 3600,
      token_type: "Bearer",
    }

    vi.mocked(fetch).mockResolvedValueOnce({
      ok: true,
      json: () => Promise.resolve(mockResponse),
    } as Response)

    const tokenManager = createTokenManager(mockConfig)

    // First call - fetches token
    const token1 = await tokenManager.getToken()
    // Second call - should return cached token
    const token2 = await tokenManager.getToken()

    expect(token1).toBe("cached-token")
    expect(token2).toBe("cached-token")
    expect(fetch).toHaveBeenCalledTimes(1) // Only called once
  })

  it("should refresh token when near expiry", async () => {
    const firstResponse = {
      access_token: "first-token",
      expires_in: 360, // 6 minutes
      token_type: "Bearer",
    }

    const secondResponse = {
      access_token: "refreshed-token",
      expires_in: 3600,
      token_type: "Bearer",
    }

    vi.mocked(fetch)
      .mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve(firstResponse),
      } as Response)
      .mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve(secondResponse),
      } as Response)

    const tokenManager = createTokenManager(mockConfig)

    // First call
    const token1 = await tokenManager.getToken()
    expect(token1).toBe("first-token")

    // Advance time by 2 minutes (within 5 minute buffer)
    vi.advanceTimersByTime(2 * 60 * 1000)

    // Should refresh because we're within 5 minute buffer of 6 minute expiry
    const token2 = await tokenManager.getToken()
    expect(token2).toBe("refreshed-token")
    expect(fetch).toHaveBeenCalledTimes(2)
  })

  it("should throw error on failed token request", async () => {
    const errorResponse = {
      error: "invalid_client",
      error_description: "Invalid client credentials",
    }

    vi.mocked(fetch).mockResolvedValueOnce({
      ok: false,
      status: 401,
      text: () => Promise.resolve(JSON.stringify(errorResponse)),
    } as Response)

    const tokenManager = createTokenManager(mockConfig)

    await expect(tokenManager.getToken()).rejects.toThrow("Invalid client credentials")
  })

  it("should handle non-JSON error response", async () => {
    vi.mocked(fetch).mockResolvedValueOnce({
      ok: false,
      status: 500,
      text: () => Promise.resolve("Internal Server Error"),
    } as Response)

    const tokenManager = createTokenManager(mockConfig)

    await expect(tokenManager.getToken()).rejects.toThrow("Token request failed: 500 - Internal Server Error")
  })

  it("should return session with expiry info", async () => {
    const now = new Date("2024-01-15T12:00:00Z")
    vi.setSystemTime(now)

    const mockResponse = {
      access_token: "session-token",
      expires_in: 3600,
      token_type: "Bearer",
    }

    vi.mocked(fetch).mockResolvedValueOnce({
      ok: true,
      json: () => Promise.resolve(mockResponse),
    } as Response)

    const tokenManager = createTokenManager(mockConfig)
    const session = await tokenManager.getSession()

    expect(session.accessToken).toBe("session-token")
    expect(session.mode).toBe("clientCredentials")
    expect(session.expiresAt).toEqual(new Date("2024-01-15T13:00:00Z"))
  })
})
