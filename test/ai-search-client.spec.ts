import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

import { aiSearchFetch, parseAiSearchError } from "../src/search/ai-search-client.js"

describe("AI Search Client", () => {
  beforeEach(() => {
    vi.stubGlobal("fetch", vi.fn())
  })

  afterEach(() => {
    vi.unstubAllGlobals()
  })

  describe("parseAiSearchError", () => {
    it("should extract error message and code", async () => {
      const response = {
        status: 403,
        statusText: "Forbidden",
        json: () => Promise.resolve({ error: { code: "AuthorizationFailure", message: "Invalid API key" } }),
      } as Response

      const result = await parseAiSearchError(response)
      expect(result).toBe("AuthorizationFailure: Invalid API key")
    })

    it("should return message without code when code is absent", async () => {
      const response = {
        status: 400,
        statusText: "Bad Request",
        json: () => Promise.resolve({ error: { message: "Invalid request" } }),
      } as Response

      const result = await parseAiSearchError(response)
      expect(result).toBe("Invalid request")
    })

    it("should fall back to HTTP status when JSON has no error field", async () => {
      const response = {
        status: 500,
        statusText: "Internal Server Error",
        json: () => Promise.resolve({ unexpected: "data" }),
      } as Response

      const result = await parseAiSearchError(response)
      expect(result).toBe("HTTP 500: Internal Server Error")
    })

    it("should fall back to HTTP status when response is not JSON", async () => {
      const response = {
        status: 502,
        statusText: "Bad Gateway",
        json: () => Promise.reject(new Error("not JSON")),
      } as Response

      const result = await parseAiSearchError(response)
      expect(result).toBe("HTTP 502: Bad Gateway")
    })
  })

  describe("aiSearchFetch", () => {
    const mockApiKey = "test-api-key"
    const url = "https://my-search.search.windows.net/indexes/my-index/docs/search?api-version=2025-09-01"

    it("should return Right with Response on success", async () => {
      const mockResponse = { ok: true, status: 200, json: () => Promise.resolve({ value: [] }) }
      vi.mocked(fetch).mockResolvedValueOnce(mockResponse as Response)

      const result = await aiSearchFetch(url, mockApiKey)

      expect(result.isRight()).toBe(true)
      expect(result.orThrow()).toBe(mockResponse)
    })

    it("should set api-key header", async () => {
      vi.mocked(fetch).mockResolvedValueOnce({ ok: true } as Response)

      await aiSearchFetch(url, mockApiKey)

      expect(fetch).toHaveBeenCalledWith(url, {
        headers: { "api-key": "test-api-key" },
      })
    })

    it("should merge custom headers with api-key", async () => {
      vi.mocked(fetch).mockResolvedValueOnce({ ok: true } as Response)

      await aiSearchFetch(url, mockApiKey, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
      })

      expect(fetch).toHaveBeenCalledWith(url, {
        method: "POST",
        headers: {
          "api-key": "test-api-key",
          "Content-Type": "application/json",
        },
      })
    })

    it("should return Left with GraphError on HTTP failure", async () => {
      vi.mocked(fetch).mockResolvedValueOnce({
        ok: false,
        status: 401,
        statusText: "Unauthorized",
        json: () => Promise.resolve({ error: { code: "AuthorizationFailure", message: "The API key is invalid." } }),
      } as Response)

      const result = await aiSearchFetch(url, mockApiKey)

      expect(result.isLeft()).toBe(true)
      result.tapLeft((err) => {
        expect(err._kind).toBe("GraphError")
        expect(err.code).toBe("Unauthorized")
        expect(err.message).toBe("AuthorizationFailure: The API key is invalid.")
        expect(err.status).toBe(401)
      })
    })

    it("should return Left with fallback message when error response is not JSON", async () => {
      vi.mocked(fetch).mockResolvedValueOnce({
        ok: false,
        status: 500,
        statusText: "Internal Server Error",
        json: () => Promise.reject(new Error("not JSON")),
      } as Response)

      const result = await aiSearchFetch(url, mockApiKey)

      expect(result.isLeft()).toBe(true)
      result.tapLeft((err) => {
        expect(err.message).toBe("HTTP 500: Internal Server Error")
        expect(err.status).toBe(500)
      })
    })
  })
})
