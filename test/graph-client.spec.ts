import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

import { graphFetch, parseGraphError } from "../src/graph/client.js"

describe("Graph Client", () => {
  beforeEach(() => {
    vi.stubGlobal("fetch", vi.fn())
  })

  afterEach(() => {
    vi.unstubAllGlobals()
  })

  describe("parseGraphError", () => {
    it("should extract error message and code from Graph API error response", async () => {
      const response = {
        status: 403,
        statusText: "Forbidden",
        json: () => Promise.resolve({ error: { code: "AccessDenied", message: "Insufficient privileges" } }),
      } as Response

      const result = await parseGraphError(response)
      expect(result).toBe("AccessDenied: Insufficient privileges")
    })

    it("should return message without code when code is absent", async () => {
      const response = {
        status: 400,
        statusText: "Bad Request",
        json: () => Promise.resolve({ error: { message: "Invalid query parameter" } }),
      } as Response

      const result = await parseGraphError(response)
      expect(result).toBe("Invalid query parameter")
    })

    it("should fall back to HTTP status when JSON has no error field", async () => {
      const response = {
        status: 500,
        statusText: "Internal Server Error",
        json: () => Promise.resolve({ unexpected: "data" }),
      } as Response

      const result = await parseGraphError(response)
      expect(result).toBe("HTTP 500: Internal Server Error")
    })

    it("should fall back to HTTP status when response is not JSON", async () => {
      const response = {
        status: 502,
        statusText: "Bad Gateway",
        json: () => Promise.reject(new Error("not JSON")),
      } as Response

      const result = await parseGraphError(response)
      expect(result).toBe("HTTP 502: Bad Gateway")
    })

    it("should fall back when error.message is empty", async () => {
      const response = {
        status: 404,
        statusText: "Not Found",
        json: () => Promise.resolve({ error: { code: "NotFound" } }),
      } as Response

      const result = await parseGraphError(response)
      expect(result).toBe("HTTP 404: Not Found")
    })
  })

  describe("graphFetch", () => {
    const mockToken = "test-access-token"

    it("should return Right with Response on success", async () => {
      const mockResponse = { ok: true, status: 200, json: () => Promise.resolve({ value: [] }) }
      vi.mocked(fetch).mockResolvedValueOnce(mockResponse as Response)

      const result = await graphFetch("https://graph.microsoft.com/v1.0/me", mockToken)

      expect(result.isRight()).toBe(true)
      expect(result.orThrow()).toBe(mockResponse)
    })

    it("should set Authorization header with bearer token", async () => {
      vi.mocked(fetch).mockResolvedValueOnce({ ok: true } as Response)

      await graphFetch("https://graph.microsoft.com/v1.0/me", mockToken)

      expect(fetch).toHaveBeenCalledWith("https://graph.microsoft.com/v1.0/me", {
        headers: { Authorization: "Bearer test-access-token" },
      })
    })

    it("should merge custom headers with Authorization", async () => {
      vi.mocked(fetch).mockResolvedValueOnce({ ok: true } as Response)

      await graphFetch("https://graph.microsoft.com/v1.0/me", mockToken, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
      })

      expect(fetch).toHaveBeenCalledWith("https://graph.microsoft.com/v1.0/me", {
        method: "POST",
        headers: {
          Authorization: "Bearer test-access-token",
          "Content-Type": "application/json",
        },
      })
    })

    it("should return Left with GraphError on HTTP failure", async () => {
      vi.mocked(fetch).mockResolvedValueOnce({
        ok: false,
        status: 403,
        statusText: "Forbidden",
        json: () => Promise.resolve({ error: { code: "AccessDenied", message: "Insufficient privileges" } }),
      } as Response)

      const result = await graphFetch("https://graph.microsoft.com/v1.0/me", mockToken)

      expect(result.isLeft()).toBe(true)
      result.tapLeft((err) => {
        expect(err._kind).toBe("GraphError")
        expect(err.code).toBe("Forbidden")
        expect(err.message).toBe("AccessDenied: Insufficient privileges")
        expect(err.status).toBe(403)
      })
    })

    it("should return Left with fallback message when error response is not JSON", async () => {
      vi.mocked(fetch).mockResolvedValueOnce({
        ok: false,
        status: 500,
        statusText: "Internal Server Error",
        json: () => Promise.reject(new Error("not JSON")),
      } as Response)

      const result = await graphFetch("https://graph.microsoft.com/v1.0/me", mockToken)

      expect(result.isLeft()).toBe(true)
      result.tapLeft((err) => {
        expect(err.message).toBe("HTTP 500: Internal Server Error")
        expect(err.status).toBe(500)
      })
    })

    it("should pass through request body in init", async () => {
      vi.mocked(fetch).mockResolvedValueOnce({ ok: true } as Response)

      const body = JSON.stringify({ query: "test" })
      await graphFetch("https://graph.microsoft.com/v1.0/search/query", mockToken, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body,
      })

      expect(fetch).toHaveBeenCalledWith("https://graph.microsoft.com/v1.0/search/query", {
        method: "POST",
        headers: {
          Authorization: "Bearer test-access-token",
          "Content-Type": "application/json",
        },
        body,
      })
    })

    it("should handle 404 responses", async () => {
      vi.mocked(fetch).mockResolvedValueOnce({
        ok: false,
        status: 404,
        statusText: "Not Found",
        json: () => Promise.resolve({ error: { code: "ResourceNotFound", message: "Resource does not exist" } }),
      } as Response)

      const result = await graphFetch("https://graph.microsoft.com/v1.0/me/messages/fake-id", mockToken)

      expect(result.isLeft()).toBe(true)
      result.tapLeft((err) => {
        expect(err.status).toBe(404)
        expect(err.message).toBe("ResourceNotFound: Resource does not exist")
      })
    })

    it("should handle 429 throttling responses", async () => {
      vi.mocked(fetch).mockResolvedValueOnce({
        ok: false,
        status: 429,
        statusText: "Too Many Requests",
        json: () => Promise.resolve({ error: { code: "Throttled", message: "Rate limit exceeded" } }),
      } as Response)

      const result = await graphFetch("https://graph.microsoft.com/v1.0/me", mockToken)

      expect(result.isLeft()).toBe(true)
      result.tapLeft((err) => {
        expect(err.status).toBe(429)
        expect(err.code).toBe("Too Many Requests")
      })
    })
  })
})
