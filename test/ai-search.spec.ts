import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

import type { AiSearchConfig } from "../src/search/ai-search-client.js"
import { buildAiSearchTool } from "../src/tools/ai-search.js"

describe("AI Search Tool", () => {
  const baseConfig: AiSearchConfig = {
    endpoint: "https://my-search.search.windows.net",
    apiKey: "test-key",
    indexName: "my-index",
    semanticConfiguration: "my-semantic-config",
    vectorFields: "contentVector",
  }

  const searchUrl = `${baseConfig.endpoint}/indexes/${baseConfig.indexName}/docs/search?api-version=2025-09-01`

  beforeEach(() => {
    vi.stubGlobal("fetch", vi.fn())
  })

  afterEach(() => {
    vi.unstubAllGlobals()
  })

  function mockSuccessResponse(value: unknown[] = [], extras: Record<string, unknown> = {}) {
    vi.mocked(fetch).mockResolvedValueOnce({
      ok: true,
      status: 200,
      json: () => Promise.resolve({ value, ...extras }),
    } as Response)
  }

  function mockErrorResponse(status: number, statusText: string, errorBody?: Record<string, unknown>) {
    vi.mocked(fetch).mockResolvedValueOnce({
      ok: false,
      status,
      statusText,
      json: () => Promise.resolve(errorBody ?? {}),
    } as Response)
  }

  function getPostedBody(): Record<string, unknown> {
    const [, init] = vi.mocked(fetch).mock.calls[0] as [string, RequestInit]
    return JSON.parse(init.body as string) as Record<string, unknown>
  }

  describe("simple text search", () => {
    it("should send correct POST body for simple search", async () => {
      mockSuccessResponse()
      const tool = buildAiSearchTool(baseConfig)

      await tool.execute({
        query: "test query",
        queryType: "simple",
        vectorSearch: false,
        top: 10,
        includeTotalCount: false,
      })

      expect(fetch).toHaveBeenCalledWith(searchUrl, expect.objectContaining({ method: "POST" }))
      const body = getPostedBody()
      expect(body.search).toBe("test query")
      expect(body.queryType).toBe("simple")
      expect(body.top).toBe(10)
      expect(body.vectorQueries).toBeUndefined()
      expect(body.semanticConfiguration).toBeUndefined()
    })
  })

  describe("semantic search", () => {
    it("should include semantic configuration, captions, and answers", async () => {
      mockSuccessResponse()
      const tool = buildAiSearchTool(baseConfig)

      await tool.execute({
        query: "what is AI?",
        queryType: "semantic",
        vectorSearch: false,
        top: 5,
        includeTotalCount: false,
      })

      const body = getPostedBody()
      expect(body.queryType).toBe("semantic")
      expect(body.semanticConfiguration).toBe("my-semantic-config")
      expect(body.captions).toBe("extractive")
      expect(body.answers).toBe("extractive")
    })

    it("should throw when semantic config is missing", async () => {
      const configNoSemantic: AiSearchConfig = {
        endpoint: "https://my-search.search.windows.net",
        apiKey: "test-key",
        indexName: "my-index",
      }
      const tool = buildAiSearchTool(configNoSemantic)

      await expect(
        tool.execute({ query: "test", queryType: "semantic", vectorSearch: false, top: 10, includeTotalCount: false }),
      ).rejects.toThrow("Semantic search requires AZURE_AI_SEARCH_SEMANTIC_CONFIG")
    })
  })

  describe("hybrid search", () => {
    it("should include vectorQueries when vectorSearch is true", async () => {
      mockSuccessResponse()
      const tool = buildAiSearchTool(baseConfig)

      await tool.execute({
        query: "vector test",
        queryType: "semantic",
        vectorSearch: true,
        top: 10,
        includeTotalCount: false,
      })

      const body = getPostedBody()
      expect(body.vectorQueries).toEqual([{ kind: "text", text: "vector test", fields: "contentVector", k: 10 }])
    })

    it("should default vectorFields to contentVector when not configured", async () => {
      mockSuccessResponse()
      const configNoVector: AiSearchConfig = {
        ...baseConfig,
        vectorFields: undefined,
      }
      const tool = buildAiSearchTool(configNoVector)

      await tool.execute({ query: "test", queryType: "simple", vectorSearch: true, top: 5, includeTotalCount: false })

      const body = getPostedBody()
      expect(body.vectorQueries).toEqual([{ kind: "text", text: "test", fields: "contentVector", k: 5 }])
    })
  })

  describe("filter and select", () => {
    it("should pass through filter", async () => {
      mockSuccessResponse()
      const tool = buildAiSearchTool(baseConfig)

      await tool.execute({
        query: "test",
        queryType: "simple",
        vectorSearch: false,
        top: 10,
        filter: "category eq 'docs'",
        includeTotalCount: false,
      })

      const body = getPostedBody()
      expect(body.filter).toBe("category eq 'docs'")
    })

    it("should use select from args over config default", async () => {
      mockSuccessResponse()
      const configWithSelect: AiSearchConfig = {
        ...baseConfig,
        selectFields: "title,content",
      }
      const tool = buildAiSearchTool(configWithSelect)

      await tool.execute({
        query: "test",
        queryType: "simple",
        vectorSearch: false,
        top: 10,
        select: "id,name",
        includeTotalCount: false,
      })

      const body = getPostedBody()
      expect(body.select).toBe("id,name")
    })

    it("should use config selectFields when args.select is not provided", async () => {
      mockSuccessResponse()
      const configWithSelect: AiSearchConfig = {
        ...baseConfig,
        selectFields: "title,content",
      }
      const tool = buildAiSearchTool(configWithSelect)

      await tool.execute({ query: "test", queryType: "simple", vectorSearch: false, top: 10, includeTotalCount: false })

      const body = getPostedBody()
      expect(body.select).toBe("title,content")
    })
  })

  describe("error handling", () => {
    it("should throw on 401 Unauthorized", async () => {
      mockErrorResponse(401, "Unauthorized", { error: { code: "AuthorizationFailure", message: "Invalid API key" } })
      const tool = buildAiSearchTool(baseConfig)

      await expect(
        tool.execute({ query: "test", queryType: "simple", vectorSearch: false, top: 10, includeTotalCount: false }),
      ).rejects.toThrow()
    })

    it("should throw on 400 Bad Request", async () => {
      mockErrorResponse(400, "Bad Request", { error: { message: "Invalid filter expression" } })
      const tool = buildAiSearchTool(baseConfig)

      await expect(
        tool.execute({
          query: "test",
          queryType: "simple",
          vectorSearch: false,
          top: 10,
          filter: "bad filter",
          includeTotalCount: false,
        }),
      ).rejects.toThrow()
    })

    it("should throw on 404 Not Found", async () => {
      mockErrorResponse(404, "Not Found", { error: { code: "IndexNotFound", message: "Index does not exist" } })
      const tool = buildAiSearchTool(baseConfig)

      await expect(
        tool.execute({ query: "test", queryType: "simple", vectorSearch: false, top: 10, includeTotalCount: false }),
      ).rejects.toThrow()
    })
  })

  describe("response mapping", () => {
    it("should return empty results for empty response", async () => {
      mockSuccessResponse([])
      const tool = buildAiSearchTool(baseConfig)

      const output = await tool.execute({
        query: "nothing",
        queryType: "simple",
        vectorSearch: false,
        top: 10,
        includeTotalCount: false,
      })

      const parsed = JSON.parse(output) as { results: unknown[] }
      expect(parsed.results).toEqual([])
    })

    it("should map hits to AiSearchResult", async () => {
      mockSuccessResponse([
        {
          "@search.score": 1.5,
          "@search.rerankerScore": 3.2,
          "@search.captions": [{ text: "This is a caption" }],
          "@search.highlights": { content: ["highlighted <em>text</em>"] },
          id: "doc-1",
          title: "Test Document",
        },
      ])
      const tool = buildAiSearchTool(baseConfig)

      const output = await tool.execute({
        query: "test",
        queryType: "semantic",
        vectorSearch: false,
        top: 10,
        includeTotalCount: false,
      })

      const parsed = JSON.parse(output) as {
        results: Array<{
          score: number
          rerankerScore: number
          document: Record<string, unknown>
          captions: string[]
          highlights: Record<string, string[]>
        }>
      }
      expect(parsed.results).toHaveLength(1)
      expect(parsed.results[0].score).toBe(1.5)
      expect(parsed.results[0].rerankerScore).toBe(3.2)
      expect(parsed.results[0].document).toEqual({ id: "doc-1", title: "Test Document" })
      expect(parsed.results[0].captions).toEqual(["This is a caption"])
      expect(parsed.results[0].highlights).toEqual({ content: ["highlighted <em>text</em>"] })
    })

    it("should include semantic answers when present", async () => {
      mockSuccessResponse([], {
        "@search.answers": [{ text: "The answer is 42", key: "doc-1", score: 0.95 }],
      })
      const tool = buildAiSearchTool(baseConfig)

      const output = await tool.execute({
        query: "what is the answer?",
        queryType: "semantic",
        vectorSearch: false,
        top: 10,
        includeTotalCount: false,
      })

      const parsed = JSON.parse(output) as { answers: Array<{ text: string }> }
      expect(parsed.answers).toHaveLength(1)
      expect(parsed.answers[0].text).toBe("The answer is 42")
    })

    it("should include totalCount when requested", async () => {
      mockSuccessResponse([], { "@odata.count": 42 })
      const tool = buildAiSearchTool(baseConfig)

      const output = await tool.execute({
        query: "test",
        queryType: "simple",
        vectorSearch: false,
        top: 10,
        includeTotalCount: true,
      })

      const parsed = JSON.parse(output) as { totalCount: number }
      expect(parsed.totalCount).toBe(42)
    })
  })
})
