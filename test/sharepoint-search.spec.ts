import { Right } from "functype"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

import type { SiteCache, SiteInfo } from "../src/cache/site-cache.js"
import { buildSearchTool } from "../src/tools/sharepoint-search.js"

describe("SharePoint Search Tool", () => {
  const mockToken = "test-access-token"
  const mockResolveToken = async () => mockToken
  const mockLog = { info: vi.fn() }

  beforeEach(() => {
    vi.stubGlobal("fetch", vi.fn())
  })

  afterEach(() => {
    vi.unstubAllGlobals()
  })

  describe("Interactive mode", () => {
    it("should search via POST /search/query", async () => {
      const searchResponse = {
        value: [
          {
            hitsContainers: [
              {
                hits: [
                  {
                    resource: {
                      id: "item-1",
                      name: "report.docx",
                      webUrl: "https://example.sharepoint.com/report.docx",
                      lastModifiedDateTime: "2024-01-15T10:00:00Z",
                      size: 1024,
                      file: { mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
                      parentReference: { driveId: "drive-1", siteId: "site-1" },
                    },
                    _summary: "...matching <b>text</b>...",
                  },
                ],
                total: 1,
              },
            ],
          },
        ],
      }

      vi.mocked(fetch).mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve(searchResponse),
      } as Response)

      const tool = buildSearchTool(mockResolveToken, "interactive")
      const result = await tool.execute({ query: "quarterly report", top: 10 }, { session: {}, log: mockLog } as never)
      const parsed = JSON.parse(result as string) as {
        results: Array<{ name: string; hitHighlights: string[] }>
        totalCount: number
      }

      expect(parsed.results).toHaveLength(1)
      expect(parsed.results[0]?.name).toBe("report.docx")
      expect(parsed.results[0]?.hitHighlights).toEqual(["...matching <b>text</b>..."])
      expect(fetch).toHaveBeenCalledWith(
        "https://graph.microsoft.com/v1.0/search/query",
        expect.objectContaining({ method: "POST" }),
      )
    })

    it("should include fileTypes in query string", async () => {
      const searchResponse = { value: [{ hitsContainers: [{ hits: [] }] }] }

      vi.mocked(fetch).mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve(searchResponse),
      } as Response)

      const tool = buildSearchTool(mockResolveToken, "interactive")
      await tool.execute({ query: "data", top: 10, fileTypes: ["docx", "pdf"] }, { session: {}, log: mockLog } as never)

      const callBody = JSON.parse(vi.mocked(fetch).mock.calls[0]![1]!.body as string) as {
        requests: Array<{ query: { queryString: string } }>
      }
      expect(callBody.requests[0]?.query.queryString).toContain("filetype:docx")
      expect(callBody.requests[0]?.query.queryString).toContain("filetype:pdf")
    })

    it("should resolve siteId to URL and scope by site in query string", async () => {
      const siteResponse = { webUrl: "https://example.sharepoint.com/sites/TestSite" }
      const searchResponse = { value: [{ hitsContainers: [{ hits: [] }] }] }

      vi.mocked(fetch)
        .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve(siteResponse) } as Response)
        .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve(searchResponse) } as Response)

      const tool = buildSearchTool(mockResolveToken, "interactive")
      await tool.execute({ query: "budget", top: 10, siteId: "site-123" }, { session: {}, log: mockLog } as never)

      // First call resolves site URL
      expect(fetch).toHaveBeenCalledWith(
        "https://graph.microsoft.com/v1.0/sites/site-123?$select=webUrl",
        expect.objectContaining({ headers: expect.objectContaining({ Authorization: `Bearer ${mockToken}` }) }),
      )
      // Second call is the search with resolved URL in KQL
      const callBody = JSON.parse(vi.mocked(fetch).mock.calls[1]![1]!.body as string) as {
        requests: Array<{ query: { queryString: string } }>
      }
      expect(callBody.requests[0]?.query.queryString).toContain("site:https://example.sharepoint.com/sites/TestSite")
    })

    it("should return empty results gracefully", async () => {
      const searchResponse = { value: [{ hitsContainers: [{ hits: [] }] }] }

      vi.mocked(fetch).mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve(searchResponse),
      } as Response)

      const tool = buildSearchTool(mockResolveToken, "interactive")
      const result = await tool.execute({ query: "nonexistent", top: 10 }, { session: {}, log: mockLog } as never)
      const parsed = JSON.parse(result as string) as { results: unknown[]; totalCount: number }

      expect(parsed.results).toHaveLength(0)
      expect(parsed.totalCount).toBe(0)
    })
  })

  describe("Client credentials mode", () => {
    const mockSites: readonly SiteInfo[] = [
      { id: "site-1", name: "Site 1", webUrl: "https://example.sharepoint.com/sites/1", driveId: "drive-1" },
      { id: "site-2", name: "Site 2", webUrl: "https://example.sharepoint.com/sites/2", driveId: "drive-2" },
    ]

    const mockSiteCache: SiteCache = {
      getSites: vi.fn().mockResolvedValue(Right(mockSites)),
      invalidate: vi.fn(),
    }

    it("should search a single site by siteId without calling getSites", async () => {
      const driveResponse = { id: "drive-1" }
      const driveSearchResponse = {
        value: [
          {
            id: "item-1",
            name: "report.pdf",
            webUrl: "https://example.sharepoint.com/report.pdf",
            lastModifiedDateTime: "2024-01-15T10:00:00Z",
            size: 2048,
            file: { mimeType: "application/pdf" },
            parentReference: { driveId: "drive-1", siteId: "site-1" },
          },
        ],
      }

      vi.mocked(fetch)
        .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve(driveResponse) } as Response)
        .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve(driveSearchResponse) } as Response)

      const tool = buildSearchTool(mockResolveToken, "clientCredentials", mockSiteCache)
      const result = await tool.execute({ query: "report", top: 10, siteId: "site-1" }, {
        session: {},
        log: mockLog,
      } as never)
      const parsed = JSON.parse(result as string) as { results: Array<{ name: string; driveId: string }> }

      expect(parsed.results).toHaveLength(1)
      expect(parsed.results[0]?.name).toBe("report.pdf")
      expect(fetch).toHaveBeenCalledTimes(2)
      expect(fetch).toHaveBeenCalledWith(
        "https://graph.microsoft.com/v1.0/sites/site-1/drive",
        expect.objectContaining({ headers: expect.objectContaining({ Authorization: `Bearer ${mockToken}` }) }),
      )
      expect(mockSiteCache.getSites).not.toHaveBeenCalled()
    })

    it("should propagate error when siteId drive resolution fails", async () => {
      vi.mocked(fetch).mockResolvedValueOnce({
        ok: false,
        status: 403,
        statusText: "Forbidden",
        json: () => Promise.resolve({ error: { code: "AccessDenied", message: "Insufficient privileges" } }),
      } as Response)

      const tool = buildSearchTool(mockResolveToken, "clientCredentials", mockSiteCache)
      await expect(
        tool.execute({ query: "report", top: 10, siteId: "site-1" }, { session: {}, log: mockLog } as never),
      ).rejects.toThrow()
      expect(mockSiteCache.getSites).not.toHaveBeenCalled()
    })

    it("should propagate error when searchDrive fails for single-site search", async () => {
      const driveResponse = { id: "drive-1" }
      vi.mocked(fetch)
        .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve(driveResponse) } as Response)
        .mockResolvedValueOnce({
          ok: false,
          status: 500,
          statusText: "Internal Server Error",
          json: () => Promise.resolve({ error: { code: "ServerError", message: "Something went wrong" } }),
        } as Response)

      const tool = buildSearchTool(mockResolveToken, "clientCredentials", mockSiteCache)
      await expect(
        tool.execute({ query: "report", top: 10, siteId: "site-1" }, { session: {}, log: mockLog } as never),
      ).rejects.toThrow()
    })

    it("should fan-out across all cached sites", async () => {
      const driveSearchResponse1 = {
        value: [
          {
            id: "item-1",
            name: "doc1.docx",
            lastModifiedDateTime: "2024-01-15T10:00:00Z",
            size: 1024,
            file: { mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
          },
        ],
      }
      const driveSearchResponse2 = {
        value: [
          {
            id: "item-2",
            name: "doc2.pdf",
            lastModifiedDateTime: "2024-01-16T10:00:00Z",
            size: 2048,
            file: { mimeType: "application/pdf" },
          },
        ],
      }

      vi.mocked(fetch)
        .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve(driveSearchResponse1) } as Response)
        .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve(driveSearchResponse2) } as Response)

      const tool = buildSearchTool(mockResolveToken, "clientCredentials", mockSiteCache)
      const result = await tool.execute({ query: "document", top: 10 }, { session: {}, log: mockLog } as never)
      const parsed = JSON.parse(result as string) as { results: Array<{ name: string }> }

      expect(parsed.results).toHaveLength(2)
      // Sorted by lastModified desc — doc2 is newer
      expect(parsed.results[0]?.name).toBe("doc2.pdf")
      expect(parsed.results[1]?.name).toBe("doc1.docx")
    })

    it("should return partial results when some sites fail", async () => {
      const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {})
      const driveSearchResponse = {
        value: [
          {
            id: "item-1",
            name: "doc.docx",
            lastModifiedDateTime: "2024-01-15T10:00:00Z",
            size: 1024,
            file: { mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
          },
        ],
      }

      vi.mocked(fetch)
        .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve(driveSearchResponse) } as Response)
        .mockResolvedValueOnce({
          ok: false,
          status: 404,
          statusText: "Not Found",
          json: () => Promise.resolve({ error: { message: "Not found" } }),
        } as Response) // site-2 fails

      const tool = buildSearchTool(mockResolveToken, "clientCredentials", mockSiteCache)
      const result = await tool.execute({ query: "doc", top: 10 }, { session: {}, log: mockLog } as never)
      const parsed = JSON.parse(result as string) as { results: Array<{ name: string }> }

      expect(parsed.results).toHaveLength(1)
      expect(parsed.results[0]?.name).toBe("doc.docx")
      expect(warnSpy).toHaveBeenCalledWith(expect.stringContaining("[sharepoint-search] Site search failed"))
      warnSpy.mockRestore()
    })

    it("should post-filter by file extension", async () => {
      const driveSearchResponse = {
        value: [
          {
            id: "item-1",
            name: "report.docx",
            lastModifiedDateTime: "2024-01-15T10:00:00Z",
            size: 1024,
            file: { mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
          },
          {
            id: "item-2",
            name: "data.xlsx",
            lastModifiedDateTime: "2024-01-16T10:00:00Z",
            size: 2048,
            file: { mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
          },
          {
            id: "item-3",
            name: "image.png",
            lastModifiedDateTime: "2024-01-17T10:00:00Z",
            size: 4096,
            file: { mimeType: "image/png" },
          },
        ],
      }

      // Only one site to simplify
      const singleSiteCache: SiteCache = {
        getSites: vi.fn().mockResolvedValue(Right([mockSites[0]])),
        invalidate: vi.fn(),
      }

      vi.mocked(fetch).mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve(driveSearchResponse),
      } as Response)

      const tool = buildSearchTool(mockResolveToken, "clientCredentials", singleSiteCache)
      const result = await tool.execute({ query: "data", top: 10, fileTypes: ["docx", "xlsx"] }, {
        session: {},
        log: mockLog,
      } as never)
      const parsed = JSON.parse(result as string) as { results: Array<{ name: string }> }

      expect(parsed.results).toHaveLength(2)
      expect(parsed.results.map((r) => r.name)).toEqual(expect.arrayContaining(["report.docx", "data.xlsx"]))
    })

    it("should use defaultSiteId when no siteId arg provided", async () => {
      const driveResponse = { id: "drive-default" }
      const driveSearchResponse = {
        value: [
          {
            id: "item-1",
            name: "default-site-doc.pdf",
            webUrl: "https://example.sharepoint.com/default-site-doc.pdf",
            lastModifiedDateTime: "2024-01-15T10:00:00Z",
            size: 2048,
            file: { mimeType: "application/pdf" },
            parentReference: { driveId: "drive-default", siteId: "default-site" },
          },
        ],
      }

      vi.mocked(fetch)
        .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve(driveResponse) } as Response)
        .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve(driveSearchResponse) } as Response)

      const unusedCache: SiteCache = {
        getSites: vi.fn().mockResolvedValue(Right([])),
        invalidate: vi.fn(),
      }
      const tool = buildSearchTool(mockResolveToken, "clientCredentials", unusedCache, "default-site")
      const result = await tool.execute({ query: "report", top: 10 }, { session: {}, log: mockLog } as never)
      const parsed = JSON.parse(result as string) as { results: Array<{ name: string; siteId: string }> }

      expect(parsed.results).toHaveLength(1)
      expect(parsed.results[0]?.name).toBe("default-site-doc.pdf")
      expect(fetch).toHaveBeenCalledWith(
        "https://graph.microsoft.com/v1.0/sites/default-site/drive",
        expect.objectContaining({ headers: expect.objectContaining({ Authorization: `Bearer ${mockToken}` }) }),
      )
      expect(unusedCache.getSites).not.toHaveBeenCalled()
    })

    it("should let explicit siteId arg override defaultSiteId", async () => {
      const driveResponse = { id: "drive-explicit" }
      const driveSearchResponse = {
        value: [
          {
            id: "item-1",
            name: "explicit-site-doc.pdf",
            webUrl: "https://example.sharepoint.com/explicit-site-doc.pdf",
            lastModifiedDateTime: "2024-01-15T10:00:00Z",
            size: 2048,
            file: { mimeType: "application/pdf" },
            parentReference: { driveId: "drive-explicit", siteId: "explicit-site" },
          },
        ],
      }

      vi.mocked(fetch)
        .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve(driveResponse) } as Response)
        .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve(driveSearchResponse) } as Response)

      const unusedCache: SiteCache = {
        getSites: vi.fn().mockResolvedValue(Right([])),
        invalidate: vi.fn(),
      }
      const tool = buildSearchTool(mockResolveToken, "clientCredentials", unusedCache, "default-site")
      const result = await tool.execute({ query: "report", top: 10, siteId: "explicit-site" }, {
        session: {},
        log: mockLog,
      } as never)
      const parsed = JSON.parse(result as string) as { results: Array<{ name: string }> }

      expect(parsed.results).toHaveLength(1)
      expect(parsed.results[0]?.name).toBe("explicit-site-doc.pdf")
      expect(fetch).toHaveBeenCalledWith(
        "https://graph.microsoft.com/v1.0/sites/explicit-site/drive",
        expect.objectContaining({ headers: expect.objectContaining({ Authorization: `Bearer ${mockToken}` }) }),
      )
      expect(unusedCache.getSites).not.toHaveBeenCalled()
    })

    it("should fan-out when no defaultSiteId and no siteId arg", async () => {
      const driveSearchResponse1 = {
        value: [
          {
            id: "item-1",
            name: "doc1.docx",
            lastModifiedDateTime: "2024-01-15T10:00:00Z",
            size: 1024,
            file: { mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
          },
        ],
      }
      const driveSearchResponse2 = { value: [] }

      const fanoutCache: SiteCache = {
        getSites: vi.fn().mockResolvedValue(Right(mockSites)),
        invalidate: vi.fn(),
      }

      vi.mocked(fetch)
        .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve(driveSearchResponse1) } as Response)
        .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve(driveSearchResponse2) } as Response)

      const tool = buildSearchTool(mockResolveToken, "clientCredentials", fanoutCache)
      const result = await tool.execute({ query: "doc", top: 10 }, { session: {}, log: mockLog } as never)
      const parsed = JSON.parse(result as string) as { results: Array<{ name: string }> }

      expect(parsed.results).toHaveLength(1)
      expect(fanoutCache.getSites).toHaveBeenCalled()
    })

    it("should return empty results when no sites are available", async () => {
      const emptySiteCache: SiteCache = {
        getSites: vi.fn().mockResolvedValue(Right([])),
        invalidate: vi.fn(),
      }

      const tool = buildSearchTool(mockResolveToken, "clientCredentials", emptySiteCache)
      const result = await tool.execute({ query: "anything", top: 10 }, { session: {}, log: mockLog } as never)
      const parsed = JSON.parse(result as string) as { results: unknown[]; totalCount: number }

      expect(parsed.results).toHaveLength(0)
      expect(parsed.totalCount).toBe(0)
    })
  })

  describe("defaultSiteUrl (SITE_URL env var)", () => {
    it("should use defaultSiteUrl directly in KQL without resolving siteId", async () => {
      const searchResponse = { value: [{ hitsContainers: [{ hits: [] }] }] }

      vi.mocked(fetch).mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve(searchResponse),
      } as Response)

      const tool = buildSearchTool(
        mockResolveToken,
        "interactive",
        undefined,
        undefined,
        undefined,
        "https://example.sharepoint.com/sites/MySite",
      )
      await tool.execute({ query: "budget", top: 10 }, { session: {}, log: mockLog } as never)

      // Only one fetch call (the search), no site URL resolution call
      expect(fetch).toHaveBeenCalledTimes(1)
      const callBody = JSON.parse(vi.mocked(fetch).mock.calls[0]![1]!.body as string) as {
        requests: Array<{ query: { queryString: string } }>
      }
      expect(callBody.requests[0]?.query.queryString).toContain("site:https://example.sharepoint.com/sites/MySite")
    })

    it("should ignore defaultSiteUrl when explicit siteId arg is passed", async () => {
      const siteResponse = { webUrl: "https://example.sharepoint.com/sites/ExplicitSite" }
      const searchResponse = { value: [{ hitsContainers: [{ hits: [] }] }] }

      vi.mocked(fetch)
        .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve(siteResponse) } as Response)
        .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve(searchResponse) } as Response)

      const tool = buildSearchTool(
        mockResolveToken,
        "interactive",
        undefined,
        undefined,
        undefined,
        "https://example.sharepoint.com/sites/DefaultSite",
      )
      await tool.execute({ query: "budget", top: 10, siteId: "explicit-site-id" }, {
        session: {},
        log: mockLog,
      } as never)

      // Should resolve explicit siteId, ignoring defaultSiteUrl
      expect(fetch).toHaveBeenCalledTimes(2)
      expect(fetch).toHaveBeenCalledWith(
        "https://graph.microsoft.com/v1.0/sites/explicit-site-id?$select=webUrl",
        expect.objectContaining({ headers: expect.objectContaining({ Authorization: `Bearer ${mockToken}` }) }),
      )
      const callBody = JSON.parse(vi.mocked(fetch).mock.calls[1]![1]!.body as string) as {
        requests: Array<{ query: { queryString: string } }>
      }
      expect(callBody.requests[0]?.query.queryString).toContain(
        "site:https://example.sharepoint.com/sites/ExplicitSite",
      )
    })

    it("should use defaultSiteUrl with search region", async () => {
      const searchResponse = { value: [{ hitsContainers: [{ hits: [] }] }] }

      vi.mocked(fetch).mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve(searchResponse),
      } as Response)

      const tool = buildSearchTool(
        mockResolveToken,
        "clientCredentials",
        undefined,
        undefined,
        "NAM",
        "https://example.sharepoint.com/sites/MySite",
      )
      await tool.execute({ query: "report", top: 10 }, { session: {}, log: mockLog } as never)

      expect(fetch).toHaveBeenCalledTimes(1)
      const callBody = JSON.parse(vi.mocked(fetch).mock.calls[0]![1]!.body as string) as {
        requests: Array<{ query: { queryString: string }; region?: string }>
      }
      expect(callBody.requests[0]?.query.queryString).toContain("site:https://example.sharepoint.com/sites/MySite")
      expect(callBody.requests[0]?.region).toBe("NAM")
    })
  })

  describe("Search region", () => {
    it("should include region in search request body when provided", async () => {
      const searchResponse = { value: [{ hitsContainers: [{ hits: [] }] }] }

      vi.mocked(fetch).mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve(searchResponse),
      } as Response)

      const tool = buildSearchTool(mockResolveToken, "clientCredentials", undefined, undefined, "NAM")
      await tool.execute({ query: "test", top: 10 }, { session: {}, log: mockLog } as never)

      const callBody = JSON.parse(vi.mocked(fetch).mock.calls[0]![1]!.body as string) as {
        requests: Array<{ region?: string }>
      }
      expect(callBody.requests[0]?.region).toBe("NAM")
    })

    it("should use search API for client credentials when region is set", async () => {
      const searchResponse = {
        value: [
          {
            hitsContainers: [
              {
                hits: [
                  {
                    resource: {
                      id: "item-1",
                      name: "found.docx",
                      webUrl: "https://example.sharepoint.com/found.docx",
                      lastModifiedDateTime: "2024-01-15T10:00:00Z",
                      size: 1024,
                      file: { mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
                      parentReference: { driveId: "drive-1", siteId: "site-1" },
                    },
                  },
                ],
              },
            ],
          },
        ],
      }

      vi.mocked(fetch).mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve(searchResponse),
      } as Response)

      const unusedCache: SiteCache = {
        getSites: vi.fn().mockResolvedValue(Right([])),
        invalidate: vi.fn(),
      }

      const tool = buildSearchTool(mockResolveToken, "clientCredentials", unusedCache, undefined, "NAM")
      const result = await tool.execute({ query: "test", top: 10 }, { session: {}, log: mockLog } as never)
      const parsed = JSON.parse(result as string) as { results: Array<{ name: string }> }

      expect(parsed.results).toHaveLength(1)
      expect(parsed.results[0]?.name).toBe("found.docx")
      // Should use /search/query, NOT drive search fan-out
      expect(fetch).toHaveBeenCalledWith(
        "https://graph.microsoft.com/v1.0/search/query",
        expect.objectContaining({ method: "POST" }),
      )
      expect(unusedCache.getSites).not.toHaveBeenCalled()
    })

    it("should resolve siteId to URL in KQL when using search API with region", async () => {
      const siteResponse = { webUrl: "https://example.sharepoint.com/sites/TestSite" }
      const searchResponse = { value: [{ hitsContainers: [{ hits: [] }] }] }

      vi.mocked(fetch)
        .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve(siteResponse) } as Response)
        .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve(searchResponse) } as Response)

      const tool = buildSearchTool(mockResolveToken, "clientCredentials", undefined, undefined, "NAM")
      await tool.execute({ query: "budget", top: 10, siteId: "site-abc" }, { session: {}, log: mockLog } as never)

      const callBody = JSON.parse(vi.mocked(fetch).mock.calls[1]![1]!.body as string) as {
        requests: Array<{ query: { queryString: string }; region?: string }>
      }
      expect(callBody.requests[0]?.query.queryString).toContain("site:https://example.sharepoint.com/sites/TestSite")
      expect(callBody.requests[0]?.region).toBe("NAM")
    })

    it("should not include region when not provided", async () => {
      const searchResponse = { value: [{ hitsContainers: [{ hits: [] }] }] }

      vi.mocked(fetch).mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve(searchResponse),
      } as Response)

      const tool = buildSearchTool(mockResolveToken, "interactive")
      await tool.execute({ query: "test", top: 10 }, { session: {}, log: mockLog } as never)

      const callBody = JSON.parse(vi.mocked(fetch).mock.calls[0]![1]!.body as string) as {
        requests: Array<{ region?: string }>
      }
      expect(callBody.requests[0]?.region).toBeUndefined()
    })
  })

  describe("Error handling", () => {
    it("should throw Error with message instead of plain GraphError object", async () => {
      vi.mocked(fetch).mockResolvedValueOnce({
        ok: false,
        status: 400,
        statusText: "Bad Request",
        json: () =>
          Promise.resolve({
            error: { code: "SearchRequest", message: "Region is required when request with application permission." },
          }),
      } as Response)

      const tool = buildSearchTool(mockResolveToken, "interactive")
      await expect(tool.execute({ query: "test", top: 10 }, { session: {}, log: mockLog } as never)).rejects.toThrow(
        "SearchRequest: Region is required when request with application permission. (HTTP 400)",
      )
    })

    it("should include error code and status in thrown Error message", async () => {
      vi.mocked(fetch).mockResolvedValueOnce({
        ok: false,
        status: 403,
        statusText: "Forbidden",
        json: () => Promise.resolve({ error: { code: "AccessDenied", message: "Insufficient privileges" } }),
      } as Response)

      const tool = buildSearchTool(mockResolveToken, "clientCredentials", undefined, undefined, "NAM")
      await expect(tool.execute({ query: "test", top: 10 }, { session: {}, log: mockLog } as never)).rejects.toThrow(
        "AccessDenied: Insufficient privileges (HTTP 403)",
      )
    })
  })
})
