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

    it("should scope by siteId in query string", async () => {
      const searchResponse = { value: [{ hitsContainers: [{ hits: [] }] }] }

      vi.mocked(fetch).mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve(searchResponse),
      } as Response)

      const tool = buildSearchTool(mockResolveToken, "interactive")
      await tool.execute({ query: "budget", top: 10, siteId: "site-123" }, { session: {}, log: mockLog } as never)

      const callBody = JSON.parse(vi.mocked(fetch).mock.calls[0]![1]!.body as string) as {
        requests: Array<{ query: { queryString: string } }>
      }
      expect(callBody.requests[0]?.query.queryString).toContain("site:site-123")
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
      getSites: vi.fn().mockResolvedValue(mockSites),
      invalidate: vi.fn(),
    }

    it("should search a single site by siteId", async () => {
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

      vi.mocked(fetch).mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve(driveSearchResponse),
      } as Response)

      const tool = buildSearchTool(mockResolveToken, "clientCredentials", mockSiteCache)
      const result = await tool.execute({ query: "report", top: 10, siteId: "site-1" }, {
        session: {},
        log: mockLog,
      } as never)
      const parsed = JSON.parse(result as string) as { results: Array<{ name: string; driveId: string }> }

      expect(parsed.results).toHaveLength(1)
      expect(parsed.results[0]?.name).toBe("report.pdf")
      expect(fetch).toHaveBeenCalledTimes(1)
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
        .mockResolvedValueOnce({ ok: false, status: 404 } as Response) // site-2 fails

      const tool = buildSearchTool(mockResolveToken, "clientCredentials", mockSiteCache)
      const result = await tool.execute({ query: "doc", top: 10 }, { session: {}, log: mockLog } as never)
      const parsed = JSON.parse(result as string) as { results: Array<{ name: string }> }

      expect(parsed.results).toHaveLength(1)
      expect(parsed.results[0]?.name).toBe("doc.docx")
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
        getSites: vi.fn().mockResolvedValue([mockSites[0]]),
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

    it("should return empty results when no sites are available", async () => {
      const emptySiteCache: SiteCache = {
        getSites: vi.fn().mockResolvedValue([]),
        invalidate: vi.fn(),
      }

      const tool = buildSearchTool(mockResolveToken, "clientCredentials", emptySiteCache)
      const result = await tool.execute({ query: "anything", top: 10 }, { session: {}, log: mockLog } as never)
      const parsed = JSON.parse(result as string) as { results: unknown[]; totalCount: number }

      expect(parsed.results).toHaveLength(0)
      expect(parsed.totalCount).toBe(0)
    })
  })
})
