import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

import { createSiteCache } from "../src/cache/site-cache.js"

describe("Site Cache", () => {
  const mockToken = "test-access-token"

  beforeEach(() => {
    vi.useFakeTimers()
    vi.stubGlobal("fetch", vi.fn())
  })

  afterEach(() => {
    vi.useRealTimers()
    vi.unstubAllGlobals()
  })

  const makeSitesResponse = (
    sites: Array<{ id: string; displayName: string; isPersonalSite?: boolean }>,
    nextLink?: string,
  ) => ({
    value: sites,
    ...(nextLink ? { "@odata.nextLink": nextLink } : {}),
  })

  const makeDriveResponse = (driveId: string) => ({ id: driveId })

  it("should fetch sites on first call", async () => {
    const sitesResponse = makeSitesResponse([{ id: "site-1", displayName: "Team Site" }])
    const driveResponse = makeDriveResponse("drive-1")

    vi.mocked(fetch)
      .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve(sitesResponse) } as Response)
      .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve(driveResponse) } as Response)

    const cache = createSiteCache()
    const sites = await cache.getSites(mockToken)

    expect(sites).toHaveLength(1)
    expect(sites[0]).toEqual({
      id: "site-1",
      name: "Team Site",
      webUrl: "",
      driveId: "drive-1",
    })
    expect(fetch).toHaveBeenCalledTimes(2) // getAllSites + drive
  })

  it("should return cached sites within TTL", async () => {
    const sitesResponse = makeSitesResponse([{ id: "site-1", displayName: "Team Site" }])
    const driveResponse = makeDriveResponse("drive-1")

    vi.mocked(fetch)
      .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve(sitesResponse) } as Response)
      .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve(driveResponse) } as Response)

    const cache = createSiteCache()

    const sites1 = await cache.getSites(mockToken)
    const sites2 = await cache.getSites(mockToken)

    expect(sites1).toEqual(sites2)
    expect(fetch).toHaveBeenCalledTimes(2) // only the initial fetch
  })

  it("should refresh after 1-hour TTL expires", async () => {
    const sitesResponse1 = makeSitesResponse([{ id: "site-1", displayName: "Site 1" }])
    const sitesResponse2 = makeSitesResponse([
      { id: "site-1", displayName: "Site 1" },
      { id: "site-2", displayName: "Site 2" },
    ])

    vi.mocked(fetch)
      // First fetch
      .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve(sitesResponse1) } as Response)
      .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve(makeDriveResponse("drive-1")) } as Response)
      // Second fetch after TTL
      .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve(sitesResponse2) } as Response)
      .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve(makeDriveResponse("drive-1")) } as Response)
      .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve(makeDriveResponse("drive-2")) } as Response)

    const cache = createSiteCache()

    const sites1 = await cache.getSites(mockToken)
    expect(sites1).toHaveLength(1)

    // Advance past 1-hour TTL
    vi.advanceTimersByTime(61 * 60 * 1000)

    const sites2 = await cache.getSites(mockToken)
    expect(sites2).toHaveLength(2)
  })

  it("should handle pagination", async () => {
    const page1 = makeSitesResponse(
      [{ id: "site-1", displayName: "Site 1" }],
      "https://graph.microsoft.com/v1.0/sites/getAllSites?$skiptoken=abc",
    )
    const page2 = makeSitesResponse([{ id: "site-2", displayName: "Site 2" }])

    vi.mocked(fetch)
      .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve(page1) } as Response)
      .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve(page2) } as Response)
      .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve(makeDriveResponse("drive-1")) } as Response)
      .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve(makeDriveResponse("drive-2")) } as Response)

    const cache = createSiteCache()
    const sites = await cache.getSites(mockToken)

    expect(sites).toHaveLength(2)
    expect(fetch).toHaveBeenCalledTimes(4) // 2 pages + 2 drive lookups
  })

  it("should filter out personal sites", async () => {
    const sitesResponse = makeSitesResponse([
      { id: "site-1", displayName: "Team Site", isPersonalSite: false },
      { id: "site-personal", displayName: "Personal", isPersonalSite: true },
    ])

    vi.mocked(fetch)
      .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve(sitesResponse) } as Response)
      .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve(makeDriveResponse("drive-1")) } as Response)

    const cache = createSiteCache()
    const sites = await cache.getSites(mockToken)

    expect(sites).toHaveLength(1)
    expect(sites[0]?.id).toBe("site-1")
  })

  it("should handle sites where drive lookup fails", async () => {
    const sitesResponse = makeSitesResponse([
      { id: "site-1", displayName: "Site 1" },
      { id: "site-2", displayName: "Site 2" },
    ])

    vi.mocked(fetch)
      .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve(sitesResponse) } as Response)
      .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve(makeDriveResponse("drive-1")) } as Response)
      .mockResolvedValueOnce({ ok: false, status: 404 } as Response) // site-2 drive fails

    const cache = createSiteCache()
    const sites = await cache.getSites(mockToken)

    expect(sites).toHaveLength(1)
    expect(sites[0]?.id).toBe("site-1")
  })

  it("should throw on getAllSites failure", async () => {
    vi.mocked(fetch).mockResolvedValueOnce({
      ok: false,
      status: 403,
      statusText: "Forbidden",
    } as Response)

    const cache = createSiteCache()
    await expect(cache.getSites(mockToken)).rejects.toThrow("Failed to fetch sites: 403 Forbidden")
  })

  it("should clear cache on invalidate()", async () => {
    const sitesResponse = makeSitesResponse([{ id: "site-1", displayName: "Site 1" }])

    vi.mocked(fetch)
      .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve(sitesResponse) } as Response)
      .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve(makeDriveResponse("drive-1")) } as Response)
      .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve(sitesResponse) } as Response)
      .mockResolvedValueOnce({ ok: true, json: () => Promise.resolve(makeDriveResponse("drive-1")) } as Response)

    const cache = createSiteCache()

    await cache.getSites(mockToken)
    expect(fetch).toHaveBeenCalledTimes(2)

    cache.invalidate()

    await cache.getSites(mockToken)
    expect(fetch).toHaveBeenCalledTimes(4) // re-fetched after invalidate
  })
})
