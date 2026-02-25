import { GRAPH_BASE_URL } from "../index.js"

export type SiteInfo = {
  readonly id: string
  readonly name: string
  readonly webUrl: string
  readonly driveId: string
}

export type SiteCache = {
  readonly getSites: (accessToken: string) => Promise<readonly SiteInfo[]>
  readonly invalidate: () => void
}

type GraphSite = {
  id: string
  displayName?: string
  name?: string
  webUrl?: string
  isPersonalSite?: boolean
}

type GraphSitesResponse = {
  value: readonly GraphSite[]
  "@odata.nextLink"?: string
}

type GraphDriveResponse = {
  id: string
}

const CACHE_TTL_MS = 60 * 60 * 1000 // 1 hour
const MAX_SITES = 100

async function fetchPage(
  url: string,
  accessToken: string,
): Promise<{ sites: readonly GraphSite[]; nextLink?: string }> {
  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` },
  })

  if (!response.ok) {
    throw new Error(`Failed to fetch sites: ${response.status} ${response.statusText}`)
  }

  const data = (await response.json()) as GraphSitesResponse
  return { sites: data.value, nextLink: data["@odata.nextLink"] }
}

async function fetchAllSites(accessToken: string): Promise<readonly GraphSite[]> {
  const firstPage = await fetchPage(`${GRAPH_BASE_URL}/v1.0/sites/getAllSites?$top=100`, accessToken)
  const accumulated: GraphSite[] = [...firstPage.sites]

  const collectPages = async (nextLink: string | undefined): Promise<void> => {
    if (!nextLink || accumulated.length >= MAX_SITES) {
      if (accumulated.length >= MAX_SITES) {
        console.warn(`[site-cache] Site count exceeds ${MAX_SITES}, truncating. Some sites may not be searched.`)
      }
      return
    }
    const page = await fetchPage(nextLink, accessToken)
    accumulated.push(...page.sites)
    return collectPages(page.nextLink)
  }

  await collectPages(firstPage.nextLink)
  return accumulated.slice(0, MAX_SITES)
}

async function resolveDriveId(accessToken: string, siteId: string): Promise<string | undefined> {
  try {
    const response = await fetch(`${GRAPH_BASE_URL}/v1.0/sites/${siteId}/drive`, {
      headers: { Authorization: `Bearer ${accessToken}` },
    })

    if (!response.ok) return undefined

    const data = (await response.json()) as GraphDriveResponse
    return data.id
  } catch {
    return undefined
  }
}

export function createSiteCache(): SiteCache {
  const state: { sites: readonly SiteInfo[] | null; fetchedAt: number } = {
    sites: null,
    fetchedAt: 0,
  }

  const getSites = async (accessToken: string): Promise<readonly SiteInfo[]> => {
    const now = Date.now()
    if (state.sites && now - state.fetchedAt < CACHE_TTL_MS) {
      return state.sites
    }

    const allSites = await fetchAllSites(accessToken)
    const nonPersonalSites = allSites.filter((s) => !s.isPersonalSite)

    const results = await Promise.allSettled(
      nonPersonalSites.map(async (site): Promise<SiteInfo | undefined> => {
        const driveId = await resolveDriveId(accessToken, site.id)
        if (!driveId) return undefined
        return {
          id: site.id,
          name: site.displayName ?? site.name ?? site.id,
          webUrl: site.webUrl ?? "",
          driveId,
        }
      }),
    )

    const sites = results
      .filter((r): r is PromiseFulfilledResult<SiteInfo | undefined> => r.status === "fulfilled")
      .map((r) => r.value)
      .filter((s): s is SiteInfo => s !== undefined)

    state.sites = sites
    state.fetchedAt = now

    return sites
  }

  const invalidate = (): void => {
    state.sites = null
    state.fetchedAt = 0
  }

  return { getSites, invalidate }
}
