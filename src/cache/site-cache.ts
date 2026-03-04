import { type Either, Left, Option, Right } from "functype"

import type { GraphError } from "../errors.js"
import { GRAPH_BASE_URL, graphFetch } from "../graph/client.js"

export type SiteInfo = {
  readonly id: string
  readonly name: string
  readonly webUrl: string
  readonly driveId: string
}

export type SiteCache = {
  readonly getSites: (accessToken: string) => Promise<Either<GraphError, readonly SiteInfo[]>>
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
): Promise<Either<GraphError, { sites: readonly GraphSite[]; nextLink: string | undefined }>> {
  const result = await graphFetch(url, accessToken)
  if (result.isLeft()) return Left(result.value)

  const data = (await result.orThrow().json()) as GraphSitesResponse
  return Right({ sites: data.value, nextLink: data["@odata.nextLink"] })
}

async function fetchAllSites(accessToken: string): Promise<Either<GraphError, readonly GraphSite[]>> {
  const firstResult = await fetchPage(`${GRAPH_BASE_URL}/v1.0/sites/getAllSites?$top=100`, accessToken)
  if (firstResult.isLeft()) return Left(firstResult.value)

  const firstPage = firstResult.orThrow()
  const accumulated: GraphSite[] = [...firstPage.sites]

  const collectPages = async (nextLink: string | undefined): Promise<Either<GraphError, void>> => {
    if (!nextLink || accumulated.length >= MAX_SITES) {
      if (accumulated.length >= MAX_SITES) {
        console.warn(`[site-cache] Site count exceeds ${MAX_SITES}, truncating. Some sites may not be searched.`)
      }
      return Right(undefined as void)
    }
    const pageResult = await fetchPage(nextLink, accessToken)
    if (pageResult.isLeft()) return Left(pageResult.value)

    accumulated.push(...pageResult.orThrow().sites)
    return collectPages(pageResult.orThrow().nextLink)
  }

  const collectResult = await collectPages(firstPage.nextLink)
  if (collectResult.isLeft()) return Left(collectResult.value)

  return Right(accumulated.slice(0, MAX_SITES) as readonly GraphSite[])
}

async function resolveDriveId(accessToken: string, siteId: string): Promise<Option<string>> {
  const result = await graphFetch(`${GRAPH_BASE_URL}/v1.0/sites/${siteId}/drive`, accessToken)
  if (result.isLeft()) return Option<string>(undefined)

  try {
    const data = (await result.orThrow().json()) as GraphDriveResponse
    return Option(data.id)
  } catch {
    return Option<string>(undefined)
  }
}

export function createSiteCache(): SiteCache {
  const state: { sites: readonly SiteInfo[] | null; fetchedAt: number } = {
    sites: null,
    fetchedAt: 0,
  }

  const getSites = async (accessToken: string): Promise<Either<GraphError, readonly SiteInfo[]>> => {
    const now = Date.now()
    if (state.sites && now - state.fetchedAt < CACHE_TTL_MS) {
      return Right(state.sites)
    }

    const allSitesResult = await fetchAllSites(accessToken)
    if (allSitesResult.isLeft()) return Left(allSitesResult.value)

    const allSites = allSitesResult.orThrow()
    const nonPersonalSites = allSites.filter((s) => !s.isPersonalSite)

    const siteOptions = await Promise.all(
      nonPersonalSites.map(async (site): Promise<Option<SiteInfo>> => {
        const driveIdOpt = await resolveDriveId(accessToken, site.id)
        return driveIdOpt.map((driveId) => ({
          id: site.id,
          name: Option(site.displayName).or(Option(site.name)).orElse(site.id),
          webUrl: Option(site.webUrl).orElse(""),
          driveId,
        }))
      }),
    )

    const sites = siteOptions.filter((opt) => opt.isSome()).map((opt) => opt.orThrow())

    state.sites = sites
    state.fetchedAt = now

    return Right(sites as readonly SiteInfo[])
  }

  const invalidate = (): void => {
    state.sites = null
    state.fetchedAt = 0
  }

  return { getSites, invalidate }
}
