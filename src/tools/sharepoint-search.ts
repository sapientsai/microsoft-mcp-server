import type { Context } from "fastmcp"
import { type Either, Left, List, Option, Right } from "functype"
import { z } from "zod"

import type { AuthMode } from "../auth/types.js"
import type { SiteCache } from "../cache/site-cache.js"
import { type GraphError, graphError } from "../errors.js"
import { GRAPH_BASE_URL, graphFetch } from "../graph/client.js"

export type SearchResult = {
  readonly name: string
  readonly driveItemId: string
  readonly driveId: string
  readonly siteId: string
  readonly webUrl: string
  readonly lastModified: string
  readonly size: number
  readonly mimeType: string
  readonly hitHighlights: readonly string[]
}

type SearchHit = {
  resource?: {
    id?: string
    name?: string
    webUrl?: string
    lastModifiedDateTime?: string
    size?: number
    file?: { mimeType?: string }
    parentReference?: { driveId?: string; siteId?: string }
  }
  _summary?: string
}

type SearchResponse = {
  value?: ReadonlyArray<{
    hitsContainers?: ReadonlyArray<{
      hits?: readonly SearchHit[]
      total?: number
      moreResultsAvailable?: boolean
    }>
  }>
}

type DriveSearchItem = {
  id?: string
  name?: string
  webUrl?: string
  lastModifiedDateTime?: string
  size?: number
  file?: { mimeType?: string }
  parentReference?: { driveId?: string; siteId?: string }
}

type DriveSearchResponse = {
  value?: readonly DriveSearchItem[]
}

const searchParameters = z.object({
  query: z.string().describe("Search query keywords"),
  siteId: z.string().optional().describe("Scope search to a specific SharePoint site ID"),
  top: z.number().min(1).max(50).default(10).describe("Maximum number of results to return (1-50)"),
  fileTypes: z.array(z.string()).optional().describe('Filter by file extensions (e.g., ["docx", "pdf", "xlsx"])'),
})

function mapHitToResult(hit: SearchHit): SearchResult {
  const resource = Option(hit.resource)
  return {
    name: resource.flatMap((r) => Option(r.name)).orElse(""),
    driveItemId: resource.flatMap((r) => Option(r.id)).orElse(""),
    driveId: resource.flatMap((r) => Option(r.parentReference?.driveId)).orElse(""),
    siteId: resource.flatMap((r) => Option(r.parentReference?.siteId)).orElse(""),
    webUrl: resource.flatMap((r) => Option(r.webUrl)).orElse(""),
    lastModified: resource.flatMap((r) => Option(r.lastModifiedDateTime)).orElse(""),
    size: resource.flatMap((r) => Option(r.size)).orElse(0),
    mimeType: resource.flatMap((r) => Option(r.file?.mimeType)).orElse(""),
    hitHighlights: Option(hit._summary)
      .map((s) => [s])
      .orElse([]),
  }
}

async function resolveSiteUrl(accessToken: string, siteId: string): Promise<Either<GraphError, string>> {
  const result = await graphFetch(`${GRAPH_BASE_URL}/v1.0/sites/${siteId}?$select=webUrl`, accessToken)
  if (result.isLeft()) return Left(result.value)
  const data = (await result.orThrow().json()) as { webUrl?: string }
  return Option(data.webUrl).match({
    Some: (url) => Right(url),
    None: () => Left(graphError("no_site_url", `Site ${siteId} has no webUrl.`, 404)),
  })
}

async function searchViaSearchApi(
  accessToken: string,
  args: z.infer<typeof searchParameters>,
  region?: string,
  defaultSiteUrl?: string,
): Promise<Either<GraphError, readonly SearchResult[]>> {
  const fileTypeClause =
    args.fileTypes && args.fileTypes.length > 0
      ? ` (${args.fileTypes.map((ext) => `filetype:${ext}`).join(" OR ")})`
      : ""

  // KQL site: requires a URL, not a composite site ID.
  // Use defaultSiteUrl directly when available (no API call needed).
  // Fall back to resolveSiteUrl() only when an explicit siteId arg is passed.
  const siteClause = defaultSiteUrl
    ? Right<GraphError, string>(` site:${defaultSiteUrl}`)
    : args.siteId
      ? await resolveSiteUrl(accessToken, args.siteId).then((r) => r.map((url) => ` site:${url}`))
      : Right<GraphError, string>("")

  if (siteClause.isLeft()) return Left(siteClause.value)
  const queryString = `${args.query}${fileTypeClause}${siteClause.orThrow()}`

  const request: Record<string, unknown> = {
    entityTypes: ["driveItem"],
    query: { queryString },
    from: 0,
    size: args.top,
  }

  if (region) {
    request.region = region
  }

  const body = { requests: [request] }

  const result = await graphFetch(`${GRAPH_BASE_URL}/v1.0/search/query`, accessToken, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  })

  if (result.isLeft()) return Left(result.value)

  const data = (await result.orThrow().json()) as SearchResponse
  const hits = data.value?.[0]?.hitsContainers?.[0]?.hits ?? []

  return Right(hits.map(mapHitToResult) as readonly SearchResult[])
}

async function searchDrive(
  accessToken: string,
  driveId: string,
  siteId: string,
  query: string,
): Promise<Either<GraphError, readonly SearchResult[]>> {
  const result = await graphFetch(
    `${GRAPH_BASE_URL}/v1.0/drives/${driveId}/root/search(q='${encodeURIComponent(query)}')`,
    accessToken,
  )

  if (result.isLeft()) return Left(result.value)

  const data = (await result.orThrow().json()) as DriveSearchResponse
  const items = data.value ?? []

  const results = items
    .filter((item) => item.file) // only files, not folders
    .map(
      (item): SearchResult => ({
        name: Option(item.name).orElse(""),
        driveItemId: Option(item.id).orElse(""),
        driveId: Option(item.parentReference?.driveId).orElse(driveId),
        siteId: Option(item.parentReference?.siteId).orElse(siteId),
        webUrl: Option(item.webUrl).orElse(""),
        lastModified: Option(item.lastModifiedDateTime).orElse(""),
        size: Option(item.size).orElse(0),
        mimeType: Option(item.file?.mimeType).orElse(""),
        hitHighlights: [],
      }),
    )

  return Right(results as readonly SearchResult[])
}

async function resolveDriveForSite(accessToken: string, siteId: string): Promise<Either<GraphError, string>> {
  const result = await graphFetch(`${GRAPH_BASE_URL}/v1.0/sites/${siteId}/drive`, accessToken)
  if (result.isLeft()) return Left(result.value)
  const data = (await result.orThrow().json()) as { id?: string }
  return Option(data.id).match({
    Some: (id) => Right(id),
    None: () => Left(graphError("no_drive", `Site ${siteId} has no default drive.`, 404)),
  })
}

async function searchClientCredentials(
  accessToken: string,
  args: z.infer<typeof searchParameters>,
  siteCache: SiteCache,
): Promise<Either<GraphError, readonly SearchResult[]>> {
  // Single site search — resolve drive directly, no getAllSites needed
  if (args.siteId) {
    const driveResult = await resolveDriveForSite(accessToken, args.siteId)
    if (driveResult.isLeft()) return Left(driveResult.value)
    const driveId = driveResult.orThrow()
    const results = await searchDrive(accessToken, driveId, args.siteId, args.query)
    if (results.isLeft()) return Left(results.value)
    return results.map((r) => filterAndSort(r, args))
  }

  // Fan-out across all cached sites
  const sitesResult = await siteCache.getSites(accessToken)
  if (sitesResult.isLeft()) return Left(sitesResult.value)

  const sites = sitesResult.orThrow()
  if (sites.length === 0) {
    return Right([] as readonly SearchResult[])
  }

  const driveResults = await Promise.all(
    sites.map((site) => searchDrive(accessToken, site.driveId, site.id, args.query)),
  )

  const allResults: SearchResult[] = []
  for (const result of driveResults) {
    if (result.isRight()) {
      allResults.push(...result.orThrow())
    } else {
      result.tapLeft((err) => {
        console.warn(`[sharepoint-search] Site search failed: ${err.message}`)
      })
    }
  }

  return Right(filterAndSort(allResults, args) as readonly SearchResult[])
}

function filterAndSort(
  results: readonly SearchResult[],
  args: z.infer<typeof searchParameters>,
): readonly SearchResult[] {
  const list = List(results as SearchResult[])

  const filtered =
    args.fileTypes && args.fileTypes.length > 0
      ? (() => {
          const extensions = new Set(args.fileTypes!.map((ext) => ext.toLowerCase().replace(/^\./, "")))
          return list.filter((r) => {
            const ext = r.name.split(".").pop()?.toLowerCase()
            return ext !== undefined && extensions.has(ext)
          })
        })()
      : list

  return filtered
    .sorted((a, b) => b.lastModified.localeCompare(a.lastModified))
    .take(args.top)
    .toArray()
}

function throwGraphError(err: GraphError): never {
  throw new Error(`${err.code}: ${err.message} (HTTP ${err.status})`)
}

export function buildSearchTool(
  resolveAccessToken: (session: unknown) => Promise<string>,
  authMode: AuthMode,
  siteCache?: SiteCache,
  defaultSiteId?: string,
  searchRegion?: string,
  defaultSiteUrl?: string,
) {
  return {
    name: "sharepoint_search" as const,
    description:
      "Search SharePoint document libraries by keyword. Returns file metadata (name, driveId, itemId) that can be used with read_document to retrieve document text. Construct the path as /drives/{driveId}/items/{driveItemId}/content. Supports filtering by site and file type.",
    parameters: searchParameters,
    execute: async (args: z.infer<typeof searchParameters>, context: Context<Record<string, never>>) => {
      const accessToken = await resolveAccessToken(context.session)
      context.log.info("SharePoint search", { query: args.query, mode: authMode, siteId: args.siteId })

      const effectiveArgs = defaultSiteId && !args.siteId ? { ...args, siteId: defaultSiteId } : args

      // When defaultSiteUrl is set and no explicit siteId arg, use it directly in KQL (zero API calls).
      // Otherwise fall back to siteId resolution.
      const effectiveSiteUrl = defaultSiteUrl && !args.siteId ? defaultSiteUrl : undefined

      // Use the Search API (/search/query) when region is available (required for app permissions).
      // Falls back to drive-based search only for client credentials without region configured.
      const resultsEither = searchRegion
        ? await searchViaSearchApi(accessToken, effectiveArgs, searchRegion, effectiveSiteUrl)
        : authMode === "interactive"
          ? await searchViaSearchApi(accessToken, effectiveArgs, undefined, effectiveSiteUrl)
          : await searchClientCredentials(accessToken, effectiveArgs, siteCache!)

      const results = resultsEither.fold(throwGraphError, (r) => r)
      return JSON.stringify({ results, totalCount: results.length }, null, 2)
    },
  }
}
