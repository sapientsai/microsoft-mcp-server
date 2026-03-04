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

async function searchInteractive(
  accessToken: string,
  args: z.infer<typeof searchParameters>,
): Promise<Either<GraphError, readonly SearchResult[]>> {
  const fileTypeClause =
    args.fileTypes && args.fileTypes.length > 0
      ? ` (${args.fileTypes.map((ext) => `filetype:${ext}`).join(" OR ")})`
      : ""
  const siteClause = args.siteId ? ` site:${args.siteId}` : ""
  const queryString = `${args.query}${fileTypeClause}${siteClause}`

  const body = {
    requests: [
      {
        entityTypes: ["driveItem"],
        query: { queryString },
        from: 0,
        size: args.top,
      },
    ],
  }

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

  if (result.isLeft()) return Right([] as readonly SearchResult[]) // Graceful fallback for individual drive failures

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

async function searchClientCredentials(
  accessToken: string,
  args: z.infer<typeof searchParameters>,
  siteCache: SiteCache,
): Promise<Either<GraphError, readonly SearchResult[]>> {
  // Single site search
  if (args.siteId) {
    const sitesResult = await siteCache.getSites(accessToken)
    if (sitesResult.isLeft()) return Left(sitesResult.value)

    const sites = sitesResult.orThrow()
    const site = sites.find((s) => s.id === args.siteId)
    if (!site) {
      return Left(graphError("site_not_found", `Site ${args.siteId} not found or not accessible.`, 404))
    }
    const results = await searchDrive(accessToken, site.driveId, site.id, args.query)
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

export function buildSearchTool(
  resolveAccessToken: (session: unknown) => Promise<string>,
  authMode: AuthMode,
  siteCache?: SiteCache,
) {
  return {
    name: "sharepoint_search" as const,
    description:
      "Search SharePoint document libraries by keyword. Returns file metadata (name, driveId, itemId) that can be used with read_document to retrieve document text. Construct the path as /drives/{driveId}/items/{driveItemId}/content. Supports filtering by site and file type.",
    parameters: searchParameters,
    execute: async (args: z.infer<typeof searchParameters>, context: Context<Record<string, never>>) => {
      const accessToken = await resolveAccessToken(context.session)
      context.log.info("SharePoint search", { query: args.query, mode: authMode, siteId: args.siteId })

      const resultsEither =
        authMode === "interactive"
          ? await searchInteractive(accessToken, args)
          : await searchClientCredentials(accessToken, args, siteCache!)

      const results = resultsEither.orThrow()
      return JSON.stringify({ results, totalCount: results.length }, null, 2)
    },
  }
}
