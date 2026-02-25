import type { Context } from "fastmcp"
import { z } from "zod"

import type { AuthMode } from "../auth/types.js"
import type { SiteCache } from "../cache/site-cache.js"
import { GRAPH_BASE_URL } from "../index.js"

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

async function searchInteractive(
  accessToken: string,
  args: z.infer<typeof searchParameters>,
): Promise<readonly SearchResult[]> {
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

  const response = await fetch(`${GRAPH_BASE_URL}/v1.0/search/query`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
  })

  if (!response.ok) {
    const errorText = await response.text()
    throw new Error(`Search API error ${response.status}: ${errorText}`)
  }

  const data = (await response.json()) as SearchResponse
  const hits = data.value?.[0]?.hitsContainers?.[0]?.hits ?? []

  return hits.map((hit): SearchResult => {
    const resource = hit.resource ?? {}
    return {
      name: resource.name ?? "",
      driveItemId: resource.id ?? "",
      driveId: resource.parentReference?.driveId ?? "",
      siteId: resource.parentReference?.siteId ?? "",
      webUrl: resource.webUrl ?? "",
      lastModified: resource.lastModifiedDateTime ?? "",
      size: resource.size ?? 0,
      mimeType: resource.file?.mimeType ?? "",
      hitHighlights: hit._summary ? [hit._summary] : [],
    }
  })
}

async function searchDrive(
  accessToken: string,
  driveId: string,
  siteId: string,
  query: string,
): Promise<readonly SearchResult[]> {
  const response = await fetch(
    `${GRAPH_BASE_URL}/v1.0/drives/${driveId}/root/search(q='${encodeURIComponent(query)}')`,
    {
      headers: { Authorization: `Bearer ${accessToken}` },
    },
  )

  if (!response.ok) return []

  const data = (await response.json()) as DriveSearchResponse
  const items = data.value ?? []

  return items
    .filter((item) => item.file) // only files, not folders
    .map(
      (item): SearchResult => ({
        name: item.name ?? "",
        driveItemId: item.id ?? "",
        driveId: item.parentReference?.driveId ?? driveId,
        siteId: item.parentReference?.siteId ?? siteId,
        webUrl: item.webUrl ?? "",
        lastModified: item.lastModifiedDateTime ?? "",
        size: item.size ?? 0,
        mimeType: item.file?.mimeType ?? "",
        hitHighlights: [],
      }),
    )
}

async function searchClientCredentials(
  accessToken: string,
  args: z.infer<typeof searchParameters>,
  siteCache: SiteCache,
): Promise<readonly SearchResult[]> {
  // Single site search
  if (args.siteId) {
    const sites = await siteCache.getSites(accessToken)
    const site = sites.find((s) => s.id === args.siteId)
    if (!site) {
      throw new Error(`Site ${args.siteId} not found or not accessible.`)
    }
    const results = await searchDrive(accessToken, site.driveId, site.id, args.query)
    return filterAndSort(results, args)
  }

  // Fan-out across all cached sites
  const sites = await siteCache.getSites(accessToken)
  if (sites.length === 0) {
    return []
  }

  const settled = await Promise.allSettled(
    sites.map((site) => searchDrive(accessToken, site.driveId, site.id, args.query)),
  )

  const allResults: SearchResult[] = []
  for (const result of settled) {
    if (result.status === "fulfilled") {
      allResults.push(...result.value)
    } else {
      console.warn(`[sharepoint-search] Site search failed: ${result.reason}`)
    }
  }

  return filterAndSort(allResults, args)
}

function filterAndSort(
  results: readonly SearchResult[],
  args: z.infer<typeof searchParameters>,
): readonly SearchResult[] {
  const filtered =
    args.fileTypes && args.fileTypes.length > 0
      ? (() => {
          const extensions = new Set(args.fileTypes!.map((ext) => ext.toLowerCase().replace(/^\./, "")))
          return results.filter((r) => {
            const ext = r.name.split(".").pop()?.toLowerCase()
            return ext !== undefined && extensions.has(ext)
          })
        })()
      : [...results]

  return filtered.toSorted((a, b) => b.lastModified.localeCompare(a.lastModified)).slice(0, args.top)
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

      const results =
        authMode === "interactive"
          ? await searchInteractive(accessToken, args)
          : await searchClientCredentials(accessToken, args, siteCache!)

      return JSON.stringify({ results, totalCount: results.length }, null, 2)
    },
  }
}
