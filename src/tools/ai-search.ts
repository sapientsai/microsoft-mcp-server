import { z } from "zod"

import { AI_SEARCH_API_VERSION, type AiSearchConfig, aiSearchFetch } from "../search/ai-search-client.js"

export type AiSearchResult = {
  readonly score: number
  readonly rerankerScore?: number
  readonly document: Record<string, unknown>
  readonly captions: readonly string[]
  readonly highlights: Record<string, readonly string[]>
}

type AiSearchHit = {
  "@search.score"?: number
  "@search.rerankerScore"?: number
  "@search.captions"?: ReadonlyArray<{ text?: string }>
  "@search.highlights"?: Record<string, readonly string[]>
  [key: string]: unknown
}

type AiSearchResponse = {
  "@odata.count"?: number
  "@search.answers"?: ReadonlyArray<{ text?: string; key?: string; score?: number }>
  value?: readonly AiSearchHit[]
}

const aiSearchParameters = z.object({
  query: z.string().describe("Search text"),
  queryType: z
    .enum(["simple", "full", "semantic"])
    .default("semantic")
    .describe('Query type: "simple" keyword, "full" Lucene syntax, or "semantic" for AI-ranked results'),
  vectorSearch: z
    .boolean()
    .default(false)
    .describe("Enable hybrid search via server-side vectorization (requires index with vector fields)"),
  filter: z.string().optional().describe("OData filter expression"),
  select: z.string().optional().describe("Comma-separated fields to return"),
  top: z.number().min(1).max(50).default(10).describe("Maximum number of results to return (1-50)"),
  skip: z.number().optional().describe("Number of results to skip for pagination"),
  orderby: z.string().optional().describe("OData $orderby expression"),
  highlightFields: z.string().optional().describe("Comma-separated fields to highlight"),
  includeTotalCount: z.boolean().default(false).describe("Include total count of matching documents"),
})

function mapHitToResult(hit: AiSearchHit): AiSearchResult {
  const {
    "@search.score": score,
    "@search.rerankerScore": rerankerScore,
    "@search.captions": captions,
    "@search.highlights": highlights,
    ...document
  } = hit

  return {
    score: score ?? 0,
    ...(rerankerScore !== undefined ? { rerankerScore } : {}),
    document,
    captions: (captions ?? []).flatMap((c) => (c.text ? [c.text] : [])),
    highlights: (highlights as Record<string, readonly string[]>) ?? {},
  }
}

export function buildAiSearchTool(config: AiSearchConfig) {
  const searchUrl = `${config.endpoint}/indexes/${config.indexName}/docs/search?api-version=${AI_SEARCH_API_VERSION}`

  return {
    name: "azure_ai_search" as const,
    description:
      "Search an Azure AI Search index with text, semantic, or hybrid (text + vector) queries. Returns ranked results with scores and optional captions/highlights.",
    parameters: aiSearchParameters,
    execute: async (args: z.infer<typeof aiSearchParameters>) => {
      if (args.queryType === "semantic" && !config.semanticConfiguration) {
        throw new Error(
          "Semantic search requires AZURE_AI_SEARCH_SEMANTIC_CONFIG to be configured. Use queryType 'simple' or 'full' instead.",
        )
      }

      const body: Record<string, unknown> = {
        search: args.query,
        queryType: args.queryType,
        top: args.top,
        count: args.includeTotalCount,
      }

      if (args.filter) body.filter = args.filter
      if (args.skip) body.skip = args.skip
      if (args.orderby) body.orderby = args.orderby
      if (args.highlightFields) body.highlightFields = args.highlightFields

      const selectFields = args.select ?? config.selectFields
      if (selectFields) body.select = selectFields

      if (args.queryType === "semantic") {
        body.semanticConfiguration = config.semanticConfiguration
        body.captions = "extractive"
        body.answers = "extractive"
      }

      if (args.vectorSearch) {
        const fields = config.vectorFields ?? "contentVector"
        body.vectorQueries = [
          {
            kind: "text",
            text: args.query,
            fields,
            k: args.top,
          },
        ]
      }

      const result = await aiSearchFetch(searchUrl, config.apiKey, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
      })

      const response = result.orThrow()
      const data = (await response.json()) as AiSearchResponse

      const results = (data.value ?? []).map(mapHitToResult)
      const output: Record<string, unknown> = { results }

      if (args.includeTotalCount && data["@odata.count"] !== undefined) {
        output.totalCount = data["@odata.count"]
      }

      if (data["@search.answers"] && data["@search.answers"].length > 0) {
        output.answers = data["@search.answers"]
      }

      return JSON.stringify(output, null, 2)
    },
  }
}
