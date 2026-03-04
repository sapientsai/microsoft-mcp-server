# Azure AI Search Integration

Optional tool that adds vector/semantic search to the MCP server via [Azure AI Search](https://learn.microsoft.com/en-us/azure/search/). When configured, an `azure_ai_search` tool becomes available alongside the existing `sharepoint_search`.

## Prerequisites

1. **Azure AI Search service** — Create one in the [Azure Portal](https://portal.azure.com/#create/Microsoft.Search)
   - Free tier works for testing; Basic or Standard for production
   - Note the **endpoint URL** (e.g. `https://my-search.search.windows.net`)

2. **Search index** — An index must already exist with documents ingested
   - Can be populated via Azure AI Search indexers, SDK, or REST API
   - For SharePoint content, use the [SharePoint Online indexer](https://learn.microsoft.com/en-us/azure/search/search-howto-index-sharepoint-online)

3. **API key** — Get an admin or query key from the Azure Portal
   - Portal → your Search service → Settings → Keys

4. **(Optional) Semantic configuration** — Required for semantic ranking
   - Configure in your index definition under `semanticConfiguration`

5. **(Optional) Vector fields** — Required for hybrid search
   - Index must have vector fields with an embedding model configured
   - [Integrated vectorization](https://learn.microsoft.com/en-us/azure/search/vector-search-integrated-vectorization) handles server-side embedding

## Environment Variables

| Variable                          | Required | Description                                                        |
| --------------------------------- | -------- | ------------------------------------------------------------------ |
| `AZURE_AI_SEARCH_ENDPOINT`        | Yes      | Service URL (e.g. `https://my-search.search.windows.net`)          |
| `AZURE_AI_SEARCH_API_KEY`         | Yes      | Admin or query API key                                             |
| `AZURE_AI_SEARCH_INDEX_NAME`      | Yes      | Name of the search index                                           |
| `AZURE_AI_SEARCH_SEMANTIC_CONFIG` | No       | Semantic configuration name (required for `queryType: "semantic"`) |
| `AZURE_AI_SEARCH_VECTOR_FIELDS`   | No       | Comma-separated vector field names (default: `contentVector`)      |
| `AZURE_AI_SEARCH_SELECT_FIELDS`   | No       | Default fields to return (overridable per query)                   |

The tool is **only registered** when all three required variables are set. Without them, the server starts normally with no AI Search tool.

## Docker Compose Setup

Add the env vars to your `docker-compose.yml`:

```yaml
services:
  microsoft-graph-mcp:
    # ... existing config ...
    environment:
      # ... existing Azure/Graph vars ...

      # Azure AI Search (optional)
      - AZURE_AI_SEARCH_ENDPOINT=${AZURE_AI_SEARCH_ENDPOINT:-}
      - AZURE_AI_SEARCH_API_KEY=${AZURE_AI_SEARCH_API_KEY:-}
      - AZURE_AI_SEARCH_INDEX_NAME=${AZURE_AI_SEARCH_INDEX_NAME:-}
      - AZURE_AI_SEARCH_SEMANTIC_CONFIG=${AZURE_AI_SEARCH_SEMANTIC_CONFIG:-}
      - AZURE_AI_SEARCH_VECTOR_FIELDS=${AZURE_AI_SEARCH_VECTOR_FIELDS:-}
      - AZURE_AI_SEARCH_SELECT_FIELDS=${AZURE_AI_SEARCH_SELECT_FIELDS:-}
```

Then set the values in your `.env`:

```bash
AZURE_AI_SEARCH_ENDPOINT=https://civala-search.search.windows.net
AZURE_AI_SEARCH_API_KEY=your-query-key-here
AZURE_AI_SEARCH_INDEX_NAME=documents
AZURE_AI_SEARCH_SEMANTIC_CONFIG=my-semantic-config
AZURE_AI_SEARCH_VECTOR_FIELDS=contentVector
AZURE_AI_SEARCH_SELECT_FIELDS=title,content,url,lastModified
```

## Tool Usage

### Simple keyword search

```json
{
  "query": "quarterly revenue report",
  "queryType": "simple",
  "top": 10
}
```

### Semantic search (AI-ranked)

```json
{
  "query": "what was our revenue last quarter?",
  "queryType": "semantic",
  "top": 5
}
```

Returns results ranked by semantic relevance with extractive captions and answers.

### Hybrid search (text + vector)

```json
{
  "query": "clinical trial phase 3 results",
  "queryType": "semantic",
  "vectorSearch": true,
  "top": 10
}
```

Combines keyword matching with vector similarity for best recall.

### With filters

```json
{
  "query": "compliance",
  "queryType": "semantic",
  "filter": "category eq 'regulatory' and lastModified gt 2025-01-01T00:00:00Z",
  "select": "title,content,category,lastModified",
  "top": 20,
  "includeTotalCount": true
}
```

## Parameters Reference

| Parameter           | Type                           | Default      | Description                      |
| ------------------- | ------------------------------ | ------------ | -------------------------------- |
| `query`             | string                         | _(required)_ | Search text                      |
| `queryType`         | `simple` / `full` / `semantic` | `semantic`   | Query mode                       |
| `vectorSearch`      | boolean                        | `false`      | Enable hybrid vector search      |
| `filter`            | string                         | —            | OData filter expression          |
| `select`            | string                         | —            | Comma-separated fields to return |
| `top`               | number (1-50)                  | `10`         | Max results                      |
| `skip`              | number                         | —            | Pagination offset                |
| `orderby`           | string                         | —            | OData $orderby expression        |
| `highlightFields`   | string                         | —            | Fields to highlight              |
| `includeTotalCount` | boolean                        | `false`      | Include total matching count     |

## Setting Up a SharePoint Content Index

To index SharePoint content into Azure AI Search:

1. **Create a SharePoint Online data source** in your search service
2. **Create an indexer** that crawls the SharePoint site
3. **Add a skillset** (optional) for AI enrichment (OCR, key phrases, embeddings)
4. **Add a semantic configuration** to your index for semantic ranking

See: [Index SharePoint content](https://learn.microsoft.com/en-us/azure/search/search-howto-index-sharepoint-online)

## Verifying the Setup

1. Start the server and check the health endpoint:

   ```bash
   curl http://localhost:8080/health
   ```

2. List available tools — `azure_ai_search` should appear in the tool list

3. Test a query:
   ```bash
   # Via MCP client or direct tool call
   azure_ai_search({ "query": "test", "queryType": "simple", "top": 1 })
   ```

## Troubleshooting

| Issue                         | Cause                               | Fix                                                                |
| ----------------------------- | ----------------------------------- | ------------------------------------------------------------------ |
| Tool not appearing            | Missing required env vars           | Set all three: `ENDPOINT`, `API_KEY`, `INDEX_NAME`                 |
| 401 Unauthorized              | Invalid API key                     | Check key in Azure Portal → Search service → Keys                  |
| 404 Not Found                 | Wrong index name                    | Verify `AZURE_AI_SEARCH_INDEX_NAME` matches your index             |
| "Semantic search requires..." | Missing semantic config env var     | Set `AZURE_AI_SEARCH_SEMANTIC_CONFIG` or use `queryType: "simple"` |
| Empty vector results          | No vector fields / wrong field name | Check `AZURE_AI_SEARCH_VECTOR_FIELDS` matches your index schema    |
