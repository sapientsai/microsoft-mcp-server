import { execFileSync } from "node:child_process"
import { createRequire } from "node:module"

import type { AzureSession, OAuthSession } from "fastmcp"
import { AzureProvider, FastMCP } from "fastmcp"
import { type Either, Left, Option, Right } from "functype"
import { z } from "zod"

import type { TokenManager } from "./auth/token-manager.js"
import { createTokenManager } from "./auth/token-manager.js"
import type { AuthMode, ServerConfig } from "./auth/types.js"
import type { SiteCache } from "./cache/site-cache.js"
import { createSiteCache } from "./cache/site-cache.js"
import { filenameFromHeaders, filenameFromPath, formatBytes, processDownloadResponse } from "./download/download.js"
import { extractTextFromBuffer, resolveUploadContentType } from "./download/extract.js"
import { type AuthError, authError, type ConfigError, configError } from "./errors.js"
import { AZURE_BASE_URL, GRAPH_BASE_URL } from "./graph/client.js"
import type { AiSearchConfig } from "./search/ai-search-client.js"
import { buildAiSearchTool } from "./tools/ai-search.js"
import { buildSearchTool } from "./tools/sharepoint-search.js"
import {
  decodeBase64Upload,
  type DriveItemResponse,
  MAX_UPLOAD_SIZE,
  sessionUpload,
  SIMPLE_UPLOAD_LIMIT,
  simpleUpload,
} from "./upload/upload.js"

// Re-exports for public API
export type { TokenManager } from "./auth/token-manager.js"
export { createTokenManager } from "./auth/token-manager.js"
export type { AuthMode, ServerConfig } from "./auth/types.js"
export type { SiteCache, SiteInfo } from "./cache/site-cache.js"
export { createSiteCache } from "./cache/site-cache.js"
export {
  type DownloadResult,
  filenameFromHeaders,
  filenameFromPath,
  filenameFromUrl,
  formatBytes,
  isTextContent,
  processDownloadResponse,
  TEXT_MIME_PREFIXES,
  TEXT_MIME_SUFFIXES,
} from "./download/download.js"
export {
  CONTENT_TYPE_MAP,
  EXTRACTABLE_TYPES,
  extractTextFromBuffer,
  resolveContentType,
  resolveUploadContentType,
} from "./download/extract.js"
export type { AppError, AuthError, ConfigError, GraphError } from "./errors.js"
export { authError, configError, graphError } from "./errors.js"
export { AZURE_BASE_URL, GRAPH_BASE_URL, graphFetch, parseGraphError } from "./graph/client.js"
export type { AiSearchConfig } from "./search/ai-search-client.js"
export { AI_SEARCH_API_VERSION, aiSearchFetch, parseAiSearchError } from "./search/ai-search-client.js"
export type { AiSearchResult } from "./tools/ai-search.js"
export { buildAiSearchTool } from "./tools/ai-search.js"
export type { SearchResult } from "./tools/sharepoint-search.js"
export { buildSearchTool } from "./tools/sharepoint-search.js"
export type { DriveItemResponse } from "./upload/upload.js"
export { decodeBase64Upload, sessionUpload, simpleUpload } from "./upload/upload.js"

const require = createRequire(import.meta.url)
const { version: PKG_VERSION } = require("../package.json") as {
  version: `${number}.${number}.${number}`
}

function getGitHash(): string | undefined {
  try {
    return execFileSync("git", ["rev-parse", "--short", "HEAD"], { encoding: "utf-8" }).trim()
  } catch {
    const envHash = process.env.GIT_HASH
    return envHash ? envHash.slice(0, 7) : undefined
  }
}

const GIT_HASH = getGitHash()
const VERSION_STRING = GIT_HASH ? `v${PKG_VERSION}+${GIT_HASH}` : `v${PKG_VERSION}`

export const DEFAULT_CLIENT_ID = "cf7d1f97-781e-4034-930c-abd420e12d49"

function parseAuthMode(value: string | undefined): Either<ConfigError, AuthMode> {
  if (!value || value === "interactive") return Right("interactive" as AuthMode)
  if (value === "clientCredentials") return Right("clientCredentials" as AuthMode)
  return Left(configError(`Invalid AZURE_AUTH_MODE: "${value}". Must be "interactive" or "clientCredentials".`))
}

function validateConfig(config: Readonly<ServerConfig>): Either<ConfigError, void> {
  if (config.authMode === "clientCredentials") {
    if (config.tenantId === "common") {
      return Left(configError('Client credentials auth requires a specific tenant ID, not "common".'))
    }
    if (!config.clientSecret) {
      return Left(configError("Client credentials auth requires AZURE_CLIENT_SECRET."))
    }
  }
  return Right(undefined as void)
}

export function createConfig(): Readonly<ServerConfig> {
  const authMode = parseAuthMode(process.env.AZURE_AUTH_MODE).orThrow()

  const config: Readonly<ServerConfig> = {
    clientId: process.env.AZURE_CLIENT_ID ?? DEFAULT_CLIENT_ID,
    clientSecret: process.env.AZURE_CLIENT_SECRET ?? "",
    tenantId: process.env.AZURE_TENANT_ID ?? "common",
    baseUrl: process.env.BASE_URL ?? "http://localhost:8080",
    port: parseInt(process.env.PORT ?? "8080", 10),
    scopes: process.env.GRAPH_SCOPES?.split(",").map((s) => s.trim()) ?? [
      "openid",
      "profile",
      "email",
      "User.Read",
      "Mail.Read",
      "Calendars.Read",
      "Files.Read",
      "Sites.Read.All",
    ],
    authMode,
    appScopes: process.env.GRAPH_APP_SCOPES?.split(",").map((s) => s.trim()) ?? [
      "https://graph.microsoft.com/.default",
    ],
    apiKey: process.env.MCP_API_KEY ?? undefined,
    aiSearchEndpoint: process.env.AZURE_AI_SEARCH_ENDPOINT ?? undefined,
    aiSearchApiKey: process.env.AZURE_AI_SEARCH_API_KEY ?? undefined,
    aiSearchIndexName: process.env.AZURE_AI_SEARCH_INDEX_NAME ?? undefined,
    aiSearchSemanticConfig: process.env.AZURE_AI_SEARCH_SEMANTIC_CONFIG ?? undefined,
    aiSearchVectorFields: process.env.AZURE_AI_SEARCH_VECTOR_FIELDS ?? undefined,
    aiSearchSelectFields: process.env.AZURE_AI_SEARCH_SELECT_FIELDS ?? undefined,
    siteId: process.env.SITE_ID ?? undefined,
    searchRegion: process.env.GRAPH_SEARCH_REGION ?? undefined,
  }

  validateConfig(config).orThrow()

  return config
}

export function createServer(config: Readonly<ServerConfig>) {
  const isClientCredentials = config.authMode === "clientCredentials"

  const authProvider = isClientCredentials
    ? undefined
    : new AzureProvider({
        clientId: config.clientId,
        clientSecret: config.clientSecret,
        baseUrl: config.baseUrl,
        tenantId: config.tenantId,
        scopes: config.scopes,
      })

  const tokenManager: TokenManager | undefined = isClientCredentials ? createTokenManager(config) : undefined
  const siteCache: SiteCache | undefined = isClientCredentials ? createSiteCache() : undefined

  const baseInstructions = `Microsoft Graph MCP Server - Access Microsoft 365 data including users, mail, calendar, files, and more. Use sharepoint_search to find documents by keyword, then sharepoint_get_content to retrieve document text for analysis.

File Upload: For uploading files from the local environment to SharePoint/OneDrive, use the HTTP upload endpoint directly with curl:
  base64 {local_file} | tr -d '\\n' | curl -X POST -H "Authorization: Bearer {api_key}" -H "Content-Type: text/plain" --data-binary @- "{server_base_url}/upload?path={graph_path}&conflictBehavior=rename&encoding=base64"
The get_upload_config tool also returns ready-to-run curl commands with authentication.`
  const customInstructions = process.env.MCP_INSTRUCTIONS
  const instructions = customInstructions ? `${baseInstructions}\n\n${customInstructions}` : baseInstructions

  const server = new FastMCP({
    name: "microsoft-graph-server",
    version: PKG_VERSION,
    instructions,
    auth: authProvider,
    health: {
      enabled: true,
      path: "/health",
      message: `healthy ${VERSION_STRING}`,
      status: 200,
    },
    authenticate: config.apiKey
      ? (request) => {
          const authHeader = request.headers.authorization
          const headerKey = typeof authHeader === "string" ? authHeader.replace(/^Bearer\s+/i, "") : undefined

          const queryKey =
            !headerKey && request.url
              ? (() => {
                  try {
                    const url = new URL(request.url, "http://localhost")
                    return url.searchParams.get("api_key") ?? undefined
                  } catch {
                    return undefined
                  }
                })()
              : undefined

          const providedKey = headerKey ?? queryKey

          if (providedKey !== config.apiKey) {
            throw new Error("Unauthorized")
          }
          return Promise.resolve({})
        }
      : undefined,
  })

  const resolveAccessToken = async (session: unknown): Promise<Either<AuthError, string>> => {
    if (isClientCredentials && tokenManager) {
      return tokenManager.getToken()
    }
    const authSession = session as OAuthSession | undefined
    return Option(authSession?.accessToken).toEither(authError("Not authenticated. Please sign in first."))
  }

  const baseToolDescription =
    "Execute Microsoft Graph API requests. Use this to access Microsoft 365 data including users, mail, calendar, files, and more."
  const customInstructionsText = process.env.MCP_INSTRUCTIONS
  const toolDescription = customInstructionsText
    ? `${baseToolDescription} ${customInstructionsText}`
    : baseToolDescription

  server.addTool({
    name: "microsoft_graph",
    description: toolDescription,
    parameters: z.object({
      apiType: z
        .enum(["graph", "azure"])
        .default("graph")
        .describe('API type: "graph" for Microsoft Graph, "azure" for Azure Management'),
      path: z.string().describe("API endpoint path (e.g., /me, /users, /me/messages)"),
      method: z.enum(["GET", "POST", "PUT", "PATCH", "DELETE"]).default("GET").describe("HTTP method"),
      apiVersion: z.enum(["v1.0", "beta"]).default("v1.0").describe("Graph API version"),
      queryParams: z
        .record(z.string(), z.string())
        .optional()
        .describe("OData query parameters ($select, $filter, $top, $orderby, etc.)"),
      body: z
        .union([z.record(z.string(), z.unknown()), z.string()])
        .optional()
        .describe("Request body for POST/PUT/PATCH operations"),
    }),
    execute: async (args, { session, log }) => {
      const accessToken = (await resolveAccessToken(session)).orThrow()

      const baseUrl = args.apiType === "azure" ? AZURE_BASE_URL : GRAPH_BASE_URL
      const basePath = args.apiType === "azure" ? `${baseUrl}${args.path}` : `${baseUrl}/${args.apiVersion}${args.path}`
      const url =
        args.queryParams && Object.keys(args.queryParams).length > 0
          ? `${basePath}?${new URLSearchParams(args.queryParams).toString()}`
          : basePath

      log.info("Calling Microsoft Graph API", { url, method: args.method })

      const fetchOptions: Readonly<RequestInit> = {
        method: args.method,
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
        ...(args.body && ["POST", "PUT", "PATCH"].includes(args.method)
          ? { body: typeof args.body === "string" ? args.body : JSON.stringify(args.body) }
          : {}),
      }

      const response = await fetch(url, fetchOptions)

      const contentType = response.headers.get("content-type")
      const data: unknown = contentType?.includes("application/json") ? await response.json() : await response.text()

      if (!response.ok) {
        const errorMessage =
          typeof data === "object" && data !== null && "error" in data
            ? (data as { error: { message?: string } }).error.message
            : `HTTP ${response.status}: ${response.statusText}`
        throw new Error(errorMessage ?? `Request failed with status ${response.status}`)
      }

      return JSON.stringify(data, null, 2)
    },
  })

  server.addTool({
    name: "microsoft_graph_batch",
    description:
      "Execute multiple Microsoft Graph API requests in a single batch call (max 20). Use this for bulk operations like creating folder trees, sending multiple requests, or any scenario requiring many Graph API calls. Individual request failures don't fail the batch — each response has its own status code.",
    parameters: z.object({
      requests: z
        .array(
          z.object({
            id: z.string().describe("Unique ID to correlate with response"),
            method: z.enum(["GET", "POST", "PUT", "PATCH", "DELETE"]).describe("HTTP method"),
            url: z.string().describe("Relative API path (e.g., /me/drive/root/children)"),
            headers: z
              .record(z.string(), z.string())
              .optional()
              .describe("Request headers. Content-Type is auto-added for requests with a body"),
            body: z
              .union([z.record(z.string(), z.unknown()), z.string()])
              .optional()
              .describe("Request body for POST/PUT/PATCH"),
            dependsOn: z.array(z.string()).optional().describe("IDs of requests that must complete before this one"),
          }),
        )
        .min(1)
        .max(20)
        .describe("Array of requests (max 20)"),
      apiVersion: z.enum(["v1.0", "beta"]).default("v1.0").describe("Graph API version"),
    }),
    execute: async (args, { session, log }) => {
      const accessToken = (await resolveAccessToken(session)).orThrow()

      const batchRequests = args.requests.map((req) => {
        const normalized: Record<string, unknown> = {
          id: req.id,
          method: req.method,
          url: req.url,
        }

        if (req.body !== undefined) {
          normalized.body = typeof req.body === "string" ? JSON.parse(req.body) : req.body

          const headers = req.headers ? { ...req.headers } : {}
          if (!Object.keys(headers).some((k) => k.toLowerCase() === "content-type")) {
            headers["Content-Type"] = "application/json"
          }
          normalized.headers = headers
        } else if (req.headers) {
          normalized.headers = req.headers
        }

        if (req.dependsOn && req.dependsOn.length > 0) {
          normalized.dependsOn = req.dependsOn
        }

        return normalized
      })

      const url = `${GRAPH_BASE_URL}/${args.apiVersion}/$batch`
      log.info("Calling Microsoft Graph Batch API", { url, requestCount: batchRequests.length })

      const response = await fetch(url, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ requests: batchRequests }),
      })

      const data: unknown = await response.json()

      if (!response.ok) {
        const errorMessage =
          typeof data === "object" && data !== null && "error" in data
            ? (data as { error: { message?: string } }).error.message
            : `HTTP ${response.status}: ${response.statusText}`
        throw new Error(errorMessage ?? `Batch request failed with status ${response.status}`)
      }

      return JSON.stringify(data, null, 2)
    },
  })

  server.addTool({
    name: "get_auth_status",
    description: "Check the current authentication status.",
    parameters: z.object({}),
    execute: async (_args, { session }) => {
      if (isClientCredentials && tokenManager) {
        const result = await tokenManager.getSession()
        return result.fold(
          (err) =>
            JSON.stringify(
              {
                authenticated: false,
                mode: "clientCredentials",
                message: `Authentication failed: ${err.message}`,
              },
              null,
              2,
            ),
          (appSession) =>
            JSON.stringify(
              {
                authenticated: true,
                mode: "clientCredentials",
                message: "Authenticated via client credentials (app-only)",
                expiresAt: appSession.expiresAt.toISOString(),
              },
              null,
              2,
            ),
        )
      }

      const authSession = session as AzureSession | undefined
      const hasAuth = !!authSession?.accessToken
      return JSON.stringify(
        {
          authenticated: hasAuth,
          mode: "interactive",
          message: hasAuth ? "Authenticated via Azure OAuth" : "Not authenticated. Please sign in.",
          scopes: authSession?.scopes,
          upn: authSession?.upn,
        },
        null,
        2,
      )
    },
  })

  server.addTool({
    name: "download_file",
    description:
      "Download a file from SharePoint or OneDrive via Microsoft Graph API. Returns file content directly to the agent: images are returned inline, text files as text content, and binary files (Office docs, PDFs) are saved to disk. Supports optional format conversion (e.g., PDF).",
    parameters: z.object({
      path: z
        .string()
        .describe(
          "Graph API path to the file content endpoint (e.g., /me/drive/items/{id}/content, /sites/{siteId}/drive/items/{id}/content, /me/drive/root:/Documents/report.pdf:/content)",
        ),
      apiVersion: z.enum(["v1.0", "beta"]).default("v1.0").describe("Graph API version"),
      format: z
        .string()
        .optional()
        .describe("Optional conversion format (e.g., 'pdf'). Only supported for certain file types."),
      outputDir: z.string().optional().describe("Directory to save the file. Defaults to system temp directory."),
      filename: z
        .string()
        .optional()
        .describe("Override filename. If not provided, uses the filename from the response headers or URL."),
    }),
    execute: async (args, { session, log }) => {
      const accessToken = (await resolveAccessToken(session)).orThrow()

      const queryParams = args.format ? `?format=${args.format}` : ""
      const url = `${GRAPH_BASE_URL}/${args.apiVersion}${args.path}${queryParams}`

      log.info("Downloading file from Microsoft Graph", { url })

      const response = await fetch(url, {
        method: "GET",
        headers: { Authorization: `Bearer ${accessToken}` },
      })

      if (!response.ok) {
        const responseContentType = response.headers.get("content-type")
        if (responseContentType?.includes("application/json")) {
          const errorData = (await response.json()) as { error?: { message?: string } }
          throw new Error(errorData.error?.message ?? `HTTP ${response.status}: ${response.statusText}`)
        }
        throw new Error(`HTTP ${response.status}: ${response.statusText}`)
      }

      const contentType = response.headers.get("content-type") ?? "application/octet-stream"
      const resolvedFilename = Option(args.filename)
        .or(filenameFromHeaders(response.headers))
        .or(filenameFromPath(args.path))
        .orElse("download")
      const buffer = Buffer.from(await response.arrayBuffer())

      return processDownloadResponse(buffer, contentType, resolvedFilename, args.outputDir)
    },
  })

  server.addTool({
    name: "get_upload_config",
    description:
      "Get the authenticated upload endpoint URL and curl command for uploading files to SharePoint/OneDrive. Call this tool first, then execute the returned curl command (POST with --data-binary) to upload. This bypasses MCP protocol limits and handles files up to 250MB.",
    parameters: z.object({
      path: z
        .string()
        .describe(
          "Graph API destination path ending with :/content (e.g., /drives/{driveId}/root:/folder/file.docx:/content, /sites/{siteId}/drive/root:/folder/file.xlsx:/content)",
        ),
      localFile: z
        .string()
        .optional()
        .describe("Local file path to include in the curl example. If omitted, a placeholder is used."),
      contentType: z.string().optional().describe("MIME type override. Auto-detected from file extension if omitted."),
      conflictBehavior: z
        .enum(["rename", "replace", "fail"])
        .default("rename")
        .describe('Conflict behavior: "rename" (default) adds a suffix, "replace" overwrites, "fail" returns an error'),
    }),
    // eslint-disable-next-line @typescript-eslint/require-await
    execute: async (args) => {
      const filename = filenameFromPath(args.path).orElse("upload")
      const contentType = resolveUploadContentType(args.contentType, filename)
      const localFile = args.localFile ?? "{local_file_path}"

      const params = new URLSearchParams({
        path: args.path,
        conflictBehavior: args.conflictBehavior,
        contentType,
        encoding: "base64",
      })
      const uploadUrl = `${config.baseUrl}/upload?${params.toString()}`

      const authHeader = config.apiKey ? `Authorization: Bearer ${config.apiKey}` : undefined

      const curlParts = [
        `base64 "${localFile}" | tr -d '\\n'`,
        "| curl -X POST",
        authHeader ? `-H "${authHeader}"` : undefined,
        `-H "Content-Type: text/plain"`,
        `--data-binary @-`,
        `"${uploadUrl}"`,
      ]
        .filter(Boolean)
        .join(" \\\n  ")

      return JSON.stringify(
        {
          uploadUrl,
          method: "PUT",
          contentType,
          ...(authHeader ? { authHeader } : {}),
          curl: curlParts,
        },
        null,
        2,
      )
    },
  })

  server.addTool({
    name: "read_document",
    description:
      "Download a file from SharePoint or OneDrive and return its readable text content. Supports DOCX, PDF, XLSX, and text-based files. Use this instead of download_file when you need to read document contents. Use with sharepoint_search results by constructing the path: /drives/{driveId}/items/{itemId}/content",
    parameters: z.object({
      path: z
        .string()
        .describe(
          "Graph API path to the file content endpoint (e.g., /me/drive/items/{id}/content, /sites/{siteId}/drive/items/{id}/content, /drives/{driveId}/items/{itemId}/content, /me/drive/root:/Documents/report.pdf:/content)",
        ),
      apiVersion: z.enum(["v1.0", "beta"]).default("v1.0").describe("Graph API version"),
      format: z
        .string()
        .optional()
        .describe("Optional conversion format (e.g., 'pdf'). Only supported for certain file types."),
      maxChars: z
        .number()
        .min(1000)
        .max(200000)
        .default(50000)
        .optional()
        .describe("Maximum characters to return (1000-200000). Content beyond this limit is truncated."),
    }),
    execute: async (args, { session, log }) => {
      const accessToken = (await resolveAccessToken(session)).orThrow()

      const queryParams = args.format ? `?format=${args.format}` : ""
      const url = `${GRAPH_BASE_URL}/${args.apiVersion}${args.path}${queryParams}`

      log.info("Downloading file for text extraction", { url })

      const response = await fetch(url, {
        method: "GET",
        headers: { Authorization: `Bearer ${accessToken}` },
      })

      if (!response.ok) {
        const responseContentType = response.headers.get("content-type")
        if (responseContentType?.includes("application/json")) {
          const errorData = (await response.json()) as { error?: { message?: string } }
          throw new Error(errorData.error?.message ?? `HTTP ${response.status}: ${response.statusText}`)
        }
        throw new Error(`HTTP ${response.status}: ${response.statusText}`)
      }

      const contentType = response.headers.get("content-type") ?? "application/octet-stream"
      const resolvedFilename = Option<string>(undefined)
        .or(filenameFromHeaders(response.headers))
        .or(filenameFromPath(args.path))
        .orElse("download")
      const buffer = Buffer.from(await response.arrayBuffer())

      const maxFileSize = 10 * 1024 * 1024 // 10 MB
      if (buffer.length > maxFileSize) {
        throw new Error(
          `File too large (${formatBytes(buffer.length)}). Maximum supported size is ${formatBytes(maxFileSize)}.`,
        )
      }

      const fullText = (await extractTextFromBuffer(buffer, contentType, resolvedFilename)).orThrow()

      const maxChars = args.maxChars ?? 50000
      const text =
        fullText.length > maxChars
          ? `${fullText.slice(0, maxChars)}\n\n[truncated at ${maxChars.toLocaleString()} chars — full document is ${fullText.length.toLocaleString()} chars]`
          : fullText

      return {
        content: [
          {
            type: "text" as const,
            text: `File: ${resolvedFilename} (${formatBytes(buffer.length)})\n\n${text}`,
          },
        ],
      }
    },
  })

  // SharePoint search tool
  server.addTool(
    buildSearchTool(
      async (session) => (await resolveAccessToken(session)).orThrow(),
      config.authMode,
      siteCache,
      config.siteId,
      config.authMode === "clientCredentials" ? (config.searchRegion ?? "NAM") : config.searchRegion,
    ),
  )

  // Azure AI Search tool (optional — only registered when configured)
  if (config.aiSearchEndpoint && config.aiSearchApiKey && config.aiSearchIndexName) {
    const aiSearchConfig: AiSearchConfig = {
      endpoint: config.aiSearchEndpoint,
      apiKey: config.aiSearchApiKey,
      indexName: config.aiSearchIndexName,
      semanticConfiguration: config.aiSearchSemanticConfig,
      vectorFields: config.aiSearchVectorFields,
      selectFields: config.aiSearchSelectFields,
    }
    server.addTool(buildAiSearchTool(aiSearchConfig))
  }

  // HTTP upload endpoint
  const app = server.getApp()

  const handleUpload = async (req: {
    header: (name: string) => string | undefined
    query: (name: string) => string | undefined
    arrayBuffer: () => Promise<ArrayBuffer>
  }) => {
    // Auth: require API key if configured
    if (config.apiKey) {
      const authHeader = req.header("authorization")
      const token = authHeader?.replace(/^Bearer\s+/i, "")
      if (token !== config.apiKey) {
        return { status: 401 as const, body: { error: "Unauthorized" } }
      }
    }

    // Params
    const path = req.query("path")
    if (!path) return { status: 400 as const, body: { error: "path query parameter is required" } }
    const apiVersion = req.query("apiVersion") ?? "v1.0"
    const conflictBehavior = req.query("conflictBehavior") ?? "rename"
    const explicitContentType = req.query("contentType")

    // Read body
    const encoding = req.query("encoding")
    const arrayBuffer = await req.arrayBuffer()
    const rawBuffer = Buffer.from(arrayBuffer)
    if (rawBuffer.length === 0) return { status: 400 as const, body: { error: "Empty request body" } }
    const buffer = encoding === "base64" ? decodeBase64Upload(rawBuffer) : rawBuffer
    if (buffer.length === 0) return { status: 400 as const, body: { error: "Invalid base64 content" } }
    if (buffer.length > MAX_UPLOAD_SIZE) return { status: 413 as const, body: { error: "File too large (max 250MB)" } }

    // Resolve
    const accessToken = (await resolveAccessToken(undefined)).orThrow()
    const filename = filenameFromPath(path).orElse("upload")
    const contentType = resolveUploadContentType(explicitContentType, filename)
    const apiBase = `${GRAPH_BASE_URL}/${apiVersion}`

    const uploadResult: Either<{ message: string }, DriveItemResponse> =
      buffer.length <= SIMPLE_UPLOAD_LIMIT
        ? await simpleUpload(apiBase, path, accessToken, buffer, contentType, conflictBehavior)
        : await sessionUpload(apiBase, path, accessToken, buffer, conflictBehavior)

    const driveItem = uploadResult.orThrow()
    return { status: 200 as const, body: driveItem }
  }

  app.post("/upload", async (c) => {
    try {
      const result = await handleUpload(c.req)
      return c.json(result.body, result.status)
    } catch (err) {
      const message = err instanceof Error ? err.message : "Unknown error"
      return c.json({ error: message }, 500)
    }
  })

  app.put("/upload", async (c) => {
    try {
      const result = await handleUpload(c.req)
      return c.json(result.body, result.status)
    } catch (err) {
      const message = err instanceof Error ? err.message : "Unknown error"
      return c.json({ error: message }, 500)
    }
  })

  return { server, authProvider, tokenManager, siteCache, config }
}

export async function runServer(): Promise<void> {
  const config = createConfig()
  const { server } = createServer(config)

  const transportType = process.env.TRANSPORT_TYPE ?? "httpStream"

  if (transportType === "stdio") {
    await server.start({
      transportType: "stdio",
    })
    console.error("Microsoft Graph MCP Server running on stdio")
  } else {
    await server.start({
      transportType: "httpStream",
      httpStream: {
        port: config.port,
        host: "0.0.0.0",
      },
    })
    console.error(`Microsoft Graph MCP Server running on http://localhost:${config.port}`)
    if (config.authMode === "interactive") {
      console.error(`OAuth callback URL: ${config.baseUrl}/oauth/callback`)
    } else {
      console.error(`Auth mode: client credentials (app-only)`)
    }
  }
}
