import { mkdir, readFile, writeFile } from "node:fs/promises"
import { basename, extname, join } from "node:path"

import ExcelJS from "exceljs"
import type { AzureSession, OAuthSession } from "fastmcp"
import { AzureProvider, FastMCP, imageContent } from "fastmcp"
import mammoth from "mammoth"
import { extractText as extractPdfText, getDocumentProxy } from "unpdf"
import { z } from "zod"

import type { TokenManager } from "./auth/token-manager.js"
import { createTokenManager } from "./auth/token-manager.js"
import type { AuthMode, ServerConfig } from "./auth/types.js"
import type { SiteCache } from "./cache/site-cache.js"
import { createSiteCache } from "./cache/site-cache.js"
import { buildSearchTool } from "./tools/sharepoint-search.js"

export type { TokenManager } from "./auth/token-manager.js"
export { createTokenManager } from "./auth/token-manager.js"
export type { AuthMode, ServerConfig } from "./auth/types.js"
export type { SiteCache, SiteInfo } from "./cache/site-cache.js"
export { createSiteCache } from "./cache/site-cache.js"
export type { SearchResult } from "./tools/sharepoint-search.js"
export { buildSearchTool } from "./tools/sharepoint-search.js"

export const DEFAULT_CLIENT_ID = "cf7d1f97-781e-4034-930c-abd420e12d49"
export const GRAPH_BASE_URL = "https://graph.microsoft.com"
export const AZURE_BASE_URL = "https://management.azure.com"

function parseAuthMode(value: string | undefined): AuthMode {
  if (!value || value === "interactive") {
    return "interactive"
  }
  if (value === "clientCredentials") {
    return "clientCredentials"
  }
  throw new Error(`Invalid AZURE_AUTH_MODE: "${value}". Must be "interactive" or "clientCredentials".`)
}

function validateConfig(config: Readonly<ServerConfig>): void {
  if (config.authMode === "clientCredentials") {
    if (config.tenantId === "common") {
      throw new Error('Client credentials auth requires a specific tenant ID, not "common".')
    }
    if (!config.clientSecret) {
      throw new Error("Client credentials auth requires AZURE_CLIENT_SECRET.")
    }
  }
}

export function createConfig(): Readonly<ServerConfig> {
  const authMode = parseAuthMode(process.env.AZURE_AUTH_MODE)

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
  }

  validateConfig(config)

  return config
}

export function createServer(config: Readonly<ServerConfig>) {
  const isClientCredentials = config.authMode === "clientCredentials"

  // Only create AzureProvider for interactive mode
  const authProvider = isClientCredentials
    ? undefined
    : new AzureProvider({
        clientId: config.clientId,
        clientSecret: config.clientSecret,
        baseUrl: config.baseUrl,
        tenantId: config.tenantId,
        scopes: config.scopes,
      })

  // Create token manager for client credentials mode
  const tokenManager: TokenManager | undefined = isClientCredentials ? createTokenManager(config) : undefined

  // Create site cache for client credentials mode (used by SharePoint search fan-out)
  const siteCache: SiteCache | undefined = isClientCredentials ? createSiteCache() : undefined

  const baseInstructions = `Microsoft Graph MCP Server - Access Microsoft 365 data including users, mail, calendar, files, and more. Use sharepoint_search to find documents by keyword, then sharepoint_get_content to retrieve document text for analysis.

File Upload: For uploading files from the local environment to SharePoint/OneDrive, use the HTTP upload endpoint directly with curl:
  curl -X PUT -H "Authorization: Bearer {api_key}" -H "Content-Type: {mime}" --data-binary @{local_file} "{server_base_url}/upload?path={graph_path}&conflictBehavior=rename"
The upload_file tool also supports localPath (stdio mode), sourceUrl (fetch from URL), and content (small base64 files).`
  const customInstructions = process.env.MCP_INSTRUCTIONS
  const instructions = customInstructions ? `${baseInstructions}\n\n${customInstructions}` : baseInstructions

  const server = new FastMCP({
    name: "microsoft-graph-server",
    version: "1.0.0",
    instructions,
    auth: authProvider,
    health: {
      enabled: true,
      path: "/health",
      message: "healthy",
      status: 200,
    },
    authenticate: config.apiKey
      ? (request) => {
          // Check Authorization header first
          const authHeader = request.headers.authorization
          const headerKey = typeof authHeader === "string" ? authHeader.replace(/^Bearer\s+/i, "") : undefined

          // Fall back to api_key query parameter if no header
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

  const resolveAccessToken = async (session: unknown): Promise<string> => {
    if (isClientCredentials && tokenManager) {
      return tokenManager.getToken()
    }
    const authSession = session as OAuthSession | undefined
    if (!authSession?.accessToken) {
      throw new Error("Not authenticated. Please sign in first.")
    }
    return authSession.accessToken
  }

  const baseToolDescription =
    "Execute Microsoft Graph API requests. Use this to access Microsoft 365 data including users, mail, calendar, files, and more."
  const toolDescription = customInstructions ? `${baseToolDescription} ${customInstructions}` : baseToolDescription

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
      const accessToken = await resolveAccessToken(session)

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
      const accessToken = await resolveAccessToken(session)

      const batchRequests = args.requests.map((req) => {
        const normalized: Record<string, unknown> = {
          id: req.id,
          method: req.method,
          url: req.url,
        }

        if (req.body !== undefined) {
          // Parse string bodies into objects — Graph batch expects JSON objects, not stringified JSON
          normalized.body = typeof req.body === "string" ? JSON.parse(req.body) : req.body

          // Auto-add Content-Type header for requests with a body
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
        // Client credentials mode - check token manager
        try {
          const appSession = await tokenManager.getSession()
          return JSON.stringify(
            {
              authenticated: true,
              mode: "clientCredentials",
              message: "Authenticated via client credentials (app-only)",
              expiresAt: appSession.expiresAt.toISOString(),
            },
            null,
            2,
          )
        } catch (error) {
          return JSON.stringify(
            {
              authenticated: false,
              mode: "clientCredentials",
              message: `Authentication failed: ${error instanceof Error ? error.message : "Unknown error"}`,
            },
            null,
            2,
          )
        }
      }

      // Interactive mode
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
      const accessToken = await resolveAccessToken(session)

      const queryParams = args.format ? `?format=${args.format}` : ""
      const url = `${GRAPH_BASE_URL}/${args.apiVersion}${args.path}${queryParams}`

      log.info("Downloading file from Microsoft Graph", { url })

      const response = await fetch(url, {
        method: "GET",
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
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
      const resolvedFilename =
        args.filename ?? filenameFromHeaders(response.headers) ?? filenameFromPath(args.path) ?? "download"
      const buffer = Buffer.from(await response.arrayBuffer())

      return processDownloadResponse(buffer, contentType, resolvedFilename, args.outputDir)
    },
  })

  server.addTool({
    name: "upload_file",
    description:
      "Upload a file to SharePoint or OneDrive via Microsoft Graph API. Supports simple upload (≤4MB) and chunked sessions (4MB–250MB). Requires Files.ReadWrite or Files.ReadWrite.All scope. Input modes: localPath (stdio/Desktop), sourceUrl (server-side fetch from URL), content (base64, small files only). For large files in cloud/httpStream mode, prefer the HTTP endpoint: PUT /upload?path={graphPath} with --data-binary.",
    parameters: z.object({
      path: z
        .string()
        .describe(
          "Graph API destination path ending with :/content (e.g., /drives/{driveId}/root:/folder/file.docx:/content, /me/drive/root:/Documents/report.pdf:/content, /sites/{siteId}/drive/root:/folder/file.xlsx:/content)",
        ),
      apiVersion: z.enum(["v1.0", "beta"]).default("v1.0").describe("Graph API version"),
      localPath: z
        .string()
        .optional()
        .describe(
          "Absolute local file path. Works in stdio/Desktop mode. Mutually exclusive with content and sourceUrl.",
        ),
      content: z
        .string()
        .optional()
        .describe(
          "Base64-encoded file content. For small files only — large base64 strings get truncated in cloud mode. Mutually exclusive with localPath and sourceUrl.",
        ),
      sourceUrl: z
        .string()
        .optional()
        .describe(
          "URL to fetch file content from server-side. Bypasses base64 limits. Mutually exclusive with localPath and content.",
        ),
      contentType: z.string().optional().describe("MIME type override. Auto-detected from file extension if omitted."),
      conflictBehavior: z
        .enum(["rename", "replace", "fail"])
        .default("rename")
        .describe('Conflict behavior: "rename" (default) adds a suffix, "replace" overwrites, "fail" returns an error'),
    }),
    execute: async (args, { session, log }) => {
      const accessToken = await resolveAccessToken(session)

      // Validate exactly one source
      const sources = [args.localPath, args.content, args.sourceUrl].filter(Boolean)
      if (sources.length !== 1) {
        throw new Error("Provide exactly one of: localPath, content (base64), or sourceUrl.")
      }

      // Read file data
      const readLocalFile = async (localPath: string): Promise<Buffer> => {
        try {
          return await readFile(localPath)
        } catch (err) {
          if ((err as NodeJS.ErrnoException).code === "ENOENT") {
            throw new Error(
              `File not found: ${localPath}. If this file is on a remote client, ` +
                `the MCP server cannot access it directly. Upload via HTTP instead: ` +
                `curl -X PUT --data-binary @"${localPath}" "${config.baseUrl}/upload?path=${encodeURIComponent(args.path)}&conflictBehavior=${args.conflictBehavior}"`,
              { cause: err },
            )
          }
          throw err
        }
      }

      const buffer = args.localPath
        ? await readLocalFile(args.localPath)
        : args.content
          ? Buffer.from(args.content, "base64")
          : await fetchUrlToBuffer(args.sourceUrl!)

      // Resolve filename
      const filename =
        filenameFromPath(args.path) ??
        (args.localPath ? basename(args.localPath) : undefined) ??
        (args.sourceUrl ? filenameFromUrl(args.sourceUrl) : undefined) ??
        "upload"

      // Resolve content type
      const contentType = resolveUploadContentType(args.contentType, filename)

      // Size guard
      if (buffer.length > MAX_UPLOAD_SIZE) {
        throw new Error(
          `File too large (${formatBytes(buffer.length)}). Maximum upload size is ${formatBytes(MAX_UPLOAD_SIZE)}.`,
        )
      }

      const apiBase = `${GRAPH_BASE_URL}/${args.apiVersion}`

      log.info("Uploading file to Microsoft Graph", {
        path: args.path,
        filename,
        size: buffer.length,
        method: buffer.length <= SIMPLE_UPLOAD_LIMIT ? "simple" : "session",
      })

      const driveItem =
        buffer.length <= SIMPLE_UPLOAD_LIMIT
          ? await simpleUpload(apiBase, args.path, accessToken, buffer, contentType, args.conflictBehavior)
          : await sessionUpload(apiBase, args.path, accessToken, buffer, args.conflictBehavior)

      return JSON.stringify(
        {
          id: driveItem.id,
          name: driveItem.name,
          size: driveItem.size,
          webUrl: driveItem.webUrl,
          createdDateTime: driveItem.createdDateTime,
          lastModifiedDateTime: driveItem.lastModifiedDateTime,
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
      const accessToken = await resolveAccessToken(session)

      const queryParams = args.format ? `?format=${args.format}` : ""
      const url = `${GRAPH_BASE_URL}/${args.apiVersion}${args.path}${queryParams}`

      log.info("Downloading file for text extraction", { url })

      const response = await fetch(url, {
        method: "GET",
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
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
      const resolvedFilename = filenameFromHeaders(response.headers) ?? filenameFromPath(args.path) ?? "download"
      const buffer = Buffer.from(await response.arrayBuffer())

      const maxFileSize = 10 * 1024 * 1024 // 10 MB
      if (buffer.length > maxFileSize) {
        throw new Error(
          `File too large (${formatBytes(buffer.length)}). Maximum supported size is ${formatBytes(maxFileSize)}.`,
        )
      }

      const fullText = await extractTextFromBuffer(buffer, contentType, resolvedFilename)

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
  server.addTool(buildSearchTool(resolveAccessToken, config.authMode, siteCache))

  // HTTP upload endpoint — bypasses MCP protocol for binary file uploads
  // Used by Claude Container: curl --data-binary @file "https://server/upload?path=..."
  const app = server.getApp()

  app.put("/upload", async (c) => {
    // Auth: require API key if configured
    if (config.apiKey) {
      const authHeader = c.req.header("authorization")
      const token = authHeader?.replace(/^Bearer\s+/i, "")
      if (token !== config.apiKey) {
        return c.json({ error: "Unauthorized" }, 401)
      }
    }

    // Params
    const path = c.req.query("path")
    if (!path) return c.json({ error: "path query parameter is required" }, 400)
    const apiVersion = c.req.query("apiVersion") ?? "v1.0"
    const conflictBehavior = c.req.query("conflictBehavior") ?? "rename"
    const explicitContentType = c.req.query("contentType")

    // Read binary body
    const arrayBuffer = await c.req.arrayBuffer()
    const buffer = Buffer.from(arrayBuffer)
    if (buffer.length === 0) return c.json({ error: "Empty request body" }, 400)
    if (buffer.length > MAX_UPLOAD_SIZE) return c.json({ error: "File too large (max 250MB)" }, 413)

    // Resolve
    const accessToken = await resolveAccessToken(undefined)
    const filename = filenameFromPath(path) ?? "upload"
    const contentType = resolveUploadContentType(explicitContentType, filename)
    const apiBase = `${GRAPH_BASE_URL}/${apiVersion}`

    const driveItem =
      buffer.length <= SIMPLE_UPLOAD_LIMIT
        ? await simpleUpload(apiBase, path, accessToken, buffer, contentType, conflictBehavior)
        : await sessionUpload(apiBase, path, accessToken, buffer, conflictBehavior)

    return c.json(driveItem)
  })

  return { server, authProvider, tokenManager, siteCache, config }
}

export const TEXT_MIME_PREFIXES = ["text/", "application/json", "application/xml", "application/csv"]
export const TEXT_MIME_SUFFIXES = ["+xml", "+json"]

export function isTextContent(contentType: string): boolean {
  const lower = contentType.toLowerCase()
  return (
    TEXT_MIME_PREFIXES.some((prefix) => lower.startsWith(prefix)) ||
    TEXT_MIME_SUFFIXES.some((suffix) => lower.includes(suffix))
  )
}

export function filenameFromHeaders(headers: Headers): string | undefined {
  const disposition = headers.get("content-disposition")
  if (!disposition) return undefined
  const match = /filename\*?=(?:UTF-8''|"?)([^";]+)"?/i.exec(disposition)
  return match?.[1] ? decodeURIComponent(match[1]) : undefined
}

export function filenameFromPath(path: string): string | undefined {
  // Handle paths like /me/drive/root:/Documents/report.pdf:/content
  const colonPathMatch = /:\/([^:]+):\/content/i.exec(path)
  if (colonPathMatch?.[1]) {
    return basename(colonPathMatch[1])
  }
  return undefined
}

export function filenameFromUrl(url: string): string | undefined {
  try {
    const { pathname } = new URL(url)
    const name = basename(pathname)
    return name && name !== "/" ? decodeURIComponent(name) : undefined
  } catch {
    return undefined
  }
}

export function formatBytes(bytes: number): string {
  if (bytes === 0) return "0 B"
  const units = ["B", "KB", "MB", "GB"]
  const i = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1)
  const value = bytes / Math.pow(1024, i)
  return `${value.toFixed(i === 0 ? 0 : 1)} ${units[i]}`
}

export type DownloadResult = {
  content: Array<{ type: "text"; text: string } | { type: "image"; data: string; mimeType: string }>
}

export async function processDownloadResponse(
  buffer: Buffer,
  contentType: string,
  filename: string,
  outputDir?: string,
): Promise<DownloadResult> {
  // Images: return inline via MCP image content so the LLM can see them
  if (contentType.startsWith("image/")) {
    const img = await imageContent({ buffer })
    return {
      content: [{ type: "text" as const, text: `Image: ${filename} (${formatBytes(buffer.length)})` }, img],
    }
  }

  // Text-based files: return content inline so the LLM can read them
  if (isTextContent(contentType)) {
    const text = buffer.toString("utf-8")
    return {
      content: [
        {
          type: "text" as const,
          text: `File: ${filename} (${formatBytes(buffer.length)}, ${contentType})\n\n${text}`,
        },
      ],
    }
  }

  // Binary files (Office docs, PDFs, etc.): return base64 content + optionally save to disk
  const base64Data = buffer.toString("base64")

  // Save to disk when outputDir is explicitly provided (useful for stdio/local mode)
  const savedPath = outputDir
    ? await (async () => {
        await mkdir(outputDir, { recursive: true })
        const outputPath = join(outputDir, filename)
        await writeFile(outputPath, buffer)
        return outputPath
      })()
    : undefined

  return {
    content: [
      {
        type: "text" as const,
        text: JSON.stringify(
          {
            filename,
            contentType,
            size: buffer.length,
            sizeFormatted: formatBytes(buffer.length),
            encoding: "base64",
            data: base64Data,
            ...(savedPath ? { savedTo: savedPath } : {}),
          },
          null,
          2,
        ),
      },
    ],
  }
}

const CONTENT_TYPE_MAP: Record<string, string> = {
  ".pdf": "application/pdf",
  ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  ".doc": "application/msword",
  ".xls": "application/vnd.ms-excel",
  ".ppt": "application/vnd.ms-powerpoint",
  ".txt": "text/plain",
  ".csv": "text/csv",
  ".json": "application/json",
  ".xml": "application/xml",
  ".html": "text/html",
  ".htm": "text/html",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".gif": "image/gif",
  ".svg": "image/svg+xml",
  ".zip": "application/zip",
  ".mp4": "video/mp4",
  ".mp3": "audio/mpeg",
}

export const EXTRACTABLE_TYPES = [
  "application/pdf",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
] as const

function resolveContentType(contentType: string, filename: string): string {
  const lower = contentType.toLowerCase()
  // If the content type is generic, try to infer from file extension
  if (lower === "application/octet-stream" || lower === "") {
    const ext = extname(filename).toLowerCase()
    return CONTENT_TYPE_MAP[ext] ?? contentType
  }
  return lower
}

export async function extractTextFromBuffer(buffer: Buffer, contentType: string, filename: string): Promise<string> {
  const resolved = resolveContentType(contentType, filename)

  // Text-based types: return directly
  if (isTextContent(resolved)) {
    return buffer.toString("utf-8")
  }

  // PDF
  if (resolved === "application/pdf") {
    const pdf = await getDocumentProxy(new Uint8Array(buffer))
    try {
      const { totalPages, text } = await extractPdfText(pdf, { mergePages: true })
      return `[PDF: ${totalPages} page${totalPages === 1 ? "" : "s"}]\n\n${text}`
    } finally {
      await pdf.destroy()
    }
  }

  // DOCX
  if (resolved === "application/vnd.openxmlformats-officedocument.wordprocessingml.document") {
    const result = await mammoth.extractRawText({ buffer })
    return result.value
  }

  // XLSX
  if (resolved === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") {
    const wb = new ExcelJS.Workbook()
    await wb.xlsx.load(buffer as unknown as ArrayBuffer)
    const parts: string[] = []
    wb.eachSheet((ws) => {
      const rows: string[] = []
      ws.eachRow((row) => {
        const cells = Array.isArray(row.values) ? row.values.slice(1) : []
        rows.push(cells.map((v) => (v == null ? "" : String(v))).join(","))
      })
      const csv = rows.join("\n")
      if (wb.worksheets.length > 1) {
        parts.push(`[Sheet: ${ws.name}]\n${csv}`)
      } else {
        parts.push(csv)
      }
    })
    return parts.join("\n\n")
  }

  const supported = [...EXTRACTABLE_TYPES, "text/*"].join(", ")
  throw new Error(`Unsupported content type "${contentType}" for text extraction. Supported: ${supported}`)
}

// --- Upload helpers ---

export type DriveItemResponse = {
  id: string
  name: string
  size: number
  webUrl: string
  createdDateTime?: string
  lastModifiedDateTime?: string
  file?: { mimeType: string }
  parentReference?: { driveId: string; path: string }
}

export function resolveUploadContentType(explicit: string | undefined, filename: string): string {
  if (explicit) return explicit
  const ext = extname(filename).toLowerCase()
  return CONTENT_TYPE_MAP[ext] ?? "application/octet-stream"
}

export async function parseGraphError(response: Response): Promise<string> {
  try {
    const data = (await response.json()) as { error?: { message?: string; code?: string } }
    if (data.error?.message) {
      return data.error.code ? `${data.error.code}: ${data.error.message}` : data.error.message
    }
  } catch {
    // Not JSON — fall through
  }
  return `HTTP ${response.status}: ${response.statusText}`
}

const SIMPLE_UPLOAD_LIMIT = 4 * 1024 * 1024 // 4 MB
const MAX_UPLOAD_SIZE = 250 * 1024 * 1024 // 250 MB
const CHUNK_SIZE = 10 * 1024 * 1024 // 10 MB (must be multiple of 320 KiB)

async function fetchUrlToBuffer(url: string): Promise<Buffer> {
  const response = await fetch(url)
  if (!response.ok) {
    throw new Error(`Failed to fetch from sourceUrl: HTTP ${response.status} ${response.statusText}`)
  }
  return Buffer.from(await response.arrayBuffer())
}

export async function simpleUpload(
  apiBase: string,
  path: string,
  accessToken: string,
  buffer: Buffer,
  contentType: string,
  conflictBehavior: string,
): Promise<DriveItemResponse> {
  const separator = path.includes("?") ? "&" : "?"
  const url = `${apiBase}${path}${separator}@microsoft.graph.conflictBehavior=${conflictBehavior}`

  const response = await fetch(url, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": contentType,
      "Content-Length": String(buffer.length),
    },
    body: new Uint8Array(buffer),
  })

  if (!response.ok) {
    const message = await parseGraphError(response)
    throw new Error(message)
  }

  return (await response.json()) as DriveItemResponse
}

export async function sessionUpload(
  apiBase: string,
  path: string,
  accessToken: string,
  buffer: Buffer,
  conflictBehavior: string,
): Promise<DriveItemResponse> {
  // The path ends with :/content — replace :/content with :/createUploadSession
  const sessionPath = path.replace(/:\/?content$/i, ":/createUploadSession")
  const sessionUrl = `${apiBase}${sessionPath}`

  const createResponse = await fetch(sessionUrl, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      item: { "@microsoft.graph.conflictBehavior": conflictBehavior },
    }),
  })

  if (!createResponse.ok) {
    const message = await parseGraphError(createResponse)
    throw new Error(`Failed to create upload session: ${message}`)
  }

  const session = (await createResponse.json()) as { uploadUrl: string }

  return uploadChunks(session.uploadUrl, buffer, 0)
}

async function uploadChunks(uploadUrl: string, buffer: Buffer, offset: number): Promise<DriveItemResponse> {
  const totalSize = buffer.length
  if (offset >= totalSize) {
    throw new Error("Upload completed but no DriveItem response received")
  }

  const end = Math.min(offset + CHUNK_SIZE, totalSize)
  const chunk = buffer.subarray(offset, end)

  const chunkResponse = await fetch(uploadUrl, {
    method: "PUT",
    headers: {
      "Content-Length": String(chunk.length),
      "Content-Range": `bytes ${offset}-${end - 1}/${totalSize}`,
    },
    body: new Uint8Array(chunk),
  })

  if (!chunkResponse.ok) {
    // Cancel the upload session on failure
    await fetch(uploadUrl, { method: "DELETE" }).catch(() => {})
    const message = await parseGraphError(chunkResponse)
    throw new Error(`Upload chunk failed at byte ${offset}: ${message}`)
  }

  // The final chunk returns the DriveItem; intermediate chunks return 202
  if (chunkResponse.status === 200 || chunkResponse.status === 201) {
    return (await chunkResponse.json()) as DriveItemResponse
  }

  return uploadChunks(uploadUrl, buffer, offset + CHUNK_SIZE)
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
