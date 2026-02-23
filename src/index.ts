import { mkdir, writeFile } from "node:fs/promises"
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

export type { TokenManager } from "./auth/token-manager.js"
export { createTokenManager } from "./auth/token-manager.js"
export type { AuthMode, ServerConfig } from "./auth/types.js"

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

  const baseInstructions =
    "Microsoft Graph MCP Server - Access Microsoft 365 data including users, mail, calendar, files, and more."
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
      body: z.unknown().optional().describe("Request body for POST/PUT/PATCH operations"),
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
        ...(args.body && ["POST", "PUT", "PATCH"].includes(args.method) ? { body: JSON.stringify(args.body) } : {}),
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
    name: "read_document",
    description:
      "Download a file from SharePoint or OneDrive and return its readable text content. Supports DOCX, PDF, XLSX, and text-based files. Use this instead of download_file when you need to read document contents.",
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

      const text = await extractTextFromBuffer(buffer, contentType, resolvedFilename)

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

  return { server, authProvider, tokenManager, config }
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
