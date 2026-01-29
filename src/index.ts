import type { AzureSession, OAuthSession } from "fastmcp"
import { AzureProvider, FastMCP } from "fastmcp"
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

function validateConfig(config: ServerConfig): void {
  if (config.authMode === "clientCredentials") {
    if (config.tenantId === "common") {
      throw new Error('Client credentials auth requires a specific tenant ID, not "common".')
    }
    if (!config.clientSecret) {
      throw new Error("Client credentials auth requires AZURE_CLIENT_SECRET.")
    }
  }
}

export function createConfig(): ServerConfig {
  const authMode = parseAuthMode(process.env.AZURE_AUTH_MODE)

  const config: ServerConfig = {
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
    apiKey: process.env.MCP_API_KEY || undefined,
  }

  validateConfig(config)

  return config
}

export function createServer(config: ServerConfig) {
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

  const server = new FastMCP({
    name: "microsoft-graph-server",
    version: "1.0.0",
    instructions:
      "Microsoft Graph MCP Server - Access Microsoft 365 data including users, mail, calendar, files, and more.",
    auth: authProvider,
    health: {
      enabled: true,
      path: "/health",
      message: "healthy",
      status: 200,
    },
    authenticate: config.apiKey
      ? async (request) => {
          // Check Authorization header first
          const authHeader = request.headers.authorization
          let providedKey = typeof authHeader === "string" ? authHeader.replace(/^Bearer\s+/i, "") : undefined

          // Fall back to api_key query parameter if no header
          if (!providedKey && request.url) {
            try {
              const url = new URL(request.url, "http://localhost")
              providedKey = url.searchParams.get("api_key") ?? undefined
            } catch {
              // Ignore URL parsing errors
            }
          }

          if (providedKey !== config.apiKey) {
            throw new Error("Unauthorized")
          }
          return {}
        }
      : undefined,
  })

  server.addTool({
    name: "microsoft_graph",
    description:
      "Execute Microsoft Graph API requests. Use this to access Microsoft 365 data including users, mail, calendar, files, and more.",
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
      let accessToken: string

      if (isClientCredentials && tokenManager) {
        // Client credentials mode - get token from token manager
        accessToken = await tokenManager.getToken()
      } else {
        // Interactive mode - get token from session
        const authSession = session as OAuthSession | undefined
        if (!authSession?.accessToken) {
          throw new Error("Not authenticated. Please sign in first.")
        }
        accessToken = authSession.accessToken
      }

      const baseUrl = args.apiType === "azure" ? AZURE_BASE_URL : GRAPH_BASE_URL
      let url: string

      if (args.apiType === "azure") {
        url = `${baseUrl}${args.path}`
      } else {
        url = `${baseUrl}/${args.apiVersion}${args.path}`
      }

      if (args.queryParams && Object.keys(args.queryParams).length > 0) {
        const params = new URLSearchParams(args.queryParams)
        url += `?${params.toString()}`
      }

      log.info("Calling Microsoft Graph API", { url, method: args.method })

      const fetchOptions: RequestInit = {
        method: args.method,
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
      }

      if (args.body && ["POST", "PUT", "PATCH"].includes(args.method)) {
        fetchOptions.body = JSON.stringify(args.body)
      }

      const response = await fetch(url, fetchOptions)

      let data: unknown
      const contentType = response.headers.get("content-type")
      if (contentType?.includes("application/json")) {
        data = await response.json()
      } else {
        data = await response.text()
      }

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

  return { server, authProvider, tokenManager, config }
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
