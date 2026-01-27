import type { AzureSession, OAuthSession } from "fastmcp"
import { AzureProvider, FastMCP, getAuthSession } from "fastmcp"
import { z } from "zod"

export const DEFAULT_CLIENT_ID = "cf7d1f97-781e-4034-930c-abd420e12d49"
export const GRAPH_BASE_URL = "https://graph.microsoft.com"
export const AZURE_BASE_URL = "https://management.azure.com"

export type ServerConfig = {
  clientId: string
  clientSecret: string
  tenantId: string
  baseUrl: string
  port: number
  scopes: string[]
}

export function createConfig(): ServerConfig {
  return {
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
  }
}

export function createServer(config: ServerConfig) {
  const authProvider = new AzureProvider({
    clientId: config.clientId,
    clientSecret: config.clientSecret,
    baseUrl: config.baseUrl,
    tenantId: config.tenantId,
    scopes: config.scopes,
  })

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
      const authSession = session as OAuthSession | undefined
      if (!authSession?.accessToken) {
        throw new Error("Not authenticated. Please sign in first.")
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
          Authorization: `Bearer ${authSession!.accessToken}`,
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
      const authSession = session as AzureSession | undefined
      const hasAuth = !!authSession?.accessToken
      return JSON.stringify(
        {
          authenticated: hasAuth,
          message: hasAuth ? "Authenticated via Azure OAuth" : "Not authenticated. Please sign in.",
          scopes: authSession?.scopes,
          upn: authSession?.upn,
        },
        null,
        2,
      )
    },
  })

  return { server, authProvider, config }
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
      },
    })
    console.error(`Microsoft Graph MCP Server running on http://localhost:${config.port}`)
    console.error(`OAuth callback URL: ${config.baseUrl}/oauth/callback`)
  }
}
