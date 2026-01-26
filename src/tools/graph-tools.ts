import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { z } from "zod"

import { GraphClient } from "../client/index.js"

const ApiTypeSchema = z.enum(["graph", "azure"]).default("graph")
const HttpMethodSchema = z.enum(["GET", "POST", "PUT", "PATCH", "DELETE"]).default("GET")
const ApiVersionSchema = z.enum(["v1.0", "beta"]).default("v1.0")

export function registerGraphTools(server: McpServer, graphClient: GraphClient): void {
  server.tool(
    "microsoft_graph",
    "Execute Microsoft Graph API requests. Use this to access Microsoft 365 data including users, mail, calendar, files, and more.",
    {
      apiType: ApiTypeSchema.describe('API type: "graph" for Microsoft Graph, "azure" for Azure Management'),
      path: z.string().describe("API endpoint path (e.g., /me, /users, /me/messages)"),
      method: HttpMethodSchema.describe("HTTP method"),
      apiVersion: ApiVersionSchema.describe("Graph API version"),
      queryParams: z
        .record(z.string(), z.string())
        .optional()
        .describe("OData query parameters ($select, $filter, $top, $orderby, etc.)"),
      body: z.unknown().optional().describe("Request body for POST/PUT/PATCH operations"),
    },
    async ({ apiType, path, method, apiVersion, queryParams, body }) => {
      try {
        const response = await graphClient.callApi({
          apiType,
          path,
          method,
          apiVersion,
          queryParams,
          body,
        })

        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify(response.data, null, 2),
            },
          ],
        }
      } catch (error) {
        const message = error instanceof Error ? error.message : "Unknown error occurred"
        return {
          content: [
            {
              type: "text" as const,
              text: `Error: ${message}`,
            },
          ],
          isError: true,
        }
      }
    },
  )
}
