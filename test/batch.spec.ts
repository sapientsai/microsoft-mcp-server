import { describe, expect, it } from "vitest"
import { z } from "zod"

/**
 * The batch request schema used by the microsoft_graph_batch tool.
 * Mirrors the schema in src/index.ts.
 */
const batchRequestSchema = z.object({
  requests: z
    .array(
      z.object({
        id: z.string(),
        method: z.enum(["GET", "POST", "PUT", "PATCH", "DELETE"]),
        url: z.string(),
        headers: z.record(z.string(), z.string()).optional(),
        body: z.union([z.record(z.string(), z.unknown()), z.string()]).optional(),
        dependsOn: z.array(z.string()).optional(),
      }),
    )
    .min(1)
    .max(20),
  apiVersion: z.enum(["v1.0", "beta"]).default("v1.0"),
})

/**
 * Normalize a batch request body the same way the tool does.
 * String bodies are parsed into objects; object bodies pass through.
 */
function normalizeBatchBody(body: Record<string, unknown> | string | undefined): Record<string, unknown> | undefined {
  if (body === undefined) return undefined
  return typeof body === "string" ? (JSON.parse(body) as Record<string, unknown>) : body
}

/**
 * Build the headers for a batch sub-request, auto-adding Content-Type when body is present.
 */
function buildBatchHeaders(
  headers: Record<string, string> | undefined,
  hasBody: boolean,
): Record<string, string> | undefined {
  if (!hasBody) return headers
  const merged = headers ? { ...headers } : {}
  if (!Object.keys(merged).some((k) => k.toLowerCase() === "content-type")) {
    merged["Content-Type"] = "application/json"
  }
  return merged
}

describe("microsoft_graph_batch", () => {
  describe("schema validation", () => {
    it("should accept a valid batch with one request", () => {
      const result = batchRequestSchema.safeParse({
        requests: [{ id: "1", method: "GET", url: "/me" }],
      })
      expect(result.success).toBe(true)
    })

    it("should accept a batch with 20 requests", () => {
      const requests = Array.from({ length: 20 }, (_, i) => ({
        id: String(i + 1),
        method: "GET" as const,
        url: `/users/${i}`,
      }))
      const result = batchRequestSchema.safeParse({ requests })
      expect(result.success).toBe(true)
    })

    it("should reject an empty requests array", () => {
      const result = batchRequestSchema.safeParse({ requests: [] })
      expect(result.success).toBe(false)
    })

    it("should reject more than 20 requests", () => {
      const requests = Array.from({ length: 21 }, (_, i) => ({
        id: String(i + 1),
        method: "GET" as const,
        url: `/users/${i}`,
      }))
      const result = batchRequestSchema.safeParse({ requests })
      expect(result.success).toBe(false)
    })

    it("should accept requests with dependsOn", () => {
      const result = batchRequestSchema.safeParse({
        requests: [
          { id: "1", method: "POST", url: "/me/drive/root/children", body: { name: "Parent", folder: {} } },
          {
            id: "2",
            method: "POST",
            url: "/me/drive/root:/Parent:/children",
            body: { name: "Child", folder: {} },
            dependsOn: ["1"],
          },
        ],
      })
      expect(result.success).toBe(true)
    })

    it("should accept requests with object body", () => {
      const result = batchRequestSchema.safeParse({
        requests: [{ id: "1", method: "POST", url: "/me/drive/root/children", body: { name: "Folder", folder: {} } }],
      })
      expect(result.success).toBe(true)
    })

    it("should accept requests with string body", () => {
      const result = batchRequestSchema.safeParse({
        requests: [
          {
            id: "1",
            method: "POST",
            url: "/me/drive/root/children",
            body: JSON.stringify({ name: "Folder", folder: {} }),
          },
        ],
      })
      expect(result.success).toBe(true)
    })

    it("should accept requests with custom headers", () => {
      const result = batchRequestSchema.safeParse({
        requests: [
          {
            id: "1",
            method: "POST",
            url: "/me/drive/root/children",
            headers: { "Content-Type": "application/json", "If-Match": "*" },
            body: { name: "Folder", folder: {} },
          },
        ],
      })
      expect(result.success).toBe(true)
    })

    it("should default apiVersion to v1.0", () => {
      const result = batchRequestSchema.safeParse({
        requests: [{ id: "1", method: "GET", url: "/me" }],
      })
      expect(result.success).toBe(true)
      if (result.success) {
        expect(result.data.apiVersion).toBe("v1.0")
      }
    })

    it("should accept beta apiVersion", () => {
      const result = batchRequestSchema.safeParse({
        requests: [{ id: "1", method: "GET", url: "/me" }],
        apiVersion: "beta",
      })
      expect(result.success).toBe(true)
      if (result.success) {
        expect(result.data.apiVersion).toBe("beta")
      }
    })
  })

  describe("body normalization", () => {
    it("should pass through object bodies unchanged", () => {
      const body = { name: "Folder", folder: {}, "@microsoft.graph.conflictBehavior": "rename" }
      expect(normalizeBatchBody(body)).toEqual(body)
    })

    it("should parse string bodies into objects", () => {
      const bodyObj = { name: "Folder", folder: {} }
      const bodyStr = JSON.stringify(bodyObj)
      expect(normalizeBatchBody(bodyStr)).toEqual(bodyObj)
    })

    it("should return undefined for undefined body", () => {
      expect(normalizeBatchBody(undefined)).toBeUndefined()
    })

    it("should handle complex nested string bodies", () => {
      const bodyObj = {
        message: {
          subject: "Test",
          body: { contentType: "Text", content: "Hello" },
          toRecipients: [{ emailAddress: { address: "user@example.com" } }],
        },
      }
      expect(normalizeBatchBody(JSON.stringify(bodyObj))).toEqual(bodyObj)
    })
  })

  describe("header auto-population", () => {
    it("should auto-add Content-Type when body is present and no headers provided", () => {
      const result = buildBatchHeaders(undefined, true)
      expect(result).toEqual({ "Content-Type": "application/json" })
    })

    it("should auto-add Content-Type when body is present and headers lack it", () => {
      const result = buildBatchHeaders({ "If-Match": "*" }, true)
      expect(result).toEqual({ "If-Match": "*", "Content-Type": "application/json" })
    })

    it("should not overwrite existing Content-Type header", () => {
      const result = buildBatchHeaders({ "Content-Type": "text/plain" }, true)
      expect(result).toEqual({ "Content-Type": "text/plain" })
    })

    it("should not overwrite case-insensitive Content-Type header", () => {
      const result = buildBatchHeaders({ "content-type": "text/plain" }, true)
      expect(result).toEqual({ "content-type": "text/plain" })
    })

    it("should return undefined when no body and no headers", () => {
      expect(buildBatchHeaders(undefined, false)).toBeUndefined()
    })

    it("should return headers as-is when no body", () => {
      const headers = { Accept: "application/json" }
      expect(buildBatchHeaders(headers, false)).toEqual(headers)
    })
  })

  describe("realistic batch payloads", () => {
    it("should validate a SharePoint folder tree creation batch", () => {
      const result = batchRequestSchema.safeParse({
        requests: [
          {
            id: "1",
            method: "POST",
            url: "/drives/driveId/root:/ONC-Staging:/children",
            body: { name: "Documents", folder: {}, "@microsoft.graph.conflictBehavior": "rename" },
          },
          {
            id: "2",
            method: "POST",
            url: "/drives/driveId/root:/ONC-Staging:/children",
            body: { name: "Reports", folder: {}, "@microsoft.graph.conflictBehavior": "rename" },
          },
          {
            id: "3",
            method: "POST",
            url: "/drives/driveId/root:/ONC-Staging/Documents:/children",
            body: { name: "Drafts", folder: {}, "@microsoft.graph.conflictBehavior": "rename" },
            dependsOn: ["1"],
          },
        ],
      })
      expect(result.success).toBe(true)
    })

    it("should validate a mixed-method batch", () => {
      const result = batchRequestSchema.safeParse({
        requests: [
          { id: "1", method: "GET", url: "/me" },
          { id: "2", method: "GET", url: "/me/drive/root/children" },
          {
            id: "3",
            method: "POST",
            url: "/me/drive/root/children",
            body: { name: "NewFolder", folder: {} },
          },
          { id: "4", method: "DELETE", url: "/me/drive/items/itemId" },
        ],
      })
      expect(result.success).toBe(true)
    })
  })
})
