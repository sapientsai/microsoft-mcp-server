import { describe, expect, it } from "vitest"
import { z } from "zod"

/**
 * The body parameter schema used by the microsoft_graph tool.
 * Mirrors the schema in src/index.ts — accepts both objects and pre-stringified JSON.
 */
const bodySchema = z.union([z.record(z.string(), z.unknown()), z.string()]).optional()

/**
 * The serialization logic used in the microsoft_graph tool execute function.
 * Mirrors the logic in src/index.ts — avoids double-encoding string bodies.
 */
function serializeBody(body: Record<string, unknown> | string | undefined, method: string): string | undefined {
  if (body && ["POST", "PUT", "PATCH"].includes(method)) {
    return typeof body === "string" ? body : JSON.stringify(body)
  }
  return undefined
}

describe("microsoft_graph body parameter", () => {
  describe("schema validation", () => {
    it("should accept an object body", () => {
      const result = bodySchema.safeParse({ name: "Test Folder", folder: {} })
      expect(result.success).toBe(true)
    })

    it("should accept a string body (pre-stringified JSON)", () => {
      const result = bodySchema.safeParse(JSON.stringify({ name: "Test Folder", folder: {} }))
      expect(result.success).toBe(true)
    })

    it("should accept undefined (optional)", () => {
      const result = bodySchema.safeParse(undefined)
      expect(result.success).toBe(true)
    })

    it("should accept an empty object", () => {
      const result = bodySchema.safeParse({})
      expect(result.success).toBe(true)
    })

    it("should accept an object with nested values", () => {
      const result = bodySchema.safeParse({
        message: {
          subject: "Test",
          body: { contentType: "Text", content: "Hello" },
          toRecipients: [{ emailAddress: { address: "user@example.com" } }],
        },
      })
      expect(result.success).toBe(true)
    })

    it("should reject a number", () => {
      const result = bodySchema.safeParse(42)
      expect(result.success).toBe(false)
    })

    it("should reject a boolean", () => {
      const result = bodySchema.safeParse(true)
      expect(result.success).toBe(false)
    })

    it("should reject an array", () => {
      const result = bodySchema.safeParse([1, 2, 3])
      expect(result.success).toBe(false)
    })

    it("should reject null", () => {
      const result = bodySchema.safeParse(null)
      expect(result.success).toBe(false)
    })
  })

  describe("serialization (no double-encoding)", () => {
    it("should JSON.stringify an object body for POST", () => {
      const bodyObj = { name: "Test Folder", folder: {}, "@microsoft.graph.conflictBehavior": "rename" }
      const result = serializeBody(bodyObj, "POST")

      expect(result).toBe(JSON.stringify(bodyObj))

      // Round-trips correctly
      const parsed: unknown = JSON.parse(result!)
      expect(parsed).toEqual(bodyObj)
    })

    it("should pass through a string body without double-encoding for POST", () => {
      const bodyObj = { name: "Test Folder", folder: {} }
      const bodyString = JSON.stringify(bodyObj)
      const result = serializeBody(bodyString, "POST")

      // Should be identical — no extra wrapping or escaping
      expect(result).toBe(bodyString)

      // Still valid JSON that matches the original object
      const parsed: unknown = JSON.parse(result!)
      expect(parsed).toEqual(bodyObj)
    })

    it("should detect double-encoding would corrupt the body", () => {
      const bodyObj = { name: "Test Folder", folder: {} }
      const bodyString = JSON.stringify(bodyObj)

      // This is what the OLD code did — always JSON.stringify
      const doubleEncoded = JSON.stringify(bodyString)

      // Double-encoded is wrapped in extra quotes and has escaped inner quotes
      expect(doubleEncoded).not.toBe(bodyString)
      expect(doubleEncoded.startsWith('"')).toBe(true)

      // The NEW code avoids this
      const result = serializeBody(bodyString, "POST")
      expect(result).toBe(bodyString)
      expect(result).not.toBe(doubleEncoded)
    })

    it("should JSON.stringify an object body for PUT", () => {
      const bodyObj = { displayName: "Updated Name" }
      const result = serializeBody(bodyObj, "PUT")
      expect(result).toBe(JSON.stringify(bodyObj))
    })

    it("should JSON.stringify an object body for PATCH", () => {
      const bodyObj = { subject: "Updated Subject" }
      const result = serializeBody(bodyObj, "PATCH")
      expect(result).toBe(JSON.stringify(bodyObj))
    })

    it("should pass through a string body for PUT", () => {
      const bodyString = '{"displayName":"Updated Name"}'
      const result = serializeBody(bodyString, "PUT")
      expect(result).toBe(bodyString)
    })

    it("should pass through a string body for PATCH", () => {
      const bodyString = '{"subject":"Updated Subject"}'
      const result = serializeBody(bodyString, "PATCH")
      expect(result).toBe(bodyString)
    })

    it("should return undefined for GET even with body", () => {
      const result = serializeBody({ shouldBeIgnored: true }, "GET")
      expect(result).toBeUndefined()
    })

    it("should return undefined for DELETE even with body", () => {
      const result = serializeBody({ shouldBeIgnored: true }, "DELETE")
      expect(result).toBeUndefined()
    })

    it("should return undefined when body is undefined", () => {
      const result = serializeBody(undefined, "POST")
      expect(result).toBeUndefined()
    })

    it("should handle a realistic SharePoint folder creation body as object", () => {
      const body = {
        name: "New Folder",
        folder: {},
        "@microsoft.graph.conflictBehavior": "rename",
      }
      const result = serializeBody(body, "POST")
      const parsed = JSON.parse(result!) as Record<string, unknown>
      expect(parsed.name).toBe("New Folder")
      expect(parsed.folder).toEqual({})
      expect(parsed["@microsoft.graph.conflictBehavior"]).toBe("rename")
    })

    it("should handle a realistic SharePoint folder creation body as string", () => {
      const body = {
        name: "New Folder",
        folder: {},
        "@microsoft.graph.conflictBehavior": "rename",
      }
      const bodyString = JSON.stringify(body)
      const result = serializeBody(bodyString, "POST")
      const parsed = JSON.parse(result!) as Record<string, unknown>
      expect(parsed.name).toBe("New Folder")
      expect(parsed.folder).toEqual({})
      expect(parsed["@microsoft.graph.conflictBehavior"]).toBe("rename")
    })

    it("should handle a realistic send mail body as object", () => {
      const body = {
        message: {
          subject: "Test Email",
          body: { contentType: "Text", content: "Hello World" },
          toRecipients: [{ emailAddress: { address: "user@example.com" } }],
        },
        saveToSentItems: true,
      }
      const result = serializeBody(body, "POST")
      const parsed = JSON.parse(result!) as { message: { subject: string } }
      expect(parsed.message.subject).toBe("Test Email")
    })
  })
})
