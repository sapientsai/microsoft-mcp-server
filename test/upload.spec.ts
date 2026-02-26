import { describe, expect, it } from "vitest"
import { z } from "zod"

import { filenameFromPath, filenameFromUrl, parseGraphError, resolveUploadContentType } from "../src/index.js"

/**
 * The get_upload_config parameter schema — mirrors src/index.ts.
 */
const uploadConfigSchema = z.object({
  path: z.string(),
  localFile: z.string().optional(),
  contentType: z.string().optional(),
  conflictBehavior: z.enum(["rename", "replace", "fail"]).default("rename"),
})

const SIMPLE_UPLOAD_LIMIT = 4 * 1024 * 1024
const MAX_UPLOAD_SIZE = 250 * 1024 * 1024

function classifyUploadMethod(size: number): "simple" | "session" | "rejected" {
  if (size > MAX_UPLOAD_SIZE) return "rejected"
  if (size <= SIMPLE_UPLOAD_LIMIT) return "simple"
  return "session"
}

describe("upload", () => {
  describe("get_upload_config schema", () => {
    it("should accept path only", () => {
      const result = uploadConfigSchema.safeParse({
        path: "/drives/driveId/root:/folder/file.pdf:/content",
      })
      expect(result.success).toBe(true)
    })

    it("should accept path with localFile", () => {
      const result = uploadConfigSchema.safeParse({
        path: "/drives/driveId/root:/folder/file.pdf:/content",
        localFile: "/mnt/user-data/uploads/file.pdf",
      })
      expect(result.success).toBe(true)
    })

    it("should accept path with contentType override", () => {
      const result = uploadConfigSchema.safeParse({
        path: "/drives/driveId/root:/folder/file.pdf:/content",
        contentType: "application/pdf",
      })
      expect(result.success).toBe(true)
    })

    it("should default conflictBehavior to rename", () => {
      const result = uploadConfigSchema.safeParse({
        path: "/drives/driveId/root:/folder/file.pdf:/content",
      })
      expect(result.success).toBe(true)
      if (result.success) {
        expect(result.data.conflictBehavior).toBe("rename")
      }
    })

    it("should accept replace conflictBehavior", () => {
      const result = uploadConfigSchema.safeParse({
        path: "/drives/driveId/root:/folder/file.pdf:/content",
        conflictBehavior: "replace",
      })
      expect(result.success).toBe(true)
      if (result.success) {
        expect(result.data.conflictBehavior).toBe("replace")
      }
    })

    it("should accept fail conflictBehavior", () => {
      const result = uploadConfigSchema.safeParse({
        path: "/drives/driveId/root:/folder/file.pdf:/content",
        conflictBehavior: "fail",
      })
      expect(result.success).toBe(true)
      if (result.success) {
        expect(result.data.conflictBehavior).toBe("fail")
      }
    })

    it("should reject invalid conflictBehavior", () => {
      const result = uploadConfigSchema.safeParse({
        path: "/drives/driveId/root:/folder/file.pdf:/content",
        conflictBehavior: "skip",
      })
      expect(result.success).toBe(false)
    })

    it("should require path", () => {
      const result = uploadConfigSchema.safeParse({})
      expect(result.success).toBe(false)
    })
  })

  describe("MIME type resolution", () => {
    it("should resolve .pdf", () => {
      expect(resolveUploadContentType(undefined, "report.pdf")).toBe("application/pdf")
    })

    it("should resolve .docx", () => {
      expect(resolveUploadContentType(undefined, "document.docx")).toBe(
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      )
    })

    it("should resolve .xlsx", () => {
      expect(resolveUploadContentType(undefined, "spreadsheet.xlsx")).toBe(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      )
    })

    it("should resolve .pptx", () => {
      expect(resolveUploadContentType(undefined, "slides.pptx")).toBe(
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      )
    })

    it("should resolve .png", () => {
      expect(resolveUploadContentType(undefined, "image.png")).toBe("image/png")
    })

    it("should resolve .jpg", () => {
      expect(resolveUploadContentType(undefined, "photo.jpg")).toBe("image/jpeg")
    })

    it("should resolve .csv", () => {
      expect(resolveUploadContentType(undefined, "data.csv")).toBe("text/csv")
    })

    it("should resolve .json", () => {
      expect(resolveUploadContentType(undefined, "config.json")).toBe("application/json")
    })

    it("should resolve .zip", () => {
      expect(resolveUploadContentType(undefined, "archive.zip")).toBe("application/zip")
    })

    it("should return octet-stream for unknown extensions", () => {
      expect(resolveUploadContentType(undefined, "data.xyz")).toBe("application/octet-stream")
    })

    it("should return octet-stream for no extension", () => {
      expect(resolveUploadContentType(undefined, "README")).toBe("application/octet-stream")
    })

    it("should use explicit contentType over auto-detection", () => {
      expect(resolveUploadContentType("text/plain", "file.pdf")).toBe("text/plain")
    })

    it("should be case-insensitive for extensions", () => {
      expect(resolveUploadContentType(undefined, "FILE.PDF")).toBe("application/pdf")
      expect(resolveUploadContentType(undefined, "image.PNG")).toBe("image/png")
    })
  })

  describe("filename extraction", () => {
    it("should extract filename from Graph API colon path", () => {
      expect(filenameFromPath("/drives/driveId/root:/folder/report.pdf:/content")).toBe("report.pdf")
    })

    it("should extract filename from nested path", () => {
      expect(filenameFromPath("/me/drive/root:/Documents/Sub/file.docx:/content")).toBe("file.docx")
    })

    it("should return undefined for item ID paths", () => {
      expect(filenameFromPath("/me/drive/items/itemId/content")).toBeUndefined()
    })
  })

  describe("size threshold", () => {
    it("should use simple upload for ≤4MB", () => {
      expect(classifyUploadMethod(0)).toBe("simple")
      expect(classifyUploadMethod(1024)).toBe("simple")
      expect(classifyUploadMethod(SIMPLE_UPLOAD_LIMIT)).toBe("simple")
    })

    it("should use session upload for >4MB", () => {
      expect(classifyUploadMethod(SIMPLE_UPLOAD_LIMIT + 1)).toBe("session")
      expect(classifyUploadMethod(10 * 1024 * 1024)).toBe("session")
      expect(classifyUploadMethod(MAX_UPLOAD_SIZE)).toBe("session")
    })

    it("should reject >250MB", () => {
      expect(classifyUploadMethod(MAX_UPLOAD_SIZE + 1)).toBe("rejected")
    })
  })

  describe("filenameFromUrl", () => {
    it("should extract filename from URL path", () => {
      expect(filenameFromUrl("https://example.com/files/report.pdf")).toBe("report.pdf")
    })

    it("should extract filename from URL with query string", () => {
      expect(filenameFromUrl("https://example.com/files/doc.docx?token=abc")).toBe("doc.docx")
    })

    it("should decode URL-encoded filenames", () => {
      expect(filenameFromUrl("https://example.com/files/my%20report.pdf")).toBe("my report.pdf")
    })

    it("should return undefined for root-only URLs", () => {
      expect(filenameFromUrl("https://example.com/")).toBeUndefined()
    })

    it("should return undefined for invalid URLs", () => {
      expect(filenameFromUrl("not-a-url")).toBeUndefined()
    })

    it("should handle deeply nested paths", () => {
      expect(filenameFromUrl("https://cdn.example.com/a/b/c/d/file.xlsx")).toBe("file.xlsx")
    })
  })

  describe("parseGraphError", () => {
    it("should extract error message from JSON response", async () => {
      const response = new Response(JSON.stringify({ error: { code: "accessDenied", message: "Access denied" } }), {
        status: 403,
        statusText: "Forbidden",
        headers: { "Content-Type": "application/json" },
      })
      const message = await parseGraphError(response)
      expect(message).toBe("accessDenied: Access denied")
    })

    it("should extract message without code", async () => {
      const response = new Response(JSON.stringify({ error: { message: "Not found" } }), {
        status: 404,
        statusText: "Not Found",
        headers: { "Content-Type": "application/json" },
      })
      const message = await parseGraphError(response)
      expect(message).toBe("Not found")
    })

    it("should fall back to HTTP status for non-JSON response", async () => {
      const response = new Response("Server Error", {
        status: 500,
        statusText: "Internal Server Error",
      })
      const message = await parseGraphError(response)
      expect(message).toBe("HTTP 500: Internal Server Error")
    })

    it("should fall back to HTTP status for empty error object", async () => {
      const response = new Response(JSON.stringify({}), {
        status: 400,
        statusText: "Bad Request",
        headers: { "Content-Type": "application/json" },
      })
      const message = await parseGraphError(response)
      expect(message).toBe("HTTP 400: Bad Request")
    })
  })
})
