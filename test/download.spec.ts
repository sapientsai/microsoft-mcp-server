import { existsSync } from "node:fs"
import { readFile, rm } from "node:fs/promises"
import { tmpdir } from "node:os"
import { join } from "node:path"

import { afterEach, describe, expect, it } from "vitest"

import {
  filenameFromHeaders,
  filenameFromPath,
  formatBytes,
  isTextContent,
  processDownloadResponse,
} from "../src/download/download.js"

// Minimal 1x1 red PNG (68 bytes)
const TINY_PNG = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg==",
  "base64",
)

const CSV_CONTENT = "name,age,city\nAlice,30,Boston\nBob,25,NYC\n"
const JSON_CONTENT = JSON.stringify({ users: [{ name: "Alice" }] })

describe("Download Helpers", () => {
  describe("isTextContent", () => {
    it("should identify text/* MIME types", () => {
      expect(isTextContent("text/plain")).toBe(true)
      expect(isTextContent("text/csv")).toBe(true)
      expect(isTextContent("text/html")).toBe(true)
      expect(isTextContent("text/xml")).toBe(true)
    })

    it("should identify application/json", () => {
      expect(isTextContent("application/json")).toBe(true)
      expect(isTextContent("application/json; charset=utf-8")).toBe(true)
    })

    it("should identify application/xml", () => {
      expect(isTextContent("application/xml")).toBe(true)
    })

    it("should identify application/csv", () => {
      expect(isTextContent("application/csv")).toBe(true)
    })

    it("should identify +xml and +json suffixes", () => {
      expect(isTextContent("application/atom+xml")).toBe(true)
      expect(isTextContent("application/hal+json")).toBe(true)
      expect(isTextContent("application/vnd.api+json")).toBe(true)
    })

    it("should reject binary MIME types", () => {
      expect(isTextContent("application/octet-stream")).toBe(false)
      expect(isTextContent("application/pdf")).toBe(false)
      expect(isTextContent("application/vnd.openxmlformats-officedocument.wordprocessingml.document")).toBe(false)
      expect(isTextContent("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")).toBe(false)
    })

    it("should reject image MIME types", () => {
      expect(isTextContent("image/png")).toBe(false)
      expect(isTextContent("image/jpeg")).toBe(false)
    })

    it("should be case-insensitive", () => {
      expect(isTextContent("TEXT/PLAIN")).toBe(true)
      expect(isTextContent("Application/JSON")).toBe(true)
    })
  })

  describe("filenameFromHeaders", () => {
    it("should extract filename from Content-Disposition", () => {
      const headers = new Headers({ "content-disposition": 'attachment; filename="report.pdf"' })
      const result = filenameFromHeaders(headers)
      expect(result.isSome()).toBe(true)
      expect(result.orThrow()).toBe("report.pdf")
    })

    it("should extract filename without quotes", () => {
      const headers = new Headers({ "content-disposition": "attachment; filename=report.pdf" })
      const result = filenameFromHeaders(headers)
      expect(result.isSome()).toBe(true)
      expect(result.orThrow()).toBe("report.pdf")
    })

    it("should handle UTF-8 encoded filenames", () => {
      const headers = new Headers({ "content-disposition": "attachment; filename*=UTF-8''budget%202024.xlsx" })
      const result = filenameFromHeaders(headers)
      expect(result.isSome()).toBe(true)
      expect(result.orThrow()).toBe("budget 2024.xlsx")
    })

    it("should return None when no Content-Disposition", () => {
      const headers = new Headers({})
      expect(filenameFromHeaders(headers).isNone()).toBe(true)
    })

    it("should return None for malformed Content-Disposition", () => {
      const headers = new Headers({ "content-disposition": "inline" })
      expect(filenameFromHeaders(headers).isNone()).toBe(true)
    })
  })

  describe("filenameFromPath", () => {
    it("should extract filename from colon-path format", () => {
      const result = filenameFromPath("/me/drive/root:/Documents/report.pdf:/content")
      expect(result.isSome()).toBe(true)
      expect(result.orThrow()).toBe("report.pdf")
    })

    it("should extract filename from nested colon-path", () => {
      const result = filenameFromPath("/me/drive/root:/Projects/2024/budget.xlsx:/content")
      expect(result.isSome()).toBe(true)
      expect(result.orThrow()).toBe("budget.xlsx")
    })

    it("should return None for item ID paths", () => {
      expect(filenameFromPath("/me/drive/items/ABC123/content").isNone()).toBe(true)
    })

    it("should return None for paths without content suffix", () => {
      expect(filenameFromPath("/me/drive/root:/Documents/report.pdf").isNone()).toBe(true)
    })
  })

  describe("formatBytes", () => {
    it("should format 0 bytes", () => {
      expect(formatBytes(0)).toBe("0 B")
    })

    it("should format bytes", () => {
      expect(formatBytes(500)).toBe("500 B")
    })

    it("should format kilobytes", () => {
      expect(formatBytes(1024)).toBe("1.0 KB")
      expect(formatBytes(1536)).toBe("1.5 KB")
    })

    it("should format megabytes", () => {
      expect(formatBytes(1048576)).toBe("1.0 MB")
      expect(formatBytes(5242880)).toBe("5.0 MB")
    })

    it("should format gigabytes", () => {
      expect(formatBytes(1073741824)).toBe("1.0 GB")
    })
  })
})

describe("processDownloadResponse", () => {
  const testOutputDir = join(tmpdir(), "microsoft-graph-test-downloads")

  afterEach(async () => {
    await rm(testOutputDir, { recursive: true, force: true })
  })

  describe("image content", () => {
    it("should return inline image content for PNG", async () => {
      const result = await processDownloadResponse(TINY_PNG, "image/png", "photo.png")

      expect(result.content).toHaveLength(2)
      expect(result.content[0]).toEqual({
        type: "text",
        text: expect.stringContaining("photo.png"),
      })
      expect(result.content[1]).toHaveProperty("type", "image")
      expect(result.content[1]).toHaveProperty("mimeType", "image/png")
      expect(result.content[1]).toHaveProperty("data")
    })

    it("should return inline image content for JPEG", async () => {
      // Minimal JPEG (just enough to be valid-ish for the content type check)
      const jpegBuffer = Buffer.from([0xff, 0xd8, 0xff, 0xe0, 0x00, 0x10])
      const result = await processDownloadResponse(jpegBuffer, "image/jpeg", "photo.jpg")

      expect(result.content).toHaveLength(2)
      expect(result.content[0]).toEqual({
        type: "text",
        text: expect.stringContaining("photo.jpg"),
      })
      expect(result.content[1]).toHaveProperty("type", "image")
    })

    it("should include file size in image text label", async () => {
      const result = await processDownloadResponse(TINY_PNG, "image/png", "large-image.png")

      const textContent = result.content[0] as { type: "text"; text: string }
      expect(textContent.text).toContain("large-image.png")
      expect(textContent.text).toMatch(/\d+(\.\d+)?\s*(B|KB|MB|GB)/)
    })

    it("should return inline image content for SVG", async () => {
      const svgBuffer = Buffer.from('<svg xmlns="http://www.w3.org/2000/svg"><circle r="10"/></svg>')
      const result = await processDownloadResponse(svgBuffer, "image/svg+xml", "icon.svg")

      expect(result.content).toHaveLength(2)
      expect(result.content[1]).toHaveProperty("type", "image")
    })
  })

  describe("text content", () => {
    it("should return inline text for CSV files", async () => {
      const buffer = Buffer.from(CSV_CONTENT)
      const result = await processDownloadResponse(buffer, "text/csv", "data.csv")

      expect(result.content).toHaveLength(1)
      const textContent = result.content[0] as { type: "text"; text: string }
      expect(textContent.type).toBe("text")
      expect(textContent.text).toContain("data.csv")
      expect(textContent.text).toContain("text/csv")
      expect(textContent.text).toContain("Alice,30,Boston")
      expect(textContent.text).toContain("Bob,25,NYC")
    })

    it("should return inline text for JSON files", async () => {
      const buffer = Buffer.from(JSON_CONTENT)
      const result = await processDownloadResponse(buffer, "application/json", "users.json")

      expect(result.content).toHaveLength(1)
      const textContent = result.content[0] as { type: "text"; text: string }
      expect(textContent.type).toBe("text")
      expect(textContent.text).toContain("users.json")
      expect(textContent.text).toContain('"Alice"')
    })

    it("should return inline text for plain text files", async () => {
      const buffer = Buffer.from("Hello, world!\nLine 2\n")
      const result = await processDownloadResponse(buffer, "text/plain", "notes.txt")

      const textContent = result.content[0] as { type: "text"; text: string }
      expect(textContent.text).toContain("notes.txt")
      expect(textContent.text).toContain("Hello, world!")
      expect(textContent.text).toContain("Line 2")
    })

    it("should return inline text for XML files", async () => {
      const buffer = Buffer.from('<?xml version="1.0"?><root><item>test</item></root>')
      const result = await processDownloadResponse(buffer, "application/xml", "config.xml")

      const textContent = result.content[0] as { type: "text"; text: string }
      expect(textContent.text).toContain("config.xml")
      expect(textContent.text).toContain("<item>test</item>")
    })

    it("should return inline text for HTML files", async () => {
      const buffer = Buffer.from("<html><body><h1>Hello</h1></body></html>")
      const result = await processDownloadResponse(buffer, "text/html", "page.html")

      const textContent = result.content[0] as { type: "text"; text: string }
      expect(textContent.text).toContain("page.html")
      expect(textContent.text).toContain("<h1>Hello</h1>")
    })

    it("should include file size and content type in header", async () => {
      const buffer = Buffer.from(CSV_CONTENT)
      const result = await processDownloadResponse(buffer, "text/csv", "data.csv")

      const textContent = result.content[0] as { type: "text"; text: string }
      expect(textContent.text).toMatch(/\d+(\.\d+)?\s*(B|KB|MB|GB)/)
      expect(textContent.text).toContain("text/csv")
    })
  })

  describe("binary content (base64)", () => {
    it("should return base64-encoded PDF content", async () => {
      const pdfBuffer = Buffer.from("%PDF-1.4 fake pdf content for testing")
      const result = await processDownloadResponse(pdfBuffer, "application/pdf", "report.pdf")

      expect(result.content).toHaveLength(1)
      const textContent = result.content[0] as { type: "text"; text: string }
      const parsed = JSON.parse(textContent.text) as {
        filename: string
        contentType: string
        size: number
        sizeFormatted: string
        encoding: string
        data: string
      }

      expect(parsed.filename).toBe("report.pdf")
      expect(parsed.contentType).toBe("application/pdf")
      expect(parsed.size).toBe(pdfBuffer.length)
      expect(parsed.sizeFormatted).toMatch(/\d+(\.\d+)?\s*(B|KB|MB|GB)/)
      expect(parsed.encoding).toBe("base64")
      expect(parsed.data).toBe(pdfBuffer.toString("base64"))

      // Verify round-trip: base64 decodes back to original content
      const decoded = Buffer.from(parsed.data, "base64")
      expect(decoded.toString()).toBe("%PDF-1.4 fake pdf content for testing")
    })

    it("should return base64-encoded DOCX content", async () => {
      const docxBuffer = Buffer.from("PK\x03\x04 fake docx content")
      const result = await processDownloadResponse(
        docxBuffer,
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "document.docx",
      )

      const textContent = result.content[0] as { type: "text"; text: string }
      const parsed = JSON.parse(textContent.text) as {
        filename: string
        contentType: string
        encoding: string
        data: string
      }
      expect(parsed.filename).toBe("document.docx")
      expect(parsed.contentType).toBe("application/vnd.openxmlformats-officedocument.wordprocessingml.document")
      expect(parsed.encoding).toBe("base64")
      expect(Buffer.from(parsed.data, "base64").toString()).toBe("PK\x03\x04 fake docx content")
    })

    it("should return base64-encoded XLSX content", async () => {
      const xlsxBuffer = Buffer.from("PK\x03\x04 fake xlsx content")
      const result = await processDownloadResponse(
        xlsxBuffer,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "data.xlsx",
      )

      const textContent = result.content[0] as { type: "text"; text: string }
      const parsed = JSON.parse(textContent.text) as { filename: string; encoding: string; data: string }
      expect(parsed.filename).toBe("data.xlsx")
      expect(parsed.encoding).toBe("base64")
      expect(Buffer.from(parsed.data, "base64").toString()).toBe("PK\x03\x04 fake xlsx content")
    })

    it("should return base64-encoded PPTX content", async () => {
      const pptxBuffer = Buffer.from("PK\x03\x04 fake pptx content")
      const result = await processDownloadResponse(
        pptxBuffer,
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "slides.pptx",
      )

      const textContent = result.content[0] as { type: "text"; text: string }
      const parsed = JSON.parse(textContent.text) as { filename: string; encoding: string }
      expect(parsed.filename).toBe("slides.pptx")
      expect(parsed.encoding).toBe("base64")
    })

    it("should return base64-encoded octet-stream content", async () => {
      const binaryBuffer = Buffer.from([0x00, 0x01, 0x02, 0x03, 0xff])
      const result = await processDownloadResponse(binaryBuffer, "application/octet-stream", "unknown.bin")

      const textContent = result.content[0] as { type: "text"; text: string }
      const parsed = JSON.parse(textContent.text) as {
        filename: string
        contentType: string
        encoding: string
        data: string
      }
      expect(parsed.filename).toBe("unknown.bin")
      expect(parsed.contentType).toBe("application/octet-stream")
      expect(parsed.encoding).toBe("base64")

      const decoded = Buffer.from(parsed.data, "base64")
      expect([...decoded]).toEqual([0x00, 0x01, 0x02, 0x03, 0xff])
    })

    it("should not include savedTo when no outputDir specified", async () => {
      const pdfBuffer = Buffer.from("%PDF-1.4 test")
      const result = await processDownloadResponse(pdfBuffer, "application/pdf", "no-disk.pdf")

      const textContent = result.content[0] as { type: "text"; text: string }
      const parsed = JSON.parse(textContent.text) as Record<string, unknown>
      expect(parsed).not.toHaveProperty("savedTo")
    })

    it("should also save to disk when outputDir is provided", async () => {
      const pdfBuffer = Buffer.from("%PDF-1.4 fake pdf content for testing")
      const result = await processDownloadResponse(pdfBuffer, "application/pdf", "report.pdf", testOutputDir)

      const textContent = result.content[0] as { type: "text"; text: string }
      const parsed = JSON.parse(textContent.text) as { savedTo: string; encoding: string; data: string }
      expect(parsed.savedTo).toContain("report.pdf")
      expect(parsed.encoding).toBe("base64")
      expect(existsSync(parsed.savedTo)).toBe(true)

      const savedContent = await readFile(parsed.savedTo)
      expect(savedContent.toString()).toBe("%PDF-1.4 fake pdf content for testing")
    })

    it("should create nested output directory when saving", async () => {
      const nestedDir = join(testOutputDir, "nested", "deep")
      const pdfBuffer = Buffer.from("%PDF-1.4 test")
      const result = await processDownloadResponse(pdfBuffer, "application/pdf", "test.pdf", nestedDir)

      const textContent = result.content[0] as { type: "text"; text: string }
      const parsed = JSON.parse(textContent.text) as { savedTo: string }
      expect(existsSync(parsed.savedTo)).toBe(true)
    })
  })
})
