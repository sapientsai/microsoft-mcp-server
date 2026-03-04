import ExcelJS from "exceljs"
import { describe, expect, it } from "vitest"

import { EXTRACTABLE_TYPES, extractTextFromBuffer } from "../src/download/extract.js"

// Helper to create a minimal valid XLSX buffer with data
async function createXlsxBuffer(sheets: Record<string, string[][]>): Promise<Buffer> {
  const wb = new ExcelJS.Workbook()
  for (const [name, rows] of Object.entries(sheets)) {
    const ws = wb.addWorksheet(name)
    for (const row of rows) {
      ws.addRow(row)
    }
  }
  const arrayBuffer = await wb.xlsx.writeBuffer()
  return Buffer.from(arrayBuffer)
}

describe("extractTextFromBuffer", () => {
  describe("text-based content", () => {
    it("should return plain text content directly", async () => {
      const buffer = Buffer.from("Hello, world!")
      const result = await extractTextFromBuffer(buffer, "text/plain", "notes.txt")
      expect(result.isRight()).toBe(true)
      expect(result.orThrow()).toBe("Hello, world!")
    })

    it("should return JSON content directly", async () => {
      const json = JSON.stringify({ name: "Alice" })
      const buffer = Buffer.from(json)
      const result = await extractTextFromBuffer(buffer, "application/json", "data.json")
      expect(result.isRight()).toBe(true)
      expect(result.orThrow()).toBe(json)
    })

    it("should return CSV content directly", async () => {
      const csv = "name,age\nAlice,30\n"
      const buffer = Buffer.from(csv)
      const result = await extractTextFromBuffer(buffer, "text/csv", "data.csv")
      expect(result.isRight()).toBe(true)
      expect(result.orThrow()).toBe(csv)
    })

    it("should return XML content directly", async () => {
      const xml = "<root><item>test</item></root>"
      const buffer = Buffer.from(xml)
      const result = await extractTextFromBuffer(buffer, "application/xml", "config.xml")
      expect(result.isRight()).toBe(true)
      expect(result.orThrow()).toBe(xml)
    })
  })

  describe("XLSX extraction", () => {
    it("should extract text from single-sheet XLSX", async () => {
      const buffer = await createXlsxBuffer({
        Sheet1: [
          ["Name", "Age"],
          ["Alice", "30"],
          ["Bob", "25"],
        ],
      })

      const result = await extractTextFromBuffer(
        buffer,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "data.xlsx",
      )

      expect(result.isRight()).toBe(true)
      const text = result.orThrow()
      expect(text).toContain("Name")
      expect(text).toContain("Alice")
      expect(text).toContain("30")
      expect(text).toContain("Bob")
      expect(text).toContain("25")
      // Single sheet should not include sheet header
      expect(text).not.toContain("[Sheet:")
    })

    it("should extract text from multi-sheet XLSX with headers", async () => {
      const buffer = await createXlsxBuffer({
        People: [
          ["Name", "Age"],
          ["Alice", "30"],
        ],
        Cities: [
          ["City", "Country"],
          ["Boston", "US"],
        ],
      })

      const result = await extractTextFromBuffer(
        buffer,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "data.xlsx",
      )

      expect(result.isRight()).toBe(true)
      const text = result.orThrow()
      expect(text).toContain("[Sheet: People]")
      expect(text).toContain("Alice")
      expect(text).toContain("[Sheet: Cities]")
      expect(text).toContain("Boston")
    })

    it("should infer XLSX type from extension when content type is octet-stream", async () => {
      const buffer = await createXlsxBuffer({
        Sheet1: [["Test", "Data"]],
      })

      const result = await extractTextFromBuffer(buffer, "application/octet-stream", "report.xlsx")

      expect(result.isRight()).toBe(true)
      const text = result.orThrow()
      expect(text).toContain("Test")
      expect(text).toContain("Data")
    })
  })

  describe("DOCX extraction", () => {
    it("should return Left for invalid DOCX buffer", async () => {
      const docxContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      const result = await extractTextFromBuffer(Buffer.from("not a docx"), docxContentType, "test.docx")
      expect(result.isLeft()).toBe(true)
    })

    it("should infer DOCX type from extension when content type is octet-stream", async () => {
      const result = await extractTextFromBuffer(Buffer.from("not a docx"), "application/octet-stream", "document.docx")
      expect(result.isLeft()).toBe(true)
    })
  })

  describe("PDF extraction", () => {
    it("should extract text from a minimal PDF", async () => {
      // Create a minimal valid PDF with text content
      const pdfContent = `%PDF-1.0
1 0 obj
<< /Type /Catalog /Pages 2 0 R >>
endobj

2 0 obj
<< /Type /Pages /Kids [3 0 R] /Count 1 >>
endobj

3 0 obj
<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>
endobj

4 0 obj
<< /Length 44 >>
stream
BT /F1 12 Tf 100 700 Td (Hello PDF) Tj ET
endstream
endobj

5 0 obj
<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>
endobj

xref
0 6
0000000000 65535 f
0000000009 00000 n
0000000058 00000 n
0000000115 00000 n
0000000266 00000 n
0000000360 00000 n

trailer
<< /Size 6 /Root 1 0 R >>
startxref
441
%%EOF`

      const buffer = Buffer.from(pdfContent)
      const result = await extractTextFromBuffer(buffer, "application/pdf", "test.pdf")

      expect(result.isRight()).toBe(true)
      const text = result.orThrow()
      expect(text).toContain("[PDF:")
      expect(text).toContain("page")
      expect(text).toContain("Hello PDF")
    })

    it("should infer PDF type from extension when content type is octet-stream", async () => {
      // Create a minimal PDF
      const pdfContent = `%PDF-1.0
1 0 obj
<< /Type /Catalog /Pages 2 0 R >>
endobj

2 0 obj
<< /Type /Pages /Kids [3 0 R] /Count 1 >>
endobj

3 0 obj
<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>
endobj

4 0 obj
<< /Length 44 >>
stream
BT /F1 12 Tf 100 700 Td (Hello PDF) Tj ET
endstream
endobj

5 0 obj
<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>
endobj

xref
0 6
0000000000 65535 f
0000000009 00000 n
0000000058 00000 n
0000000115 00000 n
0000000266 00000 n
0000000360 00000 n

trailer
<< /Size 6 /Root 1 0 R >>
startxref
441
%%EOF`

      const buffer = Buffer.from(pdfContent)
      const result = await extractTextFromBuffer(buffer, "application/octet-stream", "report.pdf")

      expect(result.isRight()).toBe(true)
      expect(result.orThrow()).toContain("Hello PDF")
    })
  })

  describe("unsupported content types", () => {
    it("should return Left for unsupported binary content types", async () => {
      const buffer = Buffer.from([0x00, 0x01, 0x02])
      const result = await extractTextFromBuffer(buffer, "application/zip", "archive.zip")
      expect(result.isLeft()).toBe(true)
      result.tapLeft((err) => {
        expect(err.message).toContain("Unsupported content type")
      })
    })

    it("should return Left for image content types", async () => {
      const buffer = Buffer.from([0xff, 0xd8, 0xff])
      const result = await extractTextFromBuffer(buffer, "image/jpeg", "photo.jpg")
      expect(result.isLeft()).toBe(true)
      result.tapLeft((err) => {
        expect(err.message).toContain("Unsupported content type")
      })
    })

    it("should include supported types in error message", async () => {
      const buffer = Buffer.from([0x00])
      const result = await extractTextFromBuffer(buffer, "application/zip", "archive.zip")
      expect(result.isLeft()).toBe(true)
      result.tapLeft((err) => {
        expect(err.message).toContain("text/*")
      })
    })
  })

  describe("EXTRACTABLE_TYPES constant", () => {
    it("should include PDF, DOCX, and XLSX", () => {
      expect(EXTRACTABLE_TYPES).toContain("application/pdf")
      expect(EXTRACTABLE_TYPES).toContain("application/vnd.openxmlformats-officedocument.wordprocessingml.document")
      expect(EXTRACTABLE_TYPES).toContain("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    })
  })
})
