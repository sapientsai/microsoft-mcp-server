import { extname } from "node:path"

import ExcelJS from "exceljs"
import { type Either, Left, Right } from "functype"
import mammoth from "mammoth"
import { extractText as extractPdfText, getDocumentProxy } from "unpdf"

import { type GraphError, graphError } from "../errors.js"
import { isTextContent } from "./download.js"

export const CONTENT_TYPE_MAP: Record<string, string> = {
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

export function resolveContentType(contentType: string, filename: string): string {
  const lower = contentType.toLowerCase()
  if (lower === "application/octet-stream" || lower === "") {
    const ext = extname(filename).toLowerCase()
    return CONTENT_TYPE_MAP[ext] ?? contentType
  }
  return lower
}

export function resolveUploadContentType(explicit: string | undefined, filename: string): string {
  if (explicit) return explicit
  const ext = extname(filename).toLowerCase()
  return CONTENT_TYPE_MAP[ext] ?? "application/octet-stream"
}

export async function extractTextFromBuffer(
  buffer: Buffer,
  contentType: string,
  filename: string,
): Promise<Either<GraphError, string>> {
  const resolved = resolveContentType(contentType, filename)

  // Text-based types: return directly
  if (isTextContent(resolved)) {
    return Right(buffer.toString("utf-8"))
  }

  // PDF
  if (resolved === "application/pdf") {
    try {
      const pdf = await getDocumentProxy(new Uint8Array(buffer))
      try {
        const { totalPages, text } = await extractPdfText(pdf, { mergePages: true })
        return Right(`[PDF: ${totalPages} page${totalPages === 1 ? "" : "s"}]\n\n${text}`)
      } finally {
        await pdf.destroy()
      }
    } catch (err) {
      return Left(graphError("pdf_extraction_failed", err instanceof Error ? err.message : "PDF extraction failed", 0))
    }
  }

  // DOCX
  if (resolved === "application/vnd.openxmlformats-officedocument.wordprocessingml.document") {
    try {
      const result = await mammoth.extractRawText({ buffer })
      return Right(result.value)
    } catch (err) {
      return Left(
        graphError("docx_extraction_failed", err instanceof Error ? err.message : "DOCX extraction failed", 0),
      )
    }
  }

  // XLSX
  if (resolved === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") {
    try {
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
      return Right(parts.join("\n\n"))
    } catch (err) {
      return Left(
        graphError("xlsx_extraction_failed", err instanceof Error ? err.message : "XLSX extraction failed", 0),
      )
    }
  }

  const supported = [...EXTRACTABLE_TYPES, "text/*"].join(", ")
  return Left(
    graphError(
      "unsupported_type",
      `Unsupported content type "${contentType}" for text extraction. Supported: ${supported}`,
      0,
    ),
  )
}
