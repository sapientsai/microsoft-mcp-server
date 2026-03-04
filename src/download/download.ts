import { mkdir, writeFile } from "node:fs/promises"
import { basename } from "node:path"
import { join } from "node:path"

import { imageContent } from "fastmcp"
import { Option } from "functype"

export const TEXT_MIME_PREFIXES = ["text/", "application/json", "application/xml", "application/csv"]
export const TEXT_MIME_SUFFIXES = ["+xml", "+json"]

export function isTextContent(contentType: string): boolean {
  const lower = contentType.toLowerCase()
  return (
    TEXT_MIME_PREFIXES.some((prefix) => lower.startsWith(prefix)) ||
    TEXT_MIME_SUFFIXES.some((suffix) => lower.includes(suffix))
  )
}

export function filenameFromHeaders(headers: Headers): Option<string> {
  const disposition = headers.get("content-disposition")
  if (!disposition) return Option<string>(undefined)
  const match = /filename\*?=(?:UTF-8''|"?)([^";]+)"?/i.exec(disposition)
  return Option(match?.[1]).map(decodeURIComponent)
}

export function filenameFromPath(path: string): Option<string> {
  const colonPathMatch = /:\/([^:]+):\/content/i.exec(path)
  return Option(colonPathMatch?.[1]).map(basename)
}

export function filenameFromUrl(url: string): Option<string> {
  try {
    const { pathname } = new URL(url)
    const name = basename(pathname)
    return name && name !== "/" ? Option(decodeURIComponent(name)) : Option<string>(undefined)
  } catch {
    return Option<string>(undefined)
  }
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
