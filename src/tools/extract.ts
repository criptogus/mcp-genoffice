import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { z } from 'zod'
import { loadEngine } from '../engine.js'

interface FileParse {
  parseFileToText(path: string): Promise<{ kind: string; text: string }>
}

const MAX_TEXT = 30_000

export function registerExtractTools(server: McpServer): void {
  server.registerTool(
    'genoffice_extract_text',
    {
      title: 'GenOffice extract text',
      description:
        'Extract readable text from an Office or PDF file using GenOffice engines ' +
        '(docx/xlsx/pptx/pdf). Returns the document content as markdown-ish text ' +
        'with slide/table structure. Use this to read the content of a document ' +
        'before editing it.',
      inputSchema: {
        path: z
          .string()
          .describe('Absolute path to the file (.docx, .xlsx, .pptx, .pdf)'),
      },
    },
    async ({ path }) => {
      try {
        const fp = await loadEngine<FileParse>('file-parse')
        const parsed = await fp.parseFileToText(path)
        const text =
          parsed.text.length > MAX_TEXT
            ? parsed.text.slice(0, MAX_TEXT) +
              `\n… [truncated: ${parsed.text.length - MAX_TEXT} chars. Use genoffice_docx_blocks for the full block list.]`
            : parsed.text
        return {
          content: [
            { type: 'text' as const, text: `Kind: ${parsed.kind}\n\n${text}` },
          ],
        }
      } catch (err) {
        return {
          isError: true,
          content: [
            {
              type: 'text' as const,
              text: `genoffice_extract_text failed: ${(err as Error).message}. ` +
                `Check that the path exists and is a valid Office/PDF file.`,
            },
          ],
        }
      }
    },
  )
}
