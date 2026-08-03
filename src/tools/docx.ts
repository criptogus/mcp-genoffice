import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { readFileSync } from 'node:fs'
import { basename, dirname, join } from 'node:path'
import { z } from 'zod'
import { loadEngine } from '../engine.js'

/**
 * Minimal structural view of the docx-engine surface we use. The engine is
 * loaded dynamically (pure-TS source from a genoffice clone), so these are
 * hand-declared interfaces, not imports.
 */
interface EngineBlock {
  id: string
  docxIndex: number
  originalXml?: string
  type: string
  styleId?: string
  rawPPr?: string
  runs?: Array<{ text?: string }>
  hidden?: boolean
}

interface DocxEngine {
  parseDocx(buf: Uint8Array): Promise<{
    blocks: EngineBlock[]
    [k: string]: unknown
  }>
  saveDocx(
    parsed: unknown,
    finalBlocks: unknown[],
    options?: Record<string, unknown>,
  ): Promise<Uint8Array>
  buildBlankDocx(options?: { eastAsiaFont?: string }): Promise<Uint8Array>
}

function blockPreview(b: EngineBlock): string {
  const text = (b.runs ?? [])
    .map((r) => r.text ?? '')
    .join('')
    .trim()
  return text.length > 120 ? text.slice(0, 120) + '…' : text
}

export function registerDocxTools(server: McpServer): void {
  server.registerTool(
    'genoffice_docx_blocks',
    {
      title: 'GenOffice docx blocks',
      description:
        'Parse a .docx with the GenOffice engine and list its top-level blocks ' +
        '(paragraphs/headings/tables) with their index, type, style and text. ' +
        'Use this BEFORE genoffice_docx_patch to pick the block indexes you want to rewrite.',
      inputSchema: {
        path: z.string().describe('Absolute path to the .docx file'),
      },
    },
    async ({ path }) => {
      try {
        const engine = await loadEngine<DocxEngine>('docx-engine')
        const parsed = await engine.parseDocx(readFileSync(path))
        const visible = parsed.blocks.filter((b) => !b.hidden)
        const lines = visible.map((b, i) => {
          const preview = blockPreview(b)
          return `${i}\t[${b.type}${b.styleId ? ` ${b.styleId}` : ''}]\tdocxIndex=${b.docxIndex}\t${preview || '(sem texto)'}`
        })
        return {
          content: [
            {
              type: 'text' as const,
              text:
                `${visible.length} blocos visíveis (${parsed.blocks.length} total)\n\n` +
                lines.join('\n'),
            },
          ],
        }
      } catch (err) {
        return {
          isError: true,
          content: [
            {
              type: 'text' as const,
              text: `genoffice_docx_blocks failed: ${(err as Error).message}. ` +
                `Verify the path is a valid .docx.`,
            },
          ],
        }
      }
    },
  )

  server.registerTool(
    'genoffice_docx_patch',
    {
      title: 'GenOffice docx patch (byte-preserving)',
      description:
        'Rewrite one or more paragraphs of a .docx using the GenOffice byte-preserving ' +
        'roundtrip: only the edited blocks are regenerated as OOXML fragments; every ' +
        'untouched block keeps its original bytes, so layout, styles, headers, comments ' +
        'and other parts survive. Paragraph formatting (styleId/rawPPr) is carried over. ' +
        'Input: file path + edits [{index, text}] where index comes from genoffice_docx_blocks ' +
        '(0-based visible order). Output: a new patched file (never modifies the original).',
      inputSchema: {
        path: z.string().describe('Absolute path to the source .docx (never modified)'),
        edits: z
          .array(
            z.object({
              index: z.number().int().min(0).describe('Block index from genoffice_docx_blocks'),
              text: z.string().describe('New paragraph text'),
            }),
          )
          .min(1)
          .describe('Paragraph rewrites to apply'),
        outPath: z
          .string()
          .optional()
          .describe('Absolute output path; defaults to <dir>/<name>.patched.docx'),
      },
    },
    async ({ path, edits, outPath }) => {
      try {
        const engine = await loadEngine<DocxEngine>('docx-engine')
        const parsed = await engine.parseDocx(readFileSync(path))
        const visible = parsed.blocks.filter((b: EngineBlock) => !b.hidden)

        const byIndex = new Map(edits.map((e) => [e.index, e.text]))
        for (const e of edits) {
          if (e.index >= visible.length) {
            return {
              isError: true,
              content: [
                {
                  type: 'text' as const,
                  text: `edit index ${e.index} out of range: file has ${visible.length} visible blocks. ` +
                    `Run genoffice_docx_blocks first to list valid indexes.`,
                },
              ],
            }
          }
          const t = visible[e.index].type
          if (t !== 'paragraph' && t !== 'heading' && t !== 'listItem') {
            return {
              isError: true,
              content: [
                {
                  type: 'text' as const,
                  text: `block ${e.index} is type "${t}" — only paragraph/heading/listItem blocks can be text-patched in this version.`,
                },
              ],
            }
          }
        }

        const finalBlocks = visible.map((b: EngineBlock, i: number) => {
          const newText = byIndex.get(i)
          if (newText === undefined) {
            return { kind: 'original', docxIndex: b.docxIndex }
          }
          return {
            kind: 'generated',
            block: {
              type: b.type === 'listItem' ? 'paragraph' : b.type,
              styleId: b.styleId ?? undefined,
              rawPPr: b.rawPPr ?? undefined,
              runs: [{ text: newText }],
            },
          }
        })

        const out = outPath ?? join(dirname(path), `${basename(path, '.docx')}.patched.docx`)
        const saved = await engine.saveDocx(parsed, finalBlocks)
        const { writeFileSync } = await import('node:fs')
        writeFileSync(out, Buffer.from(saved))

        return {
          content: [
            {
              type: 'text' as const,
              text:
                `Patched ${edits.length} block(s) → ${out}\n` +
                `- ${visible.length - edits.length} untouched blocks kept their original bytes (byte-preserving roundtrip)\n` +
                `- Original file unchanged`,
            },
          ],
        }
      } catch (err) {
        return {
          isError: true,
          content: [
            {
              type: 'text' as const,
              text: `genoffice_docx_patch failed: ${(err as Error).message}. ` +
                `Check the path, block indexes (genoffice_docx_blocks) and output directory permissions.`,
            },
          ],
        }
      }
    },
  )

  server.registerTool(
    'genoffice_docx_watermark',
    {
      title: 'GenOffice docx watermark',
      description:
        'Set a text watermark on a .docx using the GenOffice engine: the header ' +
        'part is regenerated with the watermark paragraph while the document body ' +
        'keeps its original bytes (byte-preserving roundtrip). Pass an empty ' +
        'string to remove the watermark. Writes a NEW file, never modifies the original.',
      inputSchema: {
        path: z.string().describe('Absolute path to the source .docx (never modified)'),
        text: z.string().describe('Watermark text ("" removes the watermark)'),
        outPath: z
          .string()
          .optional()
          .describe('Absolute output path; defaults to <dir>/<name>.watermarked.docx'),
      },
    },
    async ({ path, text, outPath }) => {
      try {
        const engine = await loadEngine<DocxEngine>('docx-engine')
        const parsed = await engine.parseDocx(readFileSync(path))
        const visible = parsed.blocks.filter((b: EngineBlock) => !b.hidden)
        const finalBlocks = visible.map((b: EngineBlock) => ({
          kind: 'original' as const,
          docxIndex: b.docxIndex,
        }))
        const out = outPath ?? join(dirname(path), `${basename(path, '.docx')}.watermarked.docx`)
        const saved = await engine.saveDocx(parsed, finalBlocks, { watermark: text || null })
        const { writeFileSync } = await import('node:fs')
        writeFileSync(out, Buffer.from(saved))
        return {
          content: [
            {
              type: 'text' as const,
              text:
                `${text ? `Watermark "${text}" aplicado` : 'Watermark removido'} → ${out}\n` +
                `- Corpo do documento preservado byte a byte (só o header foi regenerado)`,
            },
          ],
        }
      } catch (err) {
        return {
          isError: true,
          content: [
            {
              type: 'text' as const,
              text: `genoffice_docx_watermark failed: ${(err as Error).message}. ` +
                `Check the path is a valid .docx and the output directory is writable.`,
            },
          ],
        }
      }
    },
  )

  server.registerTool(
    'genoffice_docx_create',
    {
      title: 'GenOffice docx create',
      description:
        'Create a NEW .docx from scratch using the GenOffice engine (buildBlankDocx). ' +
        'Optionally pass paragraphs (each string becomes one paragraph; use \\n inside ' +
        'a string for line breaks within the same paragraph). Writes the file to outPath.',
      inputSchema: {
        outPath: z.string().describe('Absolute path where the new .docx will be written'),
        paragraphs: z
          .array(z.string())
          .optional()
          .describe('Initial paragraphs; omit for a blank document'),
      },
    },
    async ({ outPath, paragraphs }) => {
      try {
        const engine = await loadEngine<DocxEngine>('docx-engine')
        let bytes = await engine.buildBlankDocx()
        if (paragraphs && paragraphs.length > 0) {
          const parsed = await engine.parseDocx(bytes)
          const finalBlocks = paragraphs.map((text) => ({
            kind: 'generated' as const,
            block: { type: 'paragraph', runs: [{ text }] },
          }))
          bytes = await engine.saveDocx(parsed, finalBlocks)
        }
        const { writeFileSync } = await import('node:fs')
        writeFileSync(outPath, Buffer.from(bytes))
        return {
          content: [
            {
              type: 'text' as const,
              text:
                `Novo .docx criado → ${outPath}\n` +
                (paragraphs?.length
                  ? `- ${paragraphs.length} parágrafo(s) inicial(is)\n`
                  : '- Documento em branco\n') +
                `Use genoffice_docx_blocks + genoffice_docx_patch para editar.`,
            },
          ],
        }
      } catch (err) {
        return {
          isError: true,
          content: [
            {
              type: 'text' as const,
              text: `genoffice_docx_create failed: ${(err as Error).message}. ` +
                `Check that outPath is a writable absolute path.`,
            },
          ],
        }
      }
    },
  )

  server.registerTool(
    'genoffice_docx_delete',
    {
      title: 'GenOffice docx delete blocks',
      description:
        'Delete one or more top-level blocks (paragraphs/headings) from a .docx using ' +
        'the GenOffice byte-preserving roundtrip: the deleted blocks are spliced out, ' +
        'all remaining blocks keep their original bytes. Indexes come from ' +
        'genoffice_docx_blocks (0-based visible order). Writes a NEW file.',
      inputSchema: {
        path: z.string().describe('Absolute path to the source .docx (never modified)'),
        indexes: z
          .array(z.number().int().min(0))
          .min(1)
          .describe('Block indexes to delete (from genoffice_docx_blocks)'),
        outPath: z
          .string()
          .optional()
          .describe('Absolute output path; defaults to <dir>/<name>.deleted.docx'),
      },
    },
    async ({ path, indexes, outPath }) => {
      try {
        const engine = await loadEngine<DocxEngine>('docx-engine')
        const parsed = await engine.parseDocx(readFileSync(path))
        const visible = parsed.blocks.filter((b: EngineBlock) => !b.hidden)
        for (const i of indexes) {
          if (i >= visible.length) {
            return {
              isError: true,
              content: [
                {
                  type: 'text' as const,
                  text: `index ${i} out of range: file has ${visible.length} visible blocks. ` +
                    `Run genoffice_docx_blocks first.`,
                },
              ],
            }
          }
        }
        const drop = new Set(indexes)
        const finalBlocks = visible
          .filter((_: EngineBlock, i: number) => !drop.has(i))
          .map((b: EngineBlock) => ({ kind: 'original' as const, docxIndex: b.docxIndex }))
        const out = outPath ?? join(dirname(path), `${basename(path, '.docx')}.deleted.docx`)
        const saved = await engine.saveDocx(parsed, finalBlocks)
        const { writeFileSync } = await import('node:fs')
        writeFileSync(out, Buffer.from(saved))
        return {
          content: [
            {
              type: 'text' as const,
              text:
                `${indexes.length} bloco(s) excluído(s) → ${out}\n` +
                `- ${finalBlocks.length} bloco(s) restante(s), bytes originais preservados`,
            },
          ],
        }
      } catch (err) {
        return {
          isError: true,
          content: [
            {
              type: 'text' as const,
              text: `genoffice_docx_delete failed: ${(err as Error).message}. ` +
                `Check the path, indexes (genoffice_docx_blocks) and output directory.`,
            },
          ],
        }
      }
    },
  )
}
