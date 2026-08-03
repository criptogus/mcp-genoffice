import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { readFileSync, writeFileSync } from 'node:fs'
import { basename, dirname, join } from 'node:path'
import { z } from 'zod'
import { loadEngine } from '../engine.js'

/**
 * Minimal structural view of the pptx-engine surface (loaded dynamically from
 * a genoffice clone, so hand-declared — not imports).
 */
interface TextRun {
  text?: string
  [k: string]: unknown
}
interface Paragraph {
  runs: TextRun[]
  [k: string]: unknown
}
interface SlideElement {
  id: string
  type: string
  name?: string
  placeholder?: string
  text?: { paragraphs: Paragraph[] }
  dirty?: boolean
}
interface Slide {
  path: string
  originalXml: string
  elements: SlideElement[]
}
interface OpenedPptx {
  deck: { slides: Slide[]; size: { cx: number; cy: number } }
  archive: { entries: Map<string, Buffer> }
}
interface PptxEngine {
  openPptx(bytes: Uint8Array): Promise<OpenedPptx>
  savePptx(opened: OpenedPptx): Promise<Uint8Array>
  patchSlideXml(slide: Slide): string
}

function elementText(el: SlideElement): string {
  return (el.text?.paragraphs ?? [])
    .map((p) => p.runs.map((r) => r.text ?? '').join(''))
    .join('\n')
}

function preview(s: string): string {
  const one = s.replace(/\s+/g, ' ').trim()
  return one.length > 100 ? one.slice(0, 100) + '…' : one
}

const MAX_LIST = 25_000

export function registerPptxTools(server: McpServer): void {
  server.registerTool(
    'genoffice_pptx_slides',
    {
      title: 'GenOffice pptx slides',
      description:
        'Open a .pptx with the GenOffice engine and list every slide with its ' +
        'text elements (id, name, type, text preview). Use this BEFORE ' +
        'genoffice_pptx_patch to identify which slide and element to edit.',
      inputSchema: {
        path: z.string().describe('Absolute path to the .pptx file'),
      },
    },
    async ({ path }) => {
      try {
        const engine = await loadEngine<PptxEngine>('pptx-engine')
        const opened = await engine.openPptx(readFileSync(path))
        const slides = opened.deck.slides
        const w = (opened.deck.size.cx / 914400).toFixed(2)
        const h = (opened.deck.size.cy / 914400).toFixed(2)
        const out: string[] = [`Deck: ${slides.length} slide(s), ${w} x ${h} in`]
        let budget = MAX_LIST
        for (let i = 0; i < slides.length; i++) {
          const s = slides[i]
          const header = `Slide ${i + 1} (${s.path}):`
          out.push(header)
          for (const el of s.elements) {
            if (el.type !== 'text' && el.type !== 'shape') continue
            const line = `  [${el.id}] ${el.name ?? '(sem nome)'} (${el.type}): ${preview(elementText(el))}`
            out.push(line)
            budget -= line.length
            if (budget <= 0) {
              out.push('  … [truncado: use um deck menor ou edite diretamente]')
              break
            }
          }
          if (budget <= 0) break
        }
        return { content: [{ type: 'text' as const, text: out.join('\n') }] }
      } catch (err) {
        return {
          isError: true,
          content: [
            {
              type: 'text' as const,
              text: `genoffice_pptx_slides failed: ${(err as Error).message}. ` +
                `Verify the path is a valid .pptx.`,
            },
          ],
        }
      }
    },
  )

  server.registerTool(
    'genoffice_pptx_patch',
    {
      title: 'GenOffice pptx patch (byte-preserving)',
      description:
        'Replace the text of one or more text elements on a slide of a .pptx, ' +
        'using GenOffice element-level byte-preserving patching: only the edited ' +
        'elements are regenerated; every untouched element and zip part keeps its ' +
        'original bytes. The first run/paragraph of each element is used as the ' +
        'formatting template (font, size, color, alignment, bullets survive). ' +
        'Multi-line text with \\n creates one paragraph per line. ' +
        'Writes a NEW file, never modifies the original.',
      inputSchema: {
        path: z.string().describe('Absolute path to the source .pptx (never modified)'),
        slide: z.number().int().min(1).describe('Slide number (1-based) to edit'),
        edits: z
          .array(
            z.object({
              element: z
                .string()
                .describe('Element name (e.g. "Title 1") or id (from genoffice_pptx_slides)'),
              text: z.string().describe('New text (\\n = new paragraph)'),
            }),
          )
          .min(1)
          .describe('Text replacements to apply'),
        outPath: z
          .string()
          .optional()
          .describe('Absolute output path; defaults to <dir>/<name>.patched.pptx'),
      },
    },
    async ({ path, slide, edits, outPath }) => {
      try {
        const engine = await loadEngine<PptxEngine>('pptx-engine')
        const opened = await engine.openPptx(readFileSync(path))
        const slides = opened.deck.slides
        if (slide < 1 || slide > slides.length) {
          return {
            isError: true,
            content: [
              {
                type: 'text' as const,
                text: `slide ${slide} out of range: deck has ${slides.length} slide(s). ` +
                  `Run genoffice_pptx_slides to list them.`,
              },
            ],
          }
        }
        const s = slides[slide - 1]
        const done: string[] = []
        const missing: string[] = []

        for (const edit of edits) {
          const el = s.elements.find(
            (e) =>
              (e.type === 'text' || e.type === 'shape') &&
              e.text &&
              (e.name === edit.element || e.id === edit.element),
          )
          if (!el) {
            missing.push(edit.element)
            continue
          }
          const paras = el.text!.paragraphs
          const templatePara: Paragraph = paras[0] ?? { runs: [] }
          const templateRun: TextRun = templatePara.runs[0] ?? {}
          el.text!.paragraphs = edit.text.split('\n').map((line) => ({
            ...templatePara,
            runs: [{ ...templateRun, text: line }],
          }))
          el.dirty = true
          done.push(`${el.name ?? el.id} → "${preview(edit.text)}"`)
        }

        if (done.length === 0) {
          return {
            isError: true,
            content: [
              {
                type: 'text' as const,
                text: `nenhum elemento encontrado para: ${missing.join(', ')}. ` +
                  `Run genoffice_pptx_slides to list element names/ids on slide ${slide}.`,
              },
            ],
          }
        }

        // Regenerate only the edited slide and repack the archive.
        opened.archive.entries.set(s.path, Buffer.from(engine.patchSlideXml(s), 'utf8'))
        const outBytes = await engine.savePptx(opened)
        const out = outPath ?? join(dirname(path), `${basename(path, '.pptx')}.patched.pptx`)
        writeFileSync(out, Buffer.from(outBytes))

        const editedCount = done.length
        return {
          content: [
            {
              type: 'text' as const,
              text:
                `Slide ${slide}: ${editedCount} elemento(s) editado(s)\n` +
                done.map((d) => `  - ${d}`).join('\n') +
                `\n→ ${out}\n` +
                (missing.length ? `Aviso: não encontrados: ${missing.join(', ')}\n` : '') +
                `- Elementos não editados e demais zip parts preservados byte a byte (patch em nível de elemento)`,
            },
          ],
        }
      } catch (err) {
        return {
          isError: true,
          content: [
            {
              type: 'text' as const,
              text: `genoffice_pptx_patch failed: ${(err as Error).message}. ` +
                `Check the path, slide number, element names (genoffice_pptx_slides) and output directory.`,
            },
          ],
        }
      }
    },
  )
}
