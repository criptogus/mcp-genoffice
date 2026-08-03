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
  structureDirty?: boolean
}
interface OpenedPptx {
  deck: { slides: Slide[]; size: { cx: number; cy: number } }
  archive: { entries: Map<string, Buffer> }
}
interface PptxEngine {
  openPptx(bytes: Uint8Array): Promise<OpenedPptx>
  savePptx(opened: OpenedPptx): Promise<Uint8Array>
  patchSlideXml(slide: Slide): string
  createBlankPptx(): Promise<Uint8Array>
  deleteElement(slide: Slide, elementId: string): boolean
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

  server.registerTool(
    'genoffice_pptx_create',
    {
      title: 'GenOffice pptx create',
      description:
        'Create a NEW .pptx from scratch using the GenOffice engine (createBlankPptx): ' +
        'a deck with one blank slide. Use genoffice_pptx_patch afterwards to fill in ' +
        'titles and content. Writes the file to outPath.',
      inputSchema: {
        outPath: z.string().describe('Absolute path where the new .pptx will be written'),
      },
    },
    async ({ outPath }) => {
      try {
        const engine = await loadEngine<PptxEngine>('pptx-engine')
        const bytes = await engine.createBlankPptx()
        writeFileSync(outPath, Buffer.from(bytes))
        return {
          content: [
            {
              type: 'text' as const,
              text:
                `Novo .pptx criado → ${outPath}\n` +
                `Use genoffice_pptx_slides + genoffice_pptx_patch para preencher.`,
            },
          ],
        }
      } catch (err) {
        return {
          isError: true,
          content: [
            {
              type: 'text' as const,
              text: `genoffice_pptx_create failed: ${(err as Error).message}. ` +
                `Check that outPath is a writable absolute path.`,
            },
          ],
        }
      }
    },
  )

  server.registerTool(
    'genoffice_pptx_delete',
    {
      title: 'GenOffice pptx delete elements',
      description:
        'Delete one or more text/shape elements from a slide of a .pptx using the ' +
        'GenOffice engine (element-level byte-preserving: the edited slide is rebuilt, ' +
        'every other slide and zip part keeps its original bytes). Elements are ' +
        'matched by name or id from genoffice_pptx_slides. Writes a NEW file.',
      inputSchema: {
        path: z.string().describe('Absolute path to the source .pptx (never modified)'),
        slide: z.number().int().min(1).describe('Slide number (1-based)'),
        elements: z
          .array(z.string())
          .min(1)
          .describe('Element names or ids to delete (from genoffice_pptx_slides)'),
        outPath: z
          .string()
          .optional()
          .describe('Absolute output path; defaults to <dir>/<name>.deleted.pptx'),
      },
    },
    async ({ path, slide, elements, outPath }) => {
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
        const deleted: string[] = []
        const missing: string[] = []
        for (const target of elements) {
          const el = s.elements.find((e) => e.name === target || e.id === target)
          if (!el) {
            missing.push(target)
            continue
          }
          if (engine.deleteElement(s, el.id)) {
            s.structureDirty = true
            deleted.push(el.name ?? el.id)
          } else {
            missing.push(target)
          }
        }
        if (deleted.length === 0) {
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
        opened.archive.entries.set(s.path, Buffer.from(engine.patchSlideXml(s), 'utf8'))
        const outBytes = await engine.savePptx(opened)
        const out = outPath ?? join(dirname(path), `${basename(path, '.pptx')}.deleted.pptx`)
        writeFileSync(out, Buffer.from(outBytes))
        return {
          content: [
            {
              type: 'text' as const,
              text:
                `Slide ${slide}: ${deleted.length} elemento(s) excluído(s)\n` +
                deleted.map((d) => `  - ${d}`).join('\n') +
                `\n→ ${out}\n` +
                (missing.length ? `Aviso: não encontrados: ${missing.join(', ')}\n` : '') +
                `- Demais slides e zip parts preservados byte a byte`,
            },
          ],
        }
      } catch (err) {
        return {
          isError: true,
          content: [
            {
              type: 'text' as const,
              text: `genoffice_pptx_delete failed: ${(err as Error).message}. ` +
                `Check the path, slide number, element names (genoffice_pptx_slides) and output directory.`,
            },
          ],
        }
      }
    },
  )
}
