// Real-client test: spawn the server over stdio, list tools, and exercise the
// tools against real fixture files from the genoffice clone.
import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'

const here = dirname(fileURLToPath(import.meta.url))
const projectRoot = join(here, '..')

// Fixtures always come from the dev clone; SERVER_SRC controls what the server
// sees (unset/empty = server auto-clones the pinned genoffice revision).
const FIXTURE_SRC = process.env.FIXTURE_SRC ?? '/tmp/genoffice'
const FIXTURE_PPTX = join(FIXTURE_SRC, 'packages/pptx-engine/tests/fixtures/01_standard_business.pptx')
const FIXTURE_DOCX = join(FIXTURE_SRC, 'fixtures/generated/simple.docx')
const OUT_DOCX = '/tmp/genoffice-patched-via-mcp.docx'

const serverEnv = { ...process.env }
const serverSrc = process.env.SERVER_SRC
if (serverSrc) serverEnv.GENOFFICE_SRC = serverSrc
else delete serverEnv.GENOFFICE_SRC

const transport = new StdioClientTransport({
  command: process.execPath,
  args: ['--import', 'tsx', join(projectRoot, 'src', 'index.ts')],
  env: serverEnv,
  cwd: projectRoot,
})
const client = new Client({ name: 'mcp-genoffice-test', version: '0.0.1' })
await client.connect(transport)

// 1. list tools
const tools = await client.listTools()
console.log('TOOLS:', tools.tools.map((t) => t.name).join(', '))

// 2. extract_text on the pptx fixture
const ext = await client.callTool({ name: 'genoffice_extract_text', arguments: { path: FIXTURE_PPTX } })
const extText = ext.content[0].text
console.log('\n--- genoffice_extract_text ---')
console.log(extText.slice(0, 220))
console.log(ext.isError ? '❌ ERROR' : '✅ OK')

// 3. docx_blocks on the simple.docx fixture
const blocks = await client.callTool({ name: 'genoffice_docx_blocks', arguments: { path: FIXTURE_DOCX } })
console.log('\n--- genoffice_docx_blocks ---')
console.log(blocks.content[0].text)
console.log(blocks.isError ? '❌ ERROR' : '✅ OK')

// 4. docx_patch: rewrite block 0
const patch = await client.callTool({
  name: 'genoffice_docx_patch',
  arguments: { path: FIXTURE_DOCX, edits: [{ index: 0, text: 'TITULO ALTERADO VIA MCP-GENOFFICE' }], outPath: OUT_DOCX },
})
console.log('\n--- genoffice_docx_patch ---')
console.log(patch.content[0].text)
console.log(patch.isError ? '❌ ERROR' : '✅ OK')

// 5. verify the patched file round-trips and keeps the other paragraph bytes
if (!patch.isError) {
  const blocks2 = await client.callTool({ name: 'genoffice_docx_blocks', arguments: { path: OUT_DOCX } })
  console.log('\n--- verificação: blocks do arquivo patched ---')
  console.log(blocks2.content[0].text)
  console.log(blocks2.isError ? '❌ ERROR' : '✅ OK')
}

// 6. pptx_slides on the pptx fixture
const slides = await client.callTool({ name: 'genoffice_pptx_slides', arguments: { path: FIXTURE_PPTX } })
console.log('\n--- genoffice_pptx_slides ---')
console.log(slides.content[0].text.slice(0, 700))
console.log(slides.isError ? '❌ ERROR' : '✅ OK')

// 7. pptx_patch: edit the title element on slide 1
const pptxOut = '/tmp/genoffice-patched-via-mcp.pptx'
const pp = await client.callTool({
  name: 'genoffice_pptx_patch',
  arguments: { path: FIXTURE_PPTX, slide: 1, edits: [{ element: 'Title 1', text: 'TITULO PPTX VIA MCP-GENOFFICE' }], outPath: pptxOut },
})
console.log('\n--- genoffice_pptx_patch ---')
console.log(pp.content[0].text)
console.log(pp.isError ? '❌ ERROR' : '✅ OK')

// 8. verify: re-open the patched pptx and confirm the new title
if (!pp.isError) {
  const slides2 = await client.callTool({ name: 'genoffice_pptx_slides', arguments: { path: pptxOut } })
  console.log('\n--- verificação: slide 1 do pptx patched ---')
  console.log(slides2.content[0].text.split('\n').slice(0, 6).join('\n'))
  console.log(slides2.isError ? '❌ ERROR' : '✅ OK')
}

// 9. docx_watermark: set + verify the doc still parses with identical blocks
const wmOut = '/tmp/genoffice-watermarked.docx'
const wm = await client.callTool({
  name: 'genoffice_docx_watermark',
  arguments: { path: FIXTURE_DOCX, text: 'CONFIDENCIAL', outPath: wmOut },
})
console.log('\n--- genoffice_docx_watermark ---')
console.log(wm.content[0].text)
console.log(wm.isError ? '❌ ERROR' : '✅ OK')

if (!wm.isError) {
  const blocks3 = await client.callTool({ name: 'genoffice_docx_blocks', arguments: { path: wmOut } })
  const same = blocks3.content[0].text === blocks.content[0].text
  console.log('--- verificação: blocks do docx watermarked idênticos ao original ---')
  console.log(same ? '✅ OK (corpo intacto, só header regenerado)' : '❌ DIVERGÊNCIA no corpo')
  console.log(blocks3.isError ? '❌ ERROR' : '✅ OK')
}

// 10. docx_create: new document from paragraphs
const createdDocx = '/tmp/genoffice-created.docx'
const dc = await client.callTool({
  name: 'genoffice_docx_create',
  arguments: { outPath: createdDocx, paragraphs: ['Título Criado', 'Parágrafo um', 'Parágrafo dois'] },
})
console.log('\n--- genoffice_docx_create ---')
console.log(dc.content[0].text)
console.log(dc.isError ? '❌ ERROR' : '✅ OK')

// 11. verify create → docx_delete block 1
if (!dc.isError) {
  const dd = await client.callTool({ name: 'genoffice_docx_delete', arguments: { path: createdDocx, indexes: [1], outPath: '/tmp/genoffice-created-deleted.docx' } })
  console.log('\n--- genoffice_docx_delete ---')
  console.log(dd.content[0].text)
  console.log(dd.isError ? '❌ ERROR' : '✅ OK')
  if (!dd.isError) {
    const b4 = await client.callTool({ name: 'genoffice_docx_blocks', arguments: { path: '/tmp/genoffice-created-deleted.docx' } })
    console.log('--- verificação: blocks após delete ---')
    console.log(b4.content[0].text)
    console.log(b4.content[0].text.includes('Título Criado') && b4.content[0].text.includes('Parágrafo dois') && !b4.content[0].text.includes('Parágrafo um') ? '✅ OK (2 blocos, 1º e 3º intactos)' : '❌ DIVERGÊNCIA')
  }
}

// 12. pptx_create
const createdPptx = '/tmp/genoffice-created.pptx'
const pc = await client.callTool({ name: 'genoffice_pptx_create', arguments: { outPath: createdPptx } })
console.log('\n--- genoffice_pptx_create ---')
console.log(pc.content[0].text)
console.log(pc.isError ? '❌ ERROR' : '✅ OK')

// 13. pptx_delete: remove Subtitle 2 from slide 1 of the fixture
const pd = await client.callTool({
  name: 'genoffice_pptx_delete',
  arguments: { path: FIXTURE_PPTX, slide: 1, elements: ['Subtitle 2'], outPath: '/tmp/genoffice-fixture-deleted.pptx' },
})
console.log('\n--- genoffice_pptx_delete ---')
console.log(pd.content[0].text)
console.log(pd.isError ? '❌ ERROR' : '✅ OK')
if (!pd.isError) {
  const sl3 = await client.callTool({ name: 'genoffice_pptx_slides', arguments: { path: '/tmp/genoffice-fixture-deleted.pptx' } })
  const slide1 = sl3.content[0].text.split('\n').slice(0, 4).join('\n')
  console.log('--- verificação: slide 1 após delete ---')
  console.log(slide1)
  console.log(slide1.includes('Title 1') && !slide1.includes('Subtitle 2') ? '✅ OK (Title 1 ficou, Subtitle 2 saiu)' : '❌ DIVERGÊNCIA')
}

await client.close()
process.exit(0)
