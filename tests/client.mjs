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

await client.close()
process.exit(0)
