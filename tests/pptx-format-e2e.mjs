// E2E: formatação pptx via MCP (patch com bold/color/sizePt/font)
import { spawnSync } from 'node:child_process'
import { readFileSync } from 'node:fs'
import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js'

const cwd = '/Users/gustavocaetano/Projects/mcp-genoffice'
const env = { ...process.env, GENOFFICE_SRC: '/tmp/genoffice' }

const transport = new StdioClientTransport({ command: 'node', args: ['bin/mcp-genoffice.mjs'], cwd, env })
const client = new Client({ name: 'pptx-fmt-e2e', version: '1.0.0' })
await client.connect(transport)

const out = '/tmp/fmt-mcp-deck.pptx'
const res = await client.callTool({
  name: 'genoffice_pptx_patch',
  arguments: {
    path: '/tmp/genoffice/templates/deck-base.pptx',
    slide: 1,
    edits: [
      { element: 'Title 1', text: 'Deck formatado', bold: true, color: '#6E4FF6', sizePt: 40, font: 'Inter' },
      { element: 'Subtitle 2', text: 'Subtítulo em itálico', italic: true, color: '#C00000', sizePt: 20 },
    ],
    outPath: out,
  },
})
console.log('TOOL:', JSON.stringify(res.content ?? res).slice(0, 200))

spawnSync('rm', ['-rf', '/tmp/fmt-mcp-deck-x'])
spawnSync('mkdir', ['-p', '/tmp/fmt-mcp-deck-x'])
spawnSync('unzip', ['-q', out, '-d', '/tmp/fmt-mcp-deck-x'])
const xml = readFileSync('/tmp/fmt-mcp-deck-x/ppt/slides/slide1.xml', 'utf8')

const checks = {
  bold: xml.includes('b="1"') || xml.includes('<a:b/>'),
  violet: xml.includes('6E4FF6'),
  size40: xml.includes('sz="4000"'), // 40pt → 4000 centipoints
  arial: xml.includes('Inter'),
  italic: xml.includes('i="1"') || xml.includes('<a:i/>'),
  red: xml.includes('C00000'),
  size20: xml.includes('sz="2000"'),
}
console.log('CHECKS:', JSON.stringify(checks))
const ok = Object.values(checks).every(Boolean)
console.log(ok ? 'PPTX FORMATACAO OK' : 'PPTX FORMATACAO FALHOU')
await client.close()
process.exit(ok ? 0 : 1)
