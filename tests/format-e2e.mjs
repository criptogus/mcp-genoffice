// E2E: formatação via MCP (server atualizado do conector + client real)
import { spawnSync } from 'node:child_process'
import { readFileSync } from 'node:fs'
import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js'

const cwd = '/Users/gustavocaetano/Projects/mcp-genoffice'
const env = { ...process.env, GENOFFICE_SRC: '/tmp/genoffice' }

const transport = new StdioClientTransport({
  command: 'node',
  args: ['bin/mcp-genoffice.mjs'],
  cwd,
  env,
})
const client = new Client({ name: 'format-e2e', version: '1.0.0' })
await client.connect(transport)

const out = '/tmp/fmt-mcp.docx'
const res = await client.callTool({
  name: 'genoffice_docx_create',
  arguments: {
    outPath: out,
    paragraphs: [
      { text: 'Título violeta bold 22pt', bold: true, color: '#6E4FF6', sizePt: 22, font: 'Inter' },
      'Corpo normal sem formatação',
      { text: 'Destaque vermelho itálico', italic: true, color: '#C00000', sizePt: 16 },
    ],
  },
})
console.log('TOOL:', JSON.stringify(res.content ?? res).slice(0, 220))

spawnSync('rm', ['-rf', '/tmp/fmt-mcp-x'])
spawnSync('mkdir', ['-p', '/tmp/fmt-mcp-x'])
spawnSync('unzip', ['-q', out, '-d', '/tmp/fmt-mcp-x'])
const xml = readFileSync('/tmp/fmt-mcp-x/word/document.xml', 'utf8')

const checks = {
  bold: xml.includes('<w:b/>') || xml.includes('<w:b '),
  violet: xml.includes('6E4FF6'),
  size44: xml.includes('w:val="44"'), // 22pt → 44 half-points
  arial: xml.includes('Inter'),
  red: xml.includes('C00000'),
  italic: xml.includes('<w:i/>') || xml.includes('<w:i '),
  size32: xml.includes('w:val="32"'), // 16pt → 32
}
console.log('CHECKS:', JSON.stringify(checks))
const ok = Object.values(checks).every(Boolean)
console.log(ok ? 'MCP FORMATACAO OK' : 'MCP FORMATACAO FALHOU')
await client.close()
process.exit(ok ? 0 : 1)
