// Live test of the CDP app mode against the installed GenOffice app (macOS).
// Exercises: status → launch (with CDP port) → eval → screenshot → open_file.
import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js'
import { execFileSync } from 'node:child_process'
import { existsSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))
const projectRoot = join(here, '..')
const FIXTURE = '/tmp/genoffice/fixtures/generated/simple.docx'

const transport = new StdioClientTransport({
  command: process.execPath,
  args: ['--import', 'tsx', join(projectRoot, 'src', 'index.ts')],
  env: { ...process.env, GENOFFICE_SRC: '/tmp/genoffice' },
  cwd: projectRoot,
})
const client = new Client({ name: 'mcp-genoffice-app-test', version: '0.0.1' })
await client.connect(transport)

const tools = await client.listTools()
console.log('TOOLS (inclui app):', tools.tools.map((t) => t.name).filter((n) => n.startsWith('genoffice_app_')).join(', '))

// 1. status (before launch)
const st0 = await client.callTool({ name: 'genoffice_app_status', arguments: {} })
console.log('\n--- status (antes) ---')
console.log(st0.content[0].text)
console.log(st0.isError ? '❌ ERROR' : '✅ OK')

// 2. launch with CDP
const launch = await client.callTool({ name: 'genoffice_app_launch', arguments: { killExisting: true } })
console.log('\n--- genoffice_app_launch ---')
console.log(launch.content[0].text)
console.log(launch.isError ? '❌ ERROR' : '✅ OK')

// 3. eval: DOM probe
const ev = await client.callTool({
  name: 'genoffice_app_eval',
  arguments: { expression: 'JSON.stringify({title: document.title, buttons: document.body.innerText.slice(0,120)})' },
})
console.log('\n--- genoffice_app_eval ---')
console.log(ev.content[0].text.slice(0, 300))
console.log(ev.isError ? '❌ ERROR' : '✅ OK')

// 4. screenshot
const shot = await client.callTool({ name: 'genoffice_app_screenshot', arguments: { outPath: '/tmp/genoffice-cdp.png' } })
console.log('\n--- genoffice_app_screenshot ---')
console.log(shot.content[0].text)
console.log(shot.isError ? '❌ ERROR' : `✅ OK (arquivo: ${existsSync('/tmp/genoffice-cdp.png') ? 'existe' : 'FALTANDO'})`)

// 5. open file (macOS `open -a`; the app registers document types)
const of = await client.callTool({ name: 'genoffice_app_open_file', arguments: { path: FIXTURE } })
console.log('\n--- genoffice_app_open_file ---')
console.log(of.content[0].text)
console.log(of.isError ? '❌ ERROR' : '✅ OK')

// 6. status after opening the file — the shell may now show the doc tab
await new Promise((r) => setTimeout(r, 4000))
const st1 = await client.callTool({ name: 'genoffice_app_status', arguments: {} })
console.log('\n--- status (após abrir arquivo) ---')
console.log(st1.content[0].text)
console.log(st1.isError ? '❌ ERROR' : '✅ OK')

await client.close()
process.exit(0)
