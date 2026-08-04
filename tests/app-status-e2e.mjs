// E2E: genoffice_app_status reporta o app detectado (HermesOffice no fork)
import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js'

const cwd = '/Users/gustavocaetano/Projects/mcp-genoffice'
const transport = new StdioClientTransport({
  command: 'node',
  args: ['bin/mcp-genoffice.mjs'],
  cwd,
  env: { ...process.env, GENOFFICE_SRC: '/tmp/genoffice' },
})
const client = new Client({ name: 'app-status-e2e', version: '1.0.0' })
await client.connect(transport)
const res = await client.callTool({ name: 'genoffice_app_status', arguments: {} })
const text = JSON.stringify(res.content ?? res)
console.log('STATUS:', text.slice(0, 220))
const ok = text.includes('HermesOffice') && text.includes('v0.4')
console.log(ok ? 'APP DETECT OK' : 'APP DETECT FALHOU')
await client.close()
process.exit(ok ? 0 : 1)
