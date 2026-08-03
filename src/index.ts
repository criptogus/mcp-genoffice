import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js'
import { registerAppTools } from './tools/app.js'
import { registerDocxTools } from './tools/docx.js'
import { registerExtractTools } from './tools/extract.js'
import { registerPptxTools } from './tools/pptx.js'

const server = new McpServer({
  name: 'mcp-genoffice',
  version: '0.1.0',
})

registerExtractTools(server)
registerDocxTools(server)
registerPptxTools(server)
registerAppTools(server)

const transport = new StdioServerTransport()
await server.connect(transport)
