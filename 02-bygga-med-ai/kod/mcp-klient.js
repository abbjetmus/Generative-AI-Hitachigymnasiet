// En MCP-klient som startar Nordverks server, listar vad den kan och använder det.
// Kör från repots rot: node 02-bygga-med-ai/kod/mcp-klient.js
import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js'
import { fileURLToPath } from 'node:url'

// Klienten startar servern som ett eget program: node mcp-server.js
const transport = new StdioClientTransport({
  command: 'node',
  args: [fileURLToPath(new URL('mcp-server.js', import.meta.url))],
})
const klient = new Client({ name: 'nordverk-klient', version: '1.0.0' })
await klient.connect(transport)

// 1. Vilka verktyg finns?
const { tools } = await klient.listTools()
console.log('Verktyg:')
for (const verktyg of tools) {
  console.log(`- ${verktyg.name}: ${verktyg.description}`)
  console.log('  inputSchema:', JSON.stringify(verktyg.inputSchema))
}

// 2. Vilka resurser finns?
const { resources } = await klient.listResources()
console.log('\nResurser:')
for (const resurs of resources) {
  console.log(`- ${resurs.uri}`)
}

// 3. Anropa ett verktyg.
const resultat = await klient.callTool({ name: 'slaUppFelkod', arguments: { kod: 'E05' } })
console.log('\nslaUppFelkod E05:\n' + resultat.content[0].text)

// 4. Läs en resurs.
const manual = await klient.readResource({ uri: 'nordverk://manualer/nv-200.md' })
console.log('\nBörjan på manualen:\n' + manual.contents[0].text.slice(0, 200))

await klient.close()
