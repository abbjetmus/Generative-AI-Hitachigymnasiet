// En MCP-server som ger AI-program tillgång till Nordverks data.
// Starta inte den här filen själv. Klienten startar den (se mcp-klient.js).
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js'
import { z } from 'zod'
import { readFileSync, readdirSync } from 'node:fs'

// Läs in datan från data/nordverk.
const mapp = new URL('../../data/nordverk/', import.meta.url)
function lasMapp(namn) {
  const filer = readdirSync(new URL(namn, mapp))
  return filer.map((fil) => ({ fil, text: readFileSync(new URL(namn + fil, mapp), 'utf8') }))
}
const manualer = lasMapp('manualer/')
const protokoll = lasMapp('serviceprotokoll/')

const server = new McpServer({ name: 'nordverk', version: '1.0.0' })

// Verktyg 1: servicehistorik för en pump.
server.registerTool(
  'hamtaServicehistorik',
  {
    description: 'Hämtar alla serviceprotokoll för en pump.',
    inputSchema: { serienummer: z.string().describe('Pumpens serienummer, t.ex. NV400-2022-0118') },
  },
  async ({ serienummer }) => {
    const traffar = protokoll.filter((p) => p.text.includes(`| Serienummer | ${serienummer} |`))
    const text =
      traffar.length > 0
        ? traffar.map((p) => p.text).join('\n\n')
        : 'Det finns inga serviceprotokoll för ' + serienummer
    return { content: [{ type: 'text', text }] }
  },
)

// Verktyg 2: slå upp en felkod i alla manualer.
server.registerTool(
  'slaUppFelkod',
  {
    description: 'Slår upp en felkod i manualerna.',
    inputSchema: { kod: z.string().describe('Felkoden, t.ex. E05') },
  },
  async ({ kod }) => {
    const rader = []
    for (const manual of manualer) {
      const rad = manual.text.split('\n').find((r) => r.startsWith(`| ${kod} |`))
      if (rad) rader.push(manual.fil + ': ' + rad)
    }
    const text = rader.length > 0 ? rader.join('\n') : 'Felkoden ' + kod + ' finns inte i manualerna.'
    return { content: [{ type: 'text', text }] }
  },
)

// Resurser: manualerna, som en klient kan läsa direkt.
for (const manual of manualer) {
  const uri = 'nordverk://manualer/' + manual.fil
  server.registerResource(
    manual.fil,
    uri,
    { description: 'Manual för pumpen ' + manual.fil.replace('.md', ''), mimeType: 'text/markdown' },
    async () => ({ contents: [{ uri, text: manual.text }] }),
  )
}

// Prata med klienten via stdin och stdout.
await server.connect(new StdioServerTransport())
// stdout används av protokollet. Skriv därför meddelanden till stderr.
console.error('Nordverks MCP-server är igång.')
