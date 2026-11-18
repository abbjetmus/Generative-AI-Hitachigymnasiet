// En MCP-klient där språkmodellen väljer vilket verktyg som ska köras.
// Kör från repots rot: node 02-bygga-med-ai/kod/mcp-klient-llm.js "Din fråga"
import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js'
import { fileURLToPath } from 'node:url'
import { chatta } from '../../gemensam/ai.js'

const fraga = process.argv[2] ?? 'Vad betyder E08 och har NV900-2019-0011 haft det felet?'

// 1. Starta servern och hämta listan med verktyg.
const transport = new StdioClientTransport({
  command: 'node',
  args: [fileURLToPath(new URL('mcp-server.js', import.meta.url))],
})
const klient = new Client({ name: 'nordverk-llm-klient', version: '1.0.0' })
await klient.connect(transport)
const { tools } = await klient.listTools()

// 2. Beskriv verktygen för modellen. Texten kommer från servern, inte från oss.
let beskrivning = ''
for (const verktyg of tools) {
  const parameter = Object.keys(verktyg.inputSchema.properties)[0]
  beskrivning += `- ${verktyg.name}: ${verktyg.description} Argument: ${parameter}.\n`
}

const schema = {
  type: 'object',
  properties: {
    typ: { type: 'string', enum: ['anropa', 'svar'] },
    funktion: { type: 'string', enum: [...tools.map((v) => v.name), 'ingen'] },
    argument: { type: 'string' },
    text: { type: 'string' },
  },
  required: ['typ', 'funktion', 'argument', 'text'],
  additionalProperties: false,
}

const meddelanden = [
  {
    role: 'system',
    content: `Du är Nordverks serviceassistent. Svara på svenska.
Programmet kan köra de här funktionerna åt dig:
${beskrivning}
Svara med ett JSON-objekt, ett steg i taget. Exempel:
{"typ": "anropa", "funktion": "slaUppFelkod", "argument": "E05", "text": ""}
{"typ": "svar", "funktion": "ingen", "argument": "", "text": "Ditt svar här"}`,
  },
  { role: 'user', content: fraga },
]

// 3. Samma loop som i kapitel 4, men verktygen körs av MCP-servern.
console.log('Fråga:', fraga)
for (let varv = 1; varv <= 5; varv++) {
  const text = await chatta(meddelanden, {
    response_format: { type: 'json_schema', json_schema: { name: 'steg', strict: true, schema } },
  })
  // Ibland svarar modellen med vanlig text i stället för JSON. Be den då försöka igen.
  let steg
  try {
    steg = JSON.parse(text)
  } catch {
    console.log('(Inte giltig JSON. Ber modellen försöka igen.)')
    meddelanden.push({ role: 'assistant', content: String(text) })
    meddelanden.push({ role: 'user', content: 'Svara med ett JSON-objekt som i exemplen.' })
    continue
  }

  if (steg.typ === 'svar') {
    console.log('\nSvar:', steg.text)
    break
  }

  // Bygg argumenten, t.ex. { kod: 'E08' }, och be servern köra verktyget.
  const verktyg = tools.find((v) => v.name === steg.funktion)
  const parameter = Object.keys(verktyg.inputSchema.properties)[0]
  console.log(`🔧 ${steg.funktion}({ ${parameter}: "${steg.argument}" })`)
  const resultat = await klient.callTool({
    name: steg.funktion,
    arguments: { [parameter]: steg.argument },
  })

  meddelanden.push({ role: 'assistant', content: text })
  meddelanden.push({
    role: 'user',
    content: `Resultat från ${steg.funktion}:\n${resultat.content[0].text}`,
  })
}

await klient.close()
