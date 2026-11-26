# 6. MCP

I kapitel 4 och 5 skrev du verktygen direkt i Vue-appen. Vill Nordverk använda
samma verktyg i ett annat AI-program, till exempel en kodassistent eller en
chattapp, måste de skrivas om där. **MCP** löser det.

## Vad är MCP?

**Model Context Protocol** är en öppen standard för hur AI-program hittar och
använder verktyg och data. Man skriver verktygen en gång, i en **MCP-server**.
Sedan kan alla program som förstår MCP använda dem. Jämför med USB: samma
kontakt passar många olika apparater.

Enligt [MCP:s dokumentation](https://modelcontextprotocol.io/docs/learn/architecture)
finns tre roller:

| Roll | Vad den gör | Hos oss |
|---|---|---|
| Värd (host) | AI-programmet som användaren använder | Vårt Node-skript |
| Klient (client) | Den del av värden som pratar med en server | `Client` i skripten |
| Server | Erbjuder verktyg och data | `mcp-server.js` |

En server kan erbjuda tre sorters saker:

| Sak | Betyder | Hos oss |
|---|---|---|
| Tools (verktyg) | Funktioner som klienten kan anropa | `hamtaServicehistorik`, `slaUppFelkod` |
| Resources (resurser) | Data som klienten kan läsa | Manualerna |
| Prompts | Färdiga promptmallar | Används inte här |

Klienten och servern pratar med JSON-meddelanden. Det enklaste sättet heter
**stdio**: klienten startar servern som ett eget program och de skriver till
varandra via in- och utdata, som i en terminal. Det finns också ett sätt över
HTTP för servrar på nätet.

💡 Här använder vi Node i stället för Vue. En MCP-server är ett eget program som
läser filer från disken. Det kan inte köras i en webbläsare.

## Servern

Paketet `@modelcontextprotocol/sdk` är redan installerat i repots rot. Filen
[`kod/mcp-server.js`](kod/mcp-server.js):

```js
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
```

- `registerTool` registrerar ett verktyg: namn, beskrivning, `inputSchema` och
  funktionen som körs. Schemat skrivs med paketet `zod`. SDK:n gör om det till
  JSON-schema åt klienten.
- `registerResource` registrerar en resurs med en adress (URI), t.ex.
  `nordverk://manualer/nv-200.md`.

⚠️ I en stdio-server får du inte använda `console.log`. Utdatan är reserverad
för MCP-meddelandena. Skriv till `console.error` i stället.

## En klient som listar och anropar

Filen [`kod/mcp-klient.js`](kod/mcp-klient.js) startar servern, frågar vad den
kan och använder ett verktyg och en resurs:

```js
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
```

Kör från repots rot:

```
node 02-bygga-med-ai/kod/mcp-klient.js
```

Utdata hos oss (förkortad):

```
Verktyg:
- hamtaServicehistorik: Hämtar alla serviceprotokoll för en pump.
  inputSchema: {"type":"object","properties":{"serienummer":{"type":"string", ...
- slaUppFelkod: Slår upp en felkod i manualerna.

Resurser:
- nordverk://manualer/nv-200.md
- nordverk://manualer/nv-400.md
- nordverk://manualer/nv-900.md

slaUppFelkod E05:
nv-400.md: | E05 | Vibration över 4,5 mm/s | Stoppa pumpen så snart ...
nv-900.md: | E05 | Vibration över 4,5 mm/s | Kontrollera uppriktning och lager. |
```

Lägg märke till `inputSchema`. Det är samma sorts JSON-schema som i `tools` i
kapitel 4.

## En klient där modellen väljer verktyg

Nu kopplar vi ihop MCP med språkmodellen. Loopen är samma som i kapitel 4. Två
saker är nya:

1. Beskrivningen av verktygen kommer från **servern** (`listTools`). Lägger
   någon till ett verktyg i servern, ser modellen det direkt.
2. Verktygen körs av servern med `callTool`.

Filen [`kod/mcp-klient-llm.js`](kod/mcp-klient-llm.js):

```js
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
```

```
node 02-bygga-med-ai/kod/mcp-klient-llm.js "Vad betyder E08 och har NV900-2019-0011 haft det felet?"
```

Hos oss:

```
🔧 slaUppFelkod({ kod: "E08" })
🔧 hamtaServicehistorik({ serienummer: "NV900-2019-0011" })

Svar: Felkoden E08 betyder att lagertemperaturen är över 85 °C ... Pumpen har
haft felkoden E08 vid två tillfällen och har redan åtgärdats med lagerbyte.
```

## Använda servern i andra program

Samma server kan användas av andra MCP-värdar, till exempel kodassistenter i
VS Code. Hur man lägger till en server skiljer sig mellan programmen och ändras
ofta. Se [VS Code-dokumentationen om MCP-servrar](https://code.visualstudio.com/docs/copilot/customization/mcp-servers).

Vill du testa servern utan egen klient finns
[MCP Inspector](https://modelcontextprotocol.io/docs/tools/inspector), som
öppnas i webbläsaren. Den kräver en ny version av Node (se sidan):

```
npx @modelcontextprotocol/inspector node 02-bygga-med-ai/kod/mcp-server.js
```

⚠️ En MCP-server får tillgång till din data och kan köra kod på din dator.
Installera bara servrar du litar på. Verktygens beskrivningar hamnar i
modellens prompt, så en elak server kan försöka styra modellen.

## Sammanfattning

| Begrepp | Betyder |
|---|---|
| MCP | Standard för hur AI-program använder verktyg och data |
| MCP-server | Program som erbjuder verktyg, resurser och prompts |
| MCP-klient | Del av AI-programmet som pratar med en server |
| stdio | Klienten startar servern och de pratar via in- och utdata |
| `listTools` / `callTool` | Fråga vilka verktyg som finns / kör ett verktyg |

Nästa kapitel: [Agenter i Microsoft 365](07-agenter-i-microsoft-365.md)
