// Kör samma uppgifter mot den modell som står i .env och mäter tid och kvalitet.
// Kör en gång med Aixia, byt .env till Ollama och kör igen. Jämför sedan.
//
//   node 03-ai-losningar-for-foretag/kod/jamfor-modeller.js
import { client, MODEL, EMBED_MODEL } from '../../gemensam/ai.js'

// Aixias proxy sparar svaret på en exakt likadan fråga en kort stund och
// skickar samma svar igen. Det ger falskt snabba tider, så vi ber den att
// låta bli. Andra tjänster känner inte till inställningen, så vi skickar
// den bara till Aixia.
const extra = client.baseURL.includes('aiqu.ai') ? { cache: { 'no-cache': true } } : {}

// Ett kort utdrag ur manualen till NV-200 (Nordverk AB är påhittat).
const manualutdrag = `Felkoder NV-200:
E01 Överström i motorn. Kontrollera att pumphjulet går lätt.
E02 Torrkörning, för lågt inloppstryck. Vid kallstart kan luft i systemet ge E02 som försvinner efter några minuter. Avlufta pumpen.
E04 Läckage vid axeltätningen. Stoppa pumpen och byt axeltätning.`

// Varje uppgift har en kontroll som säger om svaret är godkänt.
const uppgifter = [
  {
    namn: 'Svara utifrån manualen',
    messages: [
      { role: 'system', content: `Svara kort på svenska. Använd bara den här texten:\n${manualutdrag}` },
      { role: 'user', content: 'Pumpen visar E02 varje morgon men koden försvinner efter några minuter. Vad ska vi göra?' },
    ],
    kontroll: (svar) => /avluft/i.test(svar),
  },
  {
    namn: 'Sammanfatta ett kundärende',
    messages: [
      { role: 'system', content: 'Sammanfatta kundärendet i en mening på svenska.' },
      {
        role: 'user',
        content:
          'Vår NV-400 på linje 2 stannade igen i går med vibrationslarm. Det är andra gången den här månaden. Varje stopp kostar oss flera timmars produktion och vi funderar på att begära ersättning.',
      },
    ],
    kontroll: (svar) => /NV-400|vibration/i.test(svar) && svar.length < 400,
  },
  {
    namn: 'Klassa prioritet som JSON',
    messages: [
      { role: 'system', content: 'Bedöm prioriteten för kundärendet. akut = pumpen står still, normal = fel men pumpen går, låg = fråga.' },
      { role: 'user', content: 'Matarvattenpumpen visar E09 och har stannat. Pannan går på reservpumpen.' },
    ],
    response_format: {
      type: 'json_schema',
      json_schema: {
        name: 'prioritet',
        strict: true,
        schema: {
          type: 'object',
          additionalProperties: false,
          required: ['prioritet'],
          properties: { prioritet: { type: 'string', enum: ['låg', 'normal', 'akut'] } },
        },
      },
    },
    kontroll: (svar) => JSON.parse(svar).prioritet === 'akut',
  },
]

console.log(`Tjänst: ${client.baseURL}`)
console.log(`Modell: ${MODEL}\n`)

let godkanda = 0
let totalTid = 0

for (const uppgift of uppgifter) {
  const start = performance.now()
  const response = await client.chat.completions.create({
    model: MODEL,
    messages: uppgift.messages,
    response_format: uppgift.response_format,
    ...extra,
  })
  const sekunder = (performance.now() - start) / 1000
  totalTid += sekunder

  const svar = response.choices[0].message.content ?? ''
  let ok = false
  try {
    ok = uppgift.kontroll(svar)
  } catch {
    ok = false // till exempel om svaret inte var giltig JSON
  }
  if (ok) godkanda++

  // usage.completion_tokens = antal token som modellen skrev (inklusive resonemang)
  const tokens = response.usage?.completion_tokens ?? 0
  console.log(`${ok ? '✅' : '❌'} ${uppgift.namn}: ${sekunder.toFixed(1)} s, ${tokens} token, ${(tokens / sekunder).toFixed(0)} token/s`)
  console.log(`   ${svar.replace(/\s+/g, ' ').slice(0, 200)}\n`)
}

// Embeddings: fungerar samma kod med den lokala embeddingmodellen?
const start = performance.now()
const embedding = await client.embeddings.create({ model: EMBED_MODEL, input: 'Pumpen läcker vid axeln.' })
const embedSekunder = (performance.now() - start) / 1000
console.log(`Embedding med ${EMBED_MODEL}: ${embedding.data[0].embedding.length} dimensioner, ${embedSekunder.toFixed(1)} s\n`)

console.log(`Resultat: ${godkanda} av ${uppgifter.length} godkända, total tid ${totalTid.toFixed(1)} s`)
