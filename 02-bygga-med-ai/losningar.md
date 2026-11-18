# Lösningar

Försök själv innan du tittar här! Svaren från modellen blir inte exakt
likadana varje gång, så dina svar kan skilja sig från våra.

## Kapitel 1 – Rapporter från befintliga rapporter

**1.1** Lägg till rubriken sist i `mall` i `Manadsrapport.vue`:

```js
const mall = `## Sammanfattning
(3–4 meningar om månaden)

## Återkommande problem
(punktlista: kund, pump och vad som hänt)

## Rekommendationer
(3–5 punkter)

## Risker inför november
(2–3 punkter)`
```

**1.2** `protokoll.js` läser redan fältet `tekniker`. Lägg till i
`raknaNyckeltal`, före `return`:

```js
const timmarPerTekniker = {}
for (const p of lista) {
  timmarPerTekniker[p.tekniker] = (timmarPerTekniker[p.tekniker] ?? 0) + p.timmar
}
```

och i objektet som returneras:

```js
timmarPerTekniker,
```

I `Manadsrapport.vue`, efter raden med `Uppföljning krävs`:

```js
`Timmar per tekniker: ${JSON.stringify(nyckeltal.timmarPerTekniker)}\n\n` +
```

*Svar: Sara Lind 10, Ahmed Karimi 10,5, Elin Nyström 5 och Johan Berg 18,5 timmar.*

**1.3** En funktion i `Manadsrapport.vue`:

```js
async function allaPaEnGang() {
  const allaProtokoll = protokoll.map((p) => p.text).join('\n\n')
  rapport.value = await chatta([
    {
      role: 'user',
      content: allaProtokoll + '\n\nHur många besök var det och vad kostade reservdelarna totalt?',
    },
  ])
}
```

```html
<button @click="allaPaEnGang">Alla på en gång</button>
```

*Svar: Med bara 12 korta protokoll räknade modellen rätt hos oss (12 besök,
33 260 kr). Men du kan inte veta det utan att räkna själv. Med fler protokoll
blir felen vanligare. Koden räknar rätt varje gång.*

## Kapitel 2 – RAG med sökning

**2.1** *Utan RAG gissar modellen, till exempel vad Nordverk tillverkar eller vad
en felkod betyder. Med RAG svarar den utifrån manualen och visar källan.*

**2.2** Byt ut `<li>` i listan:

```html
<li v-for="t in traffar" :key="t.kalla">
  {{ t.kalla }} (poäng {{ t.poang }})
  <details>
    <summary>Visa text</summary>
    <pre>{{ t.text }}</pre>
  </details>
</li>
```

**2.3**

```js
traffar.value = nyckelordssok(fraga.value, 2)
```

*Svar: Färre träffar ger en kortare prompt. Det fungerar när rätt avsnitt
hamnar först, men ibland försvinner avsnittet med svaret.*

## Kapitel 3 – RAG med embeddings

**3.1** *Nyckelord fungerar bra med exakta beteckningar, t.ex. ett artikelnummer
som `400-7070`. Embeddings är sämre på att skilja sådana tal åt.*

**3.2** *Med 0,3 kommer fotbollsfrågan över gränsen och modellen får orelaterad
text (den bör ändå säga att det inte står i manualerna). Med 0,7 försvinner
träffarna för E06, eftersom de hade poäng runt 0,55.*

**3.3** Ändra i `vektorsok.js`:

```js
import { serviceprotokoll } from './nordverk.js'

const allaTexter = [
  ...avsnitt,
  ...serviceprotokoll.map((p) => ({ kalla: p.fil, text: p.text })),
]
```

Byt sedan `avsnitt` mot `allaTexter` på de två ställena i `vektorsok`.

*Svar: Svartå Kraftvärme AB (fuktig lokal, korroderad kontakt på tryckgivaren).*

## Kapitel 4 – Verktygsanrop

**4.1** *Funktionen svarar "Det finns inga serviceprotokoll för ...", och
modellen säger att den inte hittar pumpen. Den ska inte hitta på historik.*

**4.2** I `verktyg.js`:

```js
import { nyckelordssok } from './sok.js'

function sokIManualer(fraga) {
  return nyckelordssok(fraga, 2)
    .map((t) => t.kalla + '\n' + t.text)
    .join('\n\n')
}

export const verktyg = { hamtaServicehistorik, slaUppFelkod, sokIManualer }
```

Lägg till en rad i `beskrivning`:

```
- sokIManualer: söker i manualerna efter tekniska data, underhåll och säkerhet. Argument: vad du vill veta.
```

*Svar: NV-900 väger 980 kg inklusive motor.*

## Kapitel 5 – Agenter

**5.1** *När du nekar får agenten resultatet "En människa sa nej". Svaret till
kunden ska då inte påstå att ett ärende är skapat. Läs svaret noga. Gör det
ändå det, är det ett exempel på varför en människa ska läsa innan något skickas.*

**5.2** Om du har gjort 4.2 finns `sokIManualer` redan i `verktyg`, och
därmed i `agentensVerktyg`. Lägg till ett steg i systemprompten:

```
2b. Frågar kunden om underhåll eller tekniska data? Sök i manualerna.
```

*Svar: Agenten bör nu söka i manualen och svara var 12:e månad för pumpen som
går dygnet runt och var 24:e månad för den andra. Kontrollera mot manualen.*

**5.3** *Agenten hinner hämta historik och slå upp felkoden, sedan står det
"Stoppade efter 2 steg." Inget ärende skapas och inget svar skrivs.*

## Kapitel 6 – MCP

**6.1** Byt anropet i `mcp-klient.js`:

```js
const resultat = await klient.callTool({
  name: 'hamtaServicehistorik',
  arguments: { serienummer: 'NV200-2023-0390' },
})
```

**6.2** I `mcp-server.js`, efter de andra verktygen:

```js
const kundarenden = JSON.parse(readFileSync(new URL('kundarenden.json', mapp), 'utf8'))

server.registerTool(
  'hamtaKundarende',
  {
    description: 'Hämtar ett kundärende.',
    inputSchema: { id: z.string().describe('Ärendets id, t.ex. K-103') },
  },
  async ({ id }) => {
    const arende = kundarenden.find((a) => a.id === id)
    const text = arende ? JSON.stringify(arende) : 'Det finns inget ärende ' + id
    return { content: [{ type: 'text', text }] }
  },
)
```

**6.3** *Modellen anropar `hamtaKundarende("K-103")` och sedan
`slaUppFelkod("E02")`. Klienten hämtar verktygslistan från servern, så det nya
verktyget kommer med automatiskt. Det är hela poängen med MCP.*
