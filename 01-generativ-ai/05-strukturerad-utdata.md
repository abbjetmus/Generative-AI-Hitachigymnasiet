# 5. Strukturerad utdata

Nordverks tekniker skriver serviceprotokoll i fritext. Det är lätt för en
människa att läsa, men svårt för ett program. Vill Nordverk visa protokollen i
en tabell, spara dem i en databas eller räkna felkoder per månad behövs
**strukturerad data**: samma fält, i samma form, varje gång.

I det här kapitlet låter du modellen läsa ett protokoll och svara med **JSON**
som följer ett **schema**. Sedan visar du resultatet i en tabell.

## Varför inte bara be om JSON?

Du kan skriva *"Svara med JSON"* i prompten. Ofta fungerar det, men inte
alltid. Modellen kan skriva en mening före JSON:en, glömma ett fält eller
döpa ett fält till något annat. Då kraschar koden som ska läsa svaret.

Därför kan du skicka med ett **JSON-schema** i anropet. Ett schema beskriver
exakt vilka fält svaret ska ha och vilken typ de har. Med `strict: true`
tvingas modellen att följa schemat.

## JSON-schemat

Så här ser schemat för ett serviceprotokoll ut:

```js
const schema = {
  type: 'object',
  properties: {
    datum: { type: 'string' },
    kund: { type: 'string' },
    pump: { type: 'string' },
    felkod: { type: ['string', 'null'] },
    atgard: { type: 'string' },
    reservdelar: { type: 'array', items: { type: 'string' } },
  },
  required: ['datum', 'kund', 'pump', 'felkod', 'atgard', 'reservdelar'],
  additionalProperties: false,
}
```

| Del | Betyder |
|-----|---------|
| `type: 'object'` | Svaret är ett objekt `{ … }` |
| `properties` | Fälten objektet ska ha |
| `type: 'string'` | Fältet är en text |
| `type: ['string', 'null']` | En text, eller `null` om uppgiften saknas |
| `type: 'array', items: { type: 'string' }` | En lista med texter |
| `required` | Fälten som **måste** finnas. Med `strict` ska alla fält stå här. |
| `additionalProperties: false` | Inga andra fält än de vi har bestämt |

💡 Undvik å, ä och ö i fältnamnen. Därför heter fältet `atgard`, inte `åtgärd`.

## Steg 1 – Skapa komponenten Serviceprotokoll.vue

Chatten är redan stor, så protokolltolken får en egen fil, en **komponent**.
Skapa filen `src/Serviceprotokoll.vue` med den här koden:

```vue
<script setup>
import { ref } from 'vue'
import { client, MODEL } from './ai.js'

// Ett påhittat serviceprotokoll, skrivet i fritext av en tekniker.
const protokoll = ref(`Servicebesök 14/10 hos Kvarnbo Kartong AB.
Pump NV-400. Kunden har fått felkod E07 två gånger i veckan.
Bytte lager på drivsidan (6308-2RS) och båda axeltätningarna (AT-40).
Provkörde 30 min utan anmärkning. /Sara Lind`)

// Schemat: vilka fält vi vill ha och vilken typ de har.
const schema = {
  type: 'object',
  properties: {
    datum: { type: 'string' },
    kund: { type: 'string' },
    pump: { type: 'string' },
    felkod: { type: ['string', 'null'] },
    atgard: { type: 'string' },
    reservdelar: { type: 'array', items: { type: 'string' } },
  },
  required: ['datum', 'kund', 'pump', 'felkod', 'atgard', 'reservdelar'],
  additionalProperties: false,
}

// Förklara fälten i prompten också. Då blir resultatet bättre.
const instruktion = `Plocka ut uppgifterna ur serviceprotokollet.
- datum: skriv som ÅÅÅÅ-MM-DD. Året är 2026 om inget annat står.
- felkod: null om ingen felkod nämns.
- atgard: vad teknikern gjorde, en kort mening.
- reservdelar: en rad per reservdel, med antal och artikelnummer.
Hitta inte på något som inte står i texten.`

const resultat = ref(null)
const laddar = ref(false)
const fel = ref('')

async function tolka() {
  laddar.value = true
  resultat.value = null
  fel.value = ''

  try {
    const svar = await client.chat.completions.create({
      model: MODEL,
      messages: [
        { role: 'system', content: instruktion },
        { role: 'user', content: protokoll.value },
      ],
      response_format: {
        type: 'json_schema',
        json_schema: { name: 'serviceprotokoll', strict: true, schema },
      },
    })

    // Svaret är en text med JSON. Gör om den till ett objekt.
    const data = JSON.parse(svar.choices[0].message.content)

    // Kontrollera att datumet ser rätt ut, till exempel 2026-10-14
    if (!/^\d{4}-\d{2}-\d{2}$/.test(data.datum)) {
      fel.value = `Kontrollera datumet: ${data.datum}`
    }
    resultat.value = data
  } catch (e) {
    fel.value = `Kunde inte tolka protokollet: ${e.message}`
  }
  laddar.value = false
}
</script>

<template>
  <h2>Tolka serviceprotokoll</h2>
  <textarea v-model="protokoll" rows="6"></textarea>
  <button :disabled="laddar" @click="tolka">Tolka</button>
  <p v-if="laddar">Läser protokollet …</p>
  <p v-if="fel" class="fel">{{ fel }}</p>

  <table v-if="resultat">
    <tr><th>Datum</th><td>{{ resultat.datum }}</td></tr>
    <tr><th>Kund</th><td>{{ resultat.kund }}</td></tr>
    <tr><th>Pump</th><td>{{ resultat.pump }}</td></tr>
    <tr><th>Felkod</th><td>{{ resultat.felkod ?? '–' }}</td></tr>
    <tr><th>Åtgärd</th><td>{{ resultat.atgard }}</td></tr>
    <tr><th>Reservdelar</th><td>{{ resultat.reservdelar.join(', ') }}</td></tr>
  </table>
</template>

<style>
table {
  border-collapse: collapse;
  margin: 1rem 0;
}
th,
td {
  border: 1px solid #ccc;
  padding: 0.3rem 0.6rem;
  text-align: left;
}
.fel {
  color: #b00020;
}
</style>
```

## Steg 2 – Visa komponenten i App.vue

Öppna `src/App.vue`. Lägg till en import i `<script setup>`, under den andra
importen:

```js
import Serviceprotokoll from './Serviceprotokoll.vue'
```

och lägg till komponenten sist i `<template>`, före `</main>`:

```html
<Serviceprotokoll />
```

Spara och klicka **Tolka**. Efter en stund ser du en tabell ungefär så här:

| | |
|---|---|
| Datum | 2026-10-14 |
| Kund | Kvarnbo Kartong AB |
| Pump | NV-400 |
| Felkod | E07 |
| Åtgärd | Bytte lager på drivsidan och båda axeltätningarna. |
| Reservdelar | 1 x 6308-2RS, 2 x AT-40 |

Hela filerna finns i
[`facit/Serviceprotokoll.vue`](nordverk-chatt/facit/Serviceprotokoll.vue) och
[`facit/App-kapitel-5.vue`](nordverk-chatt/facit/App-kapitel-5.vue).

## Så fungerar det

**`response_format`** talar om för API:t att svaret ska följa schemat:

```js
response_format: {
  type: 'json_schema',
  json_schema: { name: 'serviceprotokoll', strict: true, schema },
},
```

`name` är ett namn du väljer själv. `strict: true` betyder att modellen
**måste** följa schemat.

**`JSON.parse`** gör om svaret, som är en text, till ett JavaScript-objekt.
Sedan kan du skriva `data.datum`, `data.pump` och så vidare.

**`instruktion`** är systemprompten. Schemat bestämmer formen, men modellen
behöver också veta vad fälten betyder. Därför förklaras de i prompten.

## Kontrollera och hantera fel

Schemat garanterar **formen**, men inte att **innehållet** är rätt. Prova att
skriva *"hej hur mår du"* i rutan och klicka **Tolka**. När vi testade fick vi
en tabell där kunden var *"Hej! Hur mår du? Jag behöver serviceprotokollet …"*
och ett påhittat datum. Alla fält fanns, precis som schemat kräver, men
innehållet var nonsens.

Därför gör koden två saker:

**1. `try` och `catch`.** Om något går fel, till exempel att servern inte
svarar eller att svaret inte är giltig JSON, hoppar koden till `catch`. Då
visas ett felmeddelande i stället för att appen slutar fungera.

**2. En kontroll av innehållet.** Koden kontrollerar att datumet ser ut som
`2026-10-14`:

```js
if (!/^\d{4}-\d{2}-\d{2}$/.test(data.datum)) {
  fel.value = `Kontrollera datumet: ${data.datum}`
}
```

`/^\d{4}-\d{2}-\d{2}$/` är ett **reguljärt uttryck**, ett mönster för text.
Det betyder: fyra siffror, bindestreck, två siffror, bindestreck, två siffror.

Kontroller i koden fångar en del fel, men inte alla. Hos Nordverk skulle en
tekniker titta igenom tabellen innan den sparas. Språkmodellen föreslår,
människan bestämmer.

## Sammanfattning

- **Strukturerad utdata** betyder att modellen svarar med samma fält i samma
  form varje gång.
- Ett **JSON-schema** beskriver fälten och deras typer.
- `response_format` med `json_schema` och `strict: true` tvingar modellen att
  följa schemat.
- `JSON.parse` gör om svaret till ett objekt.
- Schemat garanterar formen, inte innehållet. **Kontrollera** svaret och
  hantera fel med `try` och `catch`.

## Övningar

Skriv koden i `src/Serviceprotokoll.vue`. Lösningsförslag finns i
[losningar.md](losningar.md#kapitel-5--strukturerad-utdata).

**5.1** Tolka det här slarviga protokollet. Vad blev bra, och vad blev sämre?

```
Service NV900 hos Brobyverken 3 okt. ingen felkod, årlig service.
bytte filter x2 och olja 5 liter. /Ali
```

**5.2** Lägg till fältet `tekniker` (vem som gjorde jobbet). Du behöver ändra
på tre ställen: `properties`, `required` och tabellen i `<template>`.

**5.3** Visa en varning om felkoden inte ser ut som ett `E` följt av två
siffror, till exempel `E07`. Tips: mönstret är `/^E\d{2}$/`. Tänk på att
felkoden kan vara `null`.

**5.4** Lägg till ett fält `arProtokoll` med typen `boolean` (sant eller
falskt), som säger om texten verkligen är ett serviceprotokoll. Visa inte
tabellen om det är `false`, utan ett meddelande. Testa med *"hej hur mår du"*.

Nästa modul: [Bygga med AI](../02-bygga-med-ai/README.md)
