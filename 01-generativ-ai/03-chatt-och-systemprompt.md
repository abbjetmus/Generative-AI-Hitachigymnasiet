# 3. Chatt och systemprompt

Hittills svarar assistenten bara på en fråga i taget. I det här kapitlet gör
du en riktig chatt. Den minns vad ni pratat om, den vet att den jobbar på
Nordverk och svaret syns medan det skrivs.

## Steg 1 – Byt ut koden i App.vue

Byt ut hela `<script setup>` och hela `<template>` i `src/App.vue` mot koden
nedan. Låt `<style>` vara kvar.

```vue
<script setup>
import { ref } from 'vue'
import { client, MODEL } from './ai.js'

// Systemprompten bestämmer vem assistenten är och vad den vet.
// Nordverk AB är ett påhittat företag.
const systemprompt = `Du är kundtjänstassistent på Nordverk AB, ett företag i Västerås som tillverkar och servar industripumpar.

Fakta om Nordverk:
- Pumpmodeller: NV-200, NV-400 och NV-900.
- Kundtjänst har öppet vardagar kl. 8–16.
- E-post till kundtjänst: service@nordverk.example

Regler:
- Svara kort och vänligt på svenska.
- Hitta aldrig på priser eller tekniska data. Om du inte vet, säg det och hänvisa till kundtjänstens e-post.`

// Hela konversationen: varannan user, varannan assistant
const meddelanden = ref([])
const fraga = ref('')
const laddar = ref(false)
const temperatur = ref(1)
const tokens = ref(0) // hur många tokens som skickades in senast

async function skicka() {
  meddelanden.value.push({ role: 'user', content: fraga.value })
  fraga.value = ''
  laddar.value = true

  const svar = await client.chat.completions.create({
    model: MODEL,
    // Systemprompten först, sedan HELA historiken. Varje gång.
    messages: [{ role: 'system', content: systemprompt }, ...meddelanden.value],
    temperature: temperatur.value,
  })

  meddelanden.value.push({ role: 'assistant', content: svar.choices[0].message.content })
  tokens.value = svar.usage.prompt_tokens
  laddar.value = false
}

function nyKonversation() {
  meddelanden.value = []
}
</script>

<template>
  <main>
    <h1>Nordverk-assistenten</h1>

    <div class="chatt">
      <div v-for="(m, i) in meddelanden" :key="i" :class="['bubbla', m.role]">{{ m.content }}</div>
      <p v-if="laddar">Tänker …</p>
    </div>

    <form @submit.prevent="skicka">
      <input v-model="fraga" placeholder="Skriv en fråga" />
      <button :disabled="laddar || !fraga">Skicka</button>
    </form>

    <p>
      Temperatur: {{ temperatur }}
      <input v-model.number="temperatur" type="range" min="0" max="2" step="0.1" />
      <button :disabled="laddar" @click="nyKonversation">Ny konversation</button>
    </p>
    <p>Tokens in senast: {{ tokens }}</p>
  </main>
</template>
```

Lägg till det här **sist** i `<style>`, så att det ser ut som en chatt:

```css
.chatt {
  display: flex;
  flex-direction: column;
  gap: 0.5rem;
  margin-bottom: 1rem;
}
.bubbla {
  padding: 0.5rem 0.75rem;
  border-radius: 0.75rem;
  max-width: 80%;
  white-space: pre-wrap;
}
.user {
  align-self: flex-end;
  background: #d8e8ff;
}
.assistant {
  align-self: flex-start;
  background: #eeeeee;
}
```

Prova: skriv *"Hej, jag heter Omar och har en NV-200."* och sedan *"Vad heter
jag och vilken pump har jag?"*. Assistenten minns!

## Så fungerar minnet

Modellen **minns ingenting** mellan två anrop. Chatten verkar minnas för att vi
skickar med **hela historiken** varje gång:

```js
messages: [{ role: 'system', content: systemprompt }, ...meddelanden.value],
```

`...meddelanden.value` (tre punkter) lägger in alla meddelanden från listan.
Varje svar sparas i listan med rollen `assistant`. Efter tre frågor skickas
alltså:

```
1:a anropet:  system, user
2:a anropet:  system, user, assistant, user
3:e anropet:  system, user, assistant, user, assistant, user
```

💡 Titta på **Tokens in senast** när du chattar. Siffran växer för varje
fråga, eftersom allt skickas igen.

## Systemprompten

Systemprompten är det första meddelandet. Användaren ser den inte, men den
styr hur assistenten svarar. Vår systemprompt har tre delar:

| Del | Exempel |
|-----|---------|
| **Roll** | Kundtjänstassistent på Nordverk AB |
| **Fakta** | Pumpmodeller, öppettider, e-post |
| **Regler** | Svara kort. Hitta inte på priser. |

Kommer du ihåg att modellen hittade på vad Nordverk gör i kapitel 2? Nu vet den
det, eftersom det står i systemprompten. Fråga *"Vad kostar en NV-900?"*. Den
ska säga att den inte vet, i stället för att hitta på ett pris.

## Temperatur

**Temperaturen** styr hur mycket slump modellen använder när den väljer nästa
token.

| Temperatur | Effekt | Passar till |
|------------|--------|-------------|
| 0 | Nästan ingen slump. Svaren blir lika varandra. | Fakta, utdrag ur text |
| 1 | Lagom variation (standard) | Chatt |
| 2 | Mycket slump. Kreativt, men kan bli konstigt. | Idéer, namnförslag |

⚠️ **Samma fråga ger samma svar?** Aixias server sparar svar på frågor den nyss
fått (en **cache**). Skickar du exakt samma konversation igen får du det sparade
svaret, oavsett temperatur. Skriv frågan lite olika, eller lägg till den här
raden i anropet:

```js
cache: { 'no-cache': true }, // bara hos Aixia: hämta inte sparade svar
```

([Mer om cachen](https://docs.litellm.ai/docs/proxy/caching_controls).)

## Kontextfönstret

Allt som skickas, plus svaret, måste rymmas i modellens **kontextfönster**.
För gpt-oss-120b är det 131 072 tokens, ungefär en tjock bok. Det låter
mycket, men historiken växer för varje fråga. Ju längre historik, desto
långsammare och dyrare blir varje anrop. När fönstret är fullt går det inte
att skicka mer.

Det enklaste sättet att hålla nere storleken är att bara skicka de senaste
meddelandena (övning 3.4), eller att börja om med **Ny konversation**.

## Steg 2 – Streaming

Nu väntar appen på **hela** svaret innan något syns. Med **streaming** skickar
servern svaret i små bitar, så att du kan visa texten medan den skrivs.

**1.** Byt ut raden `const tokens = ref(0) …` mot:

```js
const pagaende = ref('') // svaret som håller på att skrivas
```

**2.** Byt ut allt i `skicka` från `const svar = …` till slutet av funktionen
mot:

```js
  const stream = await client.chat.completions.create({
    model: MODEL,
    messages: [{ role: 'system', content: systemprompt }, ...meddelanden.value],
    temperature: temperatur.value,
    stream: true,
  })

  // Svaret kommer i små bitar. Lägg till varje bit direkt.
  for await (const bit of stream) {
    pagaende.value += bit.choices[0]?.delta?.content ?? ''
  }

  meddelanden.value.push({ role: 'assistant', content: pagaende.value })
  pagaende.value = ''
  laddar.value = false
}
```

**3.** I `<template>`, byt raden med *Tänker …* mot:

```html
<div v-if="pagaende" class="bubbla assistant">{{ pagaende }}</div>
<p v-else-if="laddar">Tänker …</p>
```

och ta bort raden `<p>Tokens in senast: {{ tokens }}</p>`.

Så fungerar det:

- `stream: true` gör att du får svaret bit för bit.
- `for await` kör loopen en gång för varje bit som kommer.
- Texten finns i `bit.choices[0].delta.content`. Vissa bitar har ingen text,
  därför står det `?.` och `?? ''` (använd tom text om det saknas).
- När loopen är klar sparas hela svaret i historiken.

💡 Modellen resonerar först. Därför står det *Tänker …* en stund innan texten
börjar komma.

Hela filen finns i [`facit/App-kapitel-3.vue`](nordverk-chatt/facit/App-kapitel-3.vue).

## Sammanfattning

- Modellen minns inget. Chatten minns för att **hela historiken** skickas med
  varje gång.
- **Systemprompten** ger assistenten roll, fakta och regler.
- **Temperaturen** styr hur mycket slump som används.
- Allt måste rymmas i **kontextfönstret**.
- Med **streaming** syns svaret medan det skrivs.

## Övningar

Lösningsförslag finns i
[losningar.md](losningar.md#kapitel-3--chatt-och-systemprompt).

**3.1** Lägg till ett faktum i systemprompten: Nordverk har två serviceavtal,
*Bas* (ett servicebesök per år) och *Plus* (två besök per år och förtur vid
akuta fel). Fråga assistenten om serviceavtal.

**3.2** Ta bort `...meddelanden.value` ur anropet, så att bara systemprompten
skickas. Vad händer? Hur kan du skicka med bara den senaste frågan? Ändra
tillbaka efteråt.

**3.3** Lägg till `cache: { 'no-cache': true }` i anropet. Fråga *"Hitta på
ett namn på en ny pump. Svara bara med namnet."* tre gånger med temperatur 0
och tre gånger med temperatur 2. Vad är skillnaden?

**3.4** Skicka bara de **sex senaste** meddelandena. Tips:
`meddelanden.value.slice(-6)` ger de sex sista. Testa genom att säga ditt namn
först och fråga efter det efter fyra frågor till.

**3.5** Visa hur många meddelanden konversationen har, under chatten. Tips:
skriv `meddelanden.length` i `<template>`.

Nästa kapitel: [Promptteknik](04-promptteknik.md)
