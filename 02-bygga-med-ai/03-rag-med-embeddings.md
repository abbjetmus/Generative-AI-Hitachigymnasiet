# 3. RAG med embeddings

Nyckelordssökningen hittade ingenting för *"Pumpen skakar och låter konstigt"*,
eftersom manualen skriver *vibration*. Vi behöver en sökning som förstår
**betydelse**, inte bara exakta ord.

## Vad är en embedding?

En **embedding** är en text omgjord till en lista med tal, en **vektor**.
Modellen `bge-m3` gör om varje text till 1024 tal. Texter med liknande
betydelse får vektorer som pekar åt nästan samma håll, även om orden är olika.

```
(talen är påhittade exempel)
"Pumpen skakar"           → [0.012, -0.044, 0.031, ...]   (1024 tal)
"Vibration över 4,5 mm/s" → [0.015, -0.040, 0.027, ...]   liknar!
"Vem vann fotbolls-VM?"   → [-0.031, 0.008, -0.052, ...]  liknar inte
```

`bge-m3` klarar över 100 språk, svenska också, och texter på upp till 8 192
tokens ([modellkortet på Hugging Face](https://huggingface.co/BAAI/bge-m3)).

## Cosinuslikhet

Hur nära är två vektorer? Tänk dig två pilar. Pekar de åt samma håll är de lika.
Det mäts med **cosinuslikhet**: ett tal där 1 betyder samma riktning och 0
betyder att de inte har något med varandra att göra. Du behöver inte kunna
formeln, men den står i koden nedan.

## Chunkning

Innan texten görs om till vektorer delas den i bitar, **chunkar**. Om en bit är
för lång blir vektorn ett genomsnitt av många ämnen och liknar inget särskilt.
Om den är för kort saknas sammanhanget. Vi använder avsnitten från kapitel 2
som chunkar, och lägger källan (`NV-400, Felkoder`) först i varje bit så att
vektorn vet vilken pump texten gäller.

## Vektorsökning i minnet

1. Gör om alla avsnitt till vektorer en gång. Spara dem i en variabel.
2. Gör om frågan till en vektor.
3. Räkna cosinuslikheten mot alla avsnitt och ta de bästa.

Skapa filen `src/vektorsok.js`:

```js
// Vektorsökning med embeddings. Allt ligger i minnet i webbläsaren.
import { client, EMBED_MODEL } from './ai.js'
import { avsnitt } from './sok.js'

// Gör om texter till vektorer (listor med 1024 tal).
async function skapaEmbeddings(texter) {
  const svar = await client.embeddings.create({ model: EMBED_MODEL, input: texter })
  return svar.data.map((d) => d.embedding)
}

// Cosinuslikhet: nära 1 = liknande betydelse, nära 0 = inget gemensamt.
function cosinus(a, b) {
  let produkt = 0
  let langdA = 0
  let langdB = 0
  for (let i = 0; i < a.length; i++) {
    produkt += a[i] * b[i]
    langdA += a[i] * a[i]
    langdB += b[i] * b[i]
  }
  return produkt / (Math.sqrt(langdA) * Math.sqrt(langdB))
}

// Vektorerna för alla avsnitt. Skapas första gången någon söker.
let vektorer = null

export async function vektorsok(fraga, antal = 3) {
  if (!vektorer) {
    vektorer = await skapaEmbeddings(avsnitt.map((a) => a.kalla + '\n' + a.text))
  }
  const [fragaVektor] = await skapaEmbeddings([fraga])

  return avsnitt
    .map((a, i) => ({ ...a, poang: cosinus(fragaVektor, vektorer[i]) }))
    .sort((a, b) => b.poang - a.poang)
    .slice(0, antal)
}
```

💡 Här ligger vektorerna i minnet och skapas på nytt varje gång sidan laddas.
Med tusentals dokument sparar man dem i en **vektordatabas**.

## Byt sökning i FragaManualen.vue

Vi lägger till en lista där du väljer sökmetod. Vektorsökningen hittar alltid
något, även för helt orelaterade frågor. Därför har vi en **gräns**: träffar med
lägre likhet än 0,45 räknas inte. Värdet tog vi fram genom att testa.

Ersätt `src/FragaManualen.vue` med:

```vue
<script setup>
import { ref } from 'vue'
import { chatta } from './ai.js'
import { nyckelordssok } from './sok.js'
import { vektorsok } from './vektorsok.js'

const fraga = ref('')
const svar = ref('')
const traffar = ref([])
const laddar = ref(false)
const sokmetod = ref('embeddings')

// Bitar med lägre likhet än så här räknas inte som träffar.
const GRANS = 0.45

async function fragaManualen() {
  laddar.value = true
  svar.value = ''

  // 1. Sök (retrieval)
  if (sokmetod.value === 'embeddings') {
    const resultat = await vektorsok(fraga.value)
    traffar.value = resultat.filter((t) => t.poang >= GRANS)
  } else {
    traffar.value = nyckelordssok(fraga.value)
  }

  // Ingen träff: fråga inte ens modellen.
  if (traffar.value.length === 0) {
    svar.value = 'Det står inte i manualerna.'
    laddar.value = false
    return
  }

  // 2. Lägg in träffarna i prompten (augmentation)
  const underlag = traffar.value
    .map((t, i) => `[${i + 1}] Källa: ${t.kalla}\n${t.text}`)
    .join('\n\n')

  // 3. Låt modellen skriva svaret (generation)
  svar.value = await chatta([
    {
      role: 'system',
      content:
        'Du är Nordverks serviceassistent. Svara kort på svenska och bara utifrån underlaget. ' +
        'Skriv källans nummer inom hakparentes efter det du hämtat därifrån, till exempel [1]. ' +
        'Om svaret inte finns i underlaget, svara exakt: "Det står inte i manualerna."',
    },
    { role: 'user', content: `UNDERLAG:\n${underlag}\n\nFRÅGA: ${fraga.value}` },
  ])
  laddar.value = false
}
</script>

<template>
  <section>
    <h2>Fråga manualen</h2>
    <form @submit.prevent="fragaManualen">
      <select v-model="sokmetod">
        <option value="nyckelord">Nyckelord</option>
        <option value="embeddings">Embeddings</option>
      </select>
      <input v-model="fraga" placeholder="Till exempel: Vad betyder E06 på NV-400?" />
      <button :disabled="laddar || !fraga">Sök och svara</button>
    </form>
    <p v-if="laddar">Söker och tänker …</p>
    <p v-if="svar">{{ svar }}</p>
    <ol v-if="traffar.length">
      <li v-for="(t, i) in traffar" :key="i">
        {{ t.kalla }} (poäng {{ sokmetod === 'embeddings' ? t.poang.toFixed(2) : t.poang }})
      </li>
    </ol>
  </section>
</template>
```

## Jämför

| Fråga | Nyckelord | Embeddings (poäng) |
|---|---|---|
| Pumpen skakar och låter konstigt | Ingen träff | NV-400, Vibration (0,62) |
| Vad betyder E06 på NV-400? | NV-400, Felkoder | NV-400, Felkoder (0,57) |
| Vem vann fotbolls-VM 2022? | Ingen träff | Bästa träff 0,24, under gränsen |

Embeddings hittar rätt avsnitt för *skakar*. Men lägg märke till att svaret
handlade om NV-400, fast frågan inte sa vilken pump det gällde. En bra assistent
borde fråga. Nyckelord är fortfarande bra när frågan innehåller exakta ord, som
felkoder och artikelnummer. Många system använder båda.

## Sammanfattning

| Begrepp | Betyder |
|---|---|
| Embedding | En text omgjord till en vektor, en lista med tal |
| Cosinuslikhet | Hur lika två vektorer är, 1 = samma riktning |
| Chunk | En bit av ett dokument som får en egen vektor |
| Vektorsökning | Hitta de chunkar vars vektorer liknar frågans |
| Gräns | Lägsta likhet som räknas som träff |

Nästa kapitel: [Verktygsanrop](04-verktygsanrop.md)
