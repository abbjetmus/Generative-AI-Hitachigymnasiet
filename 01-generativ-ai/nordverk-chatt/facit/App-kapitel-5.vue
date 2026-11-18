<script setup>
import { ref } from 'vue'
import { client, MODEL } from './ai.js'
import Serviceprotokoll from './Serviceprotokoll.vue'

// Systemprompten bestämmer vem assistenten är och vad den vet.
// Nordverk AB är ett påhittat företag.
const systemprompt = ref(`Du är kundtjänstassistent på Nordverk AB, ett företag i Västerås som tillverkar och servar industripumpar.

Fakta om Nordverk:
- Pumpmodeller: NV-200, NV-400 och NV-900.
- Kundtjänst har öppet vardagar kl. 8–16.
- E-post till kundtjänst: service@nordverk.example

Regler:
- Svara kort och vänligt på svenska.
- Hitta aldrig på priser eller tekniska data. Om du inte vet, säg det och hänvisa till kundtjänstens e-post.`)

// Hela konversationen: varannan user, varannan assistant
const meddelanden = ref([])
const fraga = ref('')
const laddar = ref(false)
const temperatur = ref(1)
const pagaende = ref('') // svaret som håller på att skrivas

async function skicka() {
  meddelanden.value.push({ role: 'user', content: fraga.value })
  fraga.value = ''
  laddar.value = true

  const stream = await client.chat.completions.create({
    model: MODEL,
    // Systemprompten först, sedan HELA historiken. Varje gång.
    messages: [{ role: 'system', content: systemprompt.value }, ...meddelanden.value],
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

function nyKonversation() {
  meddelanden.value = []
}
</script>

<template>
  <main>
    <h1>Nordverk-assistenten</h1>

    <div class="chatt">
      <div v-for="(m, i) in meddelanden" :key="i" :class="['bubbla', m.role]">{{ m.content }}</div>
      <div v-if="pagaende" class="bubbla assistant">{{ pagaende }}</div>
      <p v-else-if="laddar">Tänker …</p>
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

    <h2>Systemprompt</h2>
    <textarea v-model="systemprompt" rows="12"></textarea>

    <Serviceprotokoll />
  </main>
</template>

<style>
body {
  font-family: system-ui, sans-serif;
  max-width: 700px;
  margin: 2rem auto;
  padding: 0 1rem;
}
form {
  display: flex;
  gap: 0.5rem;
}
input,
textarea {
  flex: 1;
  padding: 0.5rem;
  font: inherit;
}
pre {
  background: #f4f4f4;
  padding: 1rem;
  overflow-x: auto;
  white-space: pre-wrap;
}
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
textarea {
  width: 100%;
  box-sizing: border-box;
}
</style>
