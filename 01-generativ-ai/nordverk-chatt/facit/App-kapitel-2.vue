<script setup>
import { ref } from 'vue'
import { client, MODEL } from './ai.js'

const fraga = ref('')
const svar = ref('')
const helaSvaret = ref(null)
const laddar = ref(false)

async function skicka() {
  laddar.value = true
  const response = await client.chat.completions.create({
    model: MODEL,
    messages: [
      { role: 'system', content: 'Du är en hjälpsam assistent. Svara kort på svenska.' },
      { role: 'user', content: fraga.value },
    ],
  })
  helaSvaret.value = response
  svar.value = response.choices[0].message.content
  laddar.value = false
}
</script>

<template>
  <main>
    <h1>Nordverk-assistenten</h1>
    <form @submit.prevent="skicka">
      <input v-model="fraga" placeholder="Skriv en fråga" />
      <button :disabled="laddar || !fraga">Skicka</button>
    </form>
    <p v-if="laddar">Tänker …</p>
    <p v-else-if="svar">{{ svar }}</p>
    <pre v-if="helaSvaret">{{ helaSvaret }}</pre>
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
</style>
