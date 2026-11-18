<script setup>
import { ref } from 'vue'
import { chatta } from './ai.js'
import { manualer, serviceprotokoll, kundarenden } from './nordverk.js'

const fraga = ref('')
const svar = ref('')
const laddar = ref(false)

async function skicka() {
  laddar.value = true
  svar.value = await chatta([
    { role: 'system', content: 'Du är Nordverks assistent. Svara kort på svenska.' },
    { role: 'user', content: fraga.value },
  ])
  laddar.value = false
}
</script>

<template>
  <main>
    <h1>Nordverk-assistenten</h1>
    <p>
      Inläst: {{ manualer.length }} manualer, {{ serviceprotokoll.length }} serviceprotokoll
      och {{ kundarenden.length }} kundärenden.
    </p>

    <!-- Här lägger du till komponenterna du bygger i kapitel 1–5. -->

    <h2>Fråga assistenten</h2>
    <form @submit.prevent="skicka">
      <input v-model="fraga" placeholder="Skriv en fråga" />
      <button :disabled="laddar || !fraga">Skicka</button>
    </form>
    <p v-if="laddar">Tänker …</p>
    <p v-else-if="svar">{{ svar }}</p>
  </main>
</template>

<style>
body {
  font-family: system-ui, sans-serif;
  max-width: 800px;
  margin: 2rem auto;
  padding: 0 1rem;
}
section {
  border-top: 1px solid #ccc;
  margin-top: 2rem;
}
form {
  display: flex;
  gap: 0.5rem;
}
input,
textarea,
select {
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
.logg {
  font-family: monospace;
  font-size: 0.9rem;
}
</style>
