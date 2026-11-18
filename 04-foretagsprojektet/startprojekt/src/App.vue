<script setup>
import { ref } from 'vue'
import { chatta, MODEL } from './ai.js'
import { instruktioner, grundregler } from './instruktioner.js'

const filer = ref([]) // de inlästa filerna, { namn, text } för varje fil
const valdInstruktion = ref(instruktioner[0])
const uppgift = ref('')
const datanFarSkickas = ref(false)
const resultat = ref('')
const laddar = ref(false)

// Läser de valda filerna som text. Inget skickas förrän du klickar på Generera.
async function lasInFiler(event) {
  for (const fil of event.target.files) {
    filer.value.push({ namn: fil.name, text: await fil.text() })
  }
}

// Lägger ihop alla filer till en text, med filnamnet som rubrik före varje fil.
function byggUnderlag() {
  return filer.value.map((fil) => `### Fil: ${fil.namn}\n\n${fil.text}`).join('\n\n')
}

async function generera() {
  laddar.value = true
  try {
    resultat.value = await chatta([
      { role: 'system', content: valdInstruktion.value.text + '\n\n' + grundregler },
      { role: 'user', content: byggUnderlag() + '\n\nUppgift: ' + (uppgift.value || 'Gör uppgiften utifrån filerna.') },
    ])
  } catch (fel) {
    resultat.value = 'Något gick fel: ' + fel.message
  }
  laddar.value = false
}

// Sparar resultatet som en Markdown-fil på datorn.
function laddaNer() {
  const lank = document.createElement('a')
  lank.href = URL.createObjectURL(new Blob([resultat.value], { type: 'text/markdown' }))
  lank.download = 'resultat.md'
  lank.click()
}
</script>

<template>
  <main>
    <h1>Företagsprojektet</h1>

    <h2>1. Läs in testdata</h2>
    <input type="file" multiple accept=".txt,.md,.csv,.json" @change="lasInFiler" />
    <ul>
      <li v-for="fil in filer" :key="fil.namn">{{ fil.namn }} ({{ fil.text.length }} tecken)</li>
    </ul>
    <button v-if="filer.length" @click="filer = []">Töm listan</button>

    <h2>2. Välj instruktion</h2>
    <select v-model="valdInstruktion">
      <option v-for="i in instruktioner" :key="i.namn" :value="i">{{ i.namn }}</option>
    </select>
    <textarea v-model="valdInstruktion.text" rows="6"></textarea>
    <input v-model="uppgift" placeholder="Uppgift eller fråga, t.ex. Skriv rapporten för oktober" />

    <h2>3. Generera</h2>
    <label>
      <input type="checkbox" v-model="datanFarSkickas" />
      Jag har kontrollerat att datan får skickas till AI-tjänsten (modell: {{ MODEL }})
    </label>
    <p><button :disabled="laddar || !filer.length || !datanFarSkickas" @click="generera">Generera</button></p>

    <p v-if="laddar">Tänker … (det kan ta en halv minut)</p>
    <template v-else-if="resultat">
      <pre>{{ resultat }}</pre>
      <button @click="laddaNer">Ladda ner som .md</button>
    </template>
  </main>
</template>

<style>
body {
  font-family: system-ui, sans-serif;
  max-width: 800px;
  margin: 2rem auto;
  padding: 0 1rem;
}
input:not([type]),
textarea,
select {
  display: block;
  width: 100%;
  box-sizing: border-box;
  margin: 0.5rem 0;
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
