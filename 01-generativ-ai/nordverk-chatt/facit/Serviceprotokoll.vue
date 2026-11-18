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
