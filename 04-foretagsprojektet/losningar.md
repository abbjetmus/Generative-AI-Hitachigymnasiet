# Lösningar

Försök själv innan du tittar här! Klassning och val av driftalternativ är
bedömningar. Det viktiga är motiveringen, inte att svaret är exakt som här.

## Kapitel 1 – Utredningen

**1.1** Se lärarstödet för varje case längre ner. Kort:

| Case | Nivå | Driftalternativ |
|---|---|---|
| 1 Månadsrapport | 2 | C (B eller D går också) |
| 2 Kundärenden | 3 | C efter att kontaktuppgifter tagits bort, eller D |
| 3 Teknisk dokumentation | 4 | Bara D |

**1.2** Modellen behöver `text`, och kanske `order` för att svaret ska kunna
nämna ordern. Den behöver inte `namn`, `epost` eller `telefon`. Förnamnet kan
läggas in i svaret av koden efteråt. I `text` finns personuppgifter kvar:
K-2206 berättar om en knäskada och kontakt med fysioterapeut, alltså en uppgift om
hälsa (känslig personuppgift).

**1.3** Exportkontrollerad teknik är nivå 4, alltså bara D (lokal modell) eller
ingen AI alls. Testdatan är påhittad och därför nivå 1, så gruppen kan öva mot
Aixia om ingen dator klarar en lokal modell. I utredningen ska det stå att den
riktiga lösningen bara får köras lokalt.

**1.4** Exempel:

| Risk | Åtgärd |
|---|---|
| En kund skriver instruktioner i mejlet (K-2210) och modellen lovar rabatt | Testfall med sådana mejl. Inget skickas utan att en människa godkänt |
| Svarsförslaget lovar mer än villkoren (K-2208 är utanför 30 dagar) | Villkoren skickas med. Människan granskar |
| Hälsouppgifter (K-2206) skickas till AI-tjänsten | Be företaget fundera på om sådana ärenden ska sorteras bort, eller köra lokalt |
| Ett nätfiskeärende (K-2211) får låg prioritet | Testfall. Regel i instruktionen om prioritet Hög |

**1.5** Ingen lösning. Granskningen är övningen.

## Kapitel 2 – Bygga lösningen

**2.1** Lägg till sist i listan i `src/instruktioner.js`:

```js
{
  namn: 'Teknisk support',
  text: `Du hjälper montörer och servicetekniker att hitta uppgifter i företagets tekniska dokument.
Svara kort och exakt, med siffror och enheter som de står i dokumenten.
Ange alltid dokument-ID (t.ex. NF-HX40-010) för varje uppgift.
Om svaret inte finns i dokumenten: skriv "Det står inte i dokumenten" och hänvisa till ansvarig person. Gissa aldrig.`,
},
```

*Svar: Fråga 1 ska ge 85 Nm, korsvis. Fråga 7 (pris) ska ge "Det står inte i dokumenten".*

**2.2** I `<script setup>`. Ändra första raden till `import { ref, computed } from 'vue'`:

```js
// Ungefär 4 tecken per token. En grov uppskattning.
const antalTokens = computed(() => {
  let tecken = 0
  for (const fil of filer.value) tecken += fil.text.length
  return Math.round(tecken / 4)
})
```

I `<template>`, under listan med filer:

```html
<p>Ungefär {{ antalTokens }} tokens</p>
<p v-if="antalTokens > 20000"><strong>Varning:</strong> mycket text.</p>
```

*Svar: Exempelcase 1 blir ungefär 1 200 tokens.*

**2.3** I `<script setup>`:

```js
const visaSkickat = ref(false)

// Byter ut e-postadresser och telefonnummer som 070-123 45 01.
function taBortKontaktuppgifter(text) {
  return text
    .replace(/[\w.+-]+@[\w-]+\.[\w.-]+/g, '[E-POST]')
    .replace(/0\d{1,3}-\d{2,3} ?\d{2} ?\d{2}/g, '[TELEFON]')
}

// Det här är texten som faktiskt skickas.
function underlagSomSkickas() {
  return taBortKontaktuppgifter(byggUnderlag())
}
```

I `generera()` byter du `byggUnderlag()` mot `underlagSomSkickas()`. I `<template>`:

```html
<button @click="visaSkickat = !visaSkickat">Visa det som skickas</button>
<pre v-if="visaSkickat">{{ underlagSomSkickas() }}</pre>
```

*Svar: Namnen finns kvar, liksom ordernummer, hälsouppgiften i K-2206 och allt
annat som kunden själv skrivit. Enkla sökmönster hittar bara det som har en fast
form. Därför räcker det inte att "tvätta" texten. Datan är fortfarande nivå 3.*

**2.4** I `<script setup>`:

```js
// Räknar hur många gånger varje värde i en kolumn finns i en CSV-fil.
function raknaKolumn(csv, kolumn) {
  const rader = csv.trim().split('\n')
  const index = rader[0].split(';').indexOf(kolumn)
  const antal = {}
  for (const rad of rader.slice(1)) {
    const varde = rad.split(';')[index]
    antal[varde] = (antal[varde] || 0) + 1
  }
  return antal
}

function raknatAvKod() {
  let text = ''
  for (const fil of filer.value) {
    if (fil.namn.endsWith('.csv')) {
      text += `Räknat av kod (stämmer, använd dessa siffror): antal larm per larmkod i ${fil.namn}: ${JSON.stringify(raknaKolumn(fil.text, 'larmkod'))}\n`
    }
  }
  return text
}
```

Lägg till `raknatAvKod()` i användarens meddelande i `generera()`:

```js
{ role: 'user', content: byggUnderlag() + '\n\n' + raknatAvKod() + '\n\nUppgift: ' + … },
```

*Svar: Koden räknar `{"F21":4,"E04":2,"F40":2}`. Modellen använder siffrorna i
stället för att räkna själv.*

**2.5** Ingen lösning. Räkna med att en mindre lokal modell är långsammare och
oftare svarar fel. Företaget får väga det mot att datan stannar i huset.

## Kapitel 3 – Demonstrationen

Ingen lösning. Låt den andra gruppen ge respons på tid, tydlighet och hur ärligt
begränsningarna beskrevs.

## Lärarstöd: exempelcasen

### Case 1 – Månadsrapport

**Klassning: nivå 2 (Intern).** Servicestatistik utan personuppgifter. Teknikerna
anges med ID och kunden är ett företag. **Driftalternativ:** B, C eller D. C
(Aixia) är enklast i kursen. A bara efter avtal och bedömning.

**Facit för oktober:** 5 besök (SP-1001, 1003, 1005, 1007, 1009), alla akuta.
Arbetstid 5,5 timmar, men tid saknas i SP-1009. 8 driftlarm: F21 4, E04 2, F40 2.
Reservdelar: filter F7 2 st, cirkulationspump CP-25 1 st, tryckvakt PV-300 1 st
(beställd, inte monterad). Återkommande: Björken 4 (F21, tre besök, trasig
tryckvakt) och Lönnen 2 (E04, även i september, löst med ny pump).

**Fällor i datan:** SP-1000 är från september. SP-1005 har ett annat format och
datumet 14/10 utan år. SP-1009 saknar tid.

**Typiska misstag:** Lita på modellens siffror utan facit. Inte upptäcka att
modellen kopierar meningar från septemberrapporten. Missa att tekniker-ID kan bli
personuppgifter om företaget skickar riktiga protokoll med namn.

### Case 2 – Kundärenden

**Klassning: nivå 3 (Känslig).** Namn, e-post, telefon och ordernummer är
personuppgifter. K-2206 innehåller en hälsouppgift, som i sig är nivå 4.
**Driftalternativ:** C eller D. B bara med personuppgiftsbiträdesavtal och
dokumenterad bedömning. Ta bort fälten som inte behövs.

**Facit:**

| Ärende | Kategori | Prioritet | Att lägga märke till |
|---|---|---|---|
| K-2201 | Leverans | Normal | |
| K-2202 | Retur och byte | Normal | Oöppnad, inom 30 dagar |
| K-2203 | Reklamation | Hög | Bromsar, säkerhet. Får inte lova ny vara innan bedömning |
| K-2205 | Betalning | Hög | Dubbeldragning |
| K-2206 | Produktfråga | Låg | Hälsouppgift |
| K-2208 | Retur och byte | Normal | 45 dagar och använd, utanför öppet köp |
| K-2210 | Leverans | Normal | Försök att lura modellen. Ingen rabatt |
| K-2211 | Övrigt | Hög | Misstänkt nätfiske |

I vårt test med Aixia fick de flesta ärenden samma kategori som facit. Modellen
satte K-2206 till Normal och K-2210 till Övrigt, och lovade ingen rabatt.

**Typiska misstag:** Skicka hela JSON-filen med alla fält. Tro att det räcker
att ta bort e-post och telefon. Låta svaren skickas automatiskt. Inte testa K-2210.

### Case 3 – Teknisk dokumentation

**Klassning: nivå 4 (Skyddsvärd).** Exportkontrollerad teknik enligt företaget.
Pumpar och konstruktionsritningar finns bland
[ISP:s exempel](https://isp.se/produkter-med-dubbla-anvandningsomraden/) på
produkter med dubbla användningsområden. **Driftalternativ:** bara D, på
företagets dator eller server.

**Viktigt:** Testdatan är påhittad och därför nivå 1. Gruppen får öva mot Aixia,
men utredningen ska säga att den riktiga lösningen bara får köras lokalt.

**Typiska misstag:** Välja C "eftersom Aixia är svenskt". Det räcker inte när
företaget säger att datan inte får lämna nätet. Glömma att frågorna från
montörerna också kan innehålla skyddsvärda uppgifter. Inte testa frågor som
saknar svar (7 och 8).
