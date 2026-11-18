# Lösningar

Försök själv innan du tittar här! Språkmodellen svarar inte exakt likadant
varje gång, så dina svar kan se annorlunda ut än exemplen.

I kodlösningarna visas bara det du lägger till eller ändrar.

## Kapitel 1 – Vad är generativ AI?

**1.1** *Exempel:* AI är alla program som gör något "smart", till exempel en
robotdammsugare som hittar runt möblerna. Maskininlärning är AI som lär sig av
exempel, till exempel Spotify som föreslår låtar utifrån vad du lyssnat på.
Generativ AI skapar något nytt, till exempel en app som gör en bild av en
textbeskrivning.

**1.2** *Exempel:* " måndag", " fredag", " tisdag". En veckodag är mest
sannolik. Modellen väljer ibland en mindre sannolik token eftersom det finns
lite slump med (temperaturen). Annars skulle alla svar bli exakt likadana.

**1.3** Modellen känner inte till Nordverk. Den skriver det som *låter* som ett
telefonnummer, alltså hallucinerar den. Den borde ha sagt att den inte vet och
hänvisat till Nordverks webbplats.

**1.4**

1. **Passar.** Att skriva text är en styrka, och en människa kontrollerar.
2. **Passar inte.** Exakt räkning görs bättre med en databasfråga eller ett
   kalkylark.
3. **Passar.** Att läsa och sammanfatta är en styrka. Men viktiga detaljer bör
   kontrolleras.
4. **Passar inte.** Beslutet påverkar en människa, och modellen kan vara
   snedvriden. En människa ska bestämma.

## Kapitel 2 – Kom igång

**2.1** Det står `Bearer laggs-till-av-servern`. Det är låtsasnyckeln från
`src/ai.js`. Den riktiga nyckeln läggs på av Vite efter att anropet har
lämnat webbläsaren, så den syns aldrig här.

**2.2** Ändra systemmeddelandet i `skicka`:

```js
{ role: 'system', content: 'Du är en gammal sjökapten. Svara kort på svenska, som en sjökapten skulle prata.' },
```

**2.3** Modellen svarar något som låter rimligt, till exempel att Nordverk
säljer byggmaterial eller ventilation. Det är fel. Modellen **hallucinerar**:
den känner inte till det påhittade företaget. För att svara rätt måste den få
fakta om Nordverk i prompten. Det gör du i kapitel 3.

**2.4** Byt `<pre>`-raden mot:

```html
<p v-if="helaSvaret">
  Tokens in: {{ helaSvaret.usage.prompt_tokens }},
  tokens ut: {{ helaSvaret.usage.completion_tokens }}
</p>
```

**2.5** *Exempel:* "Pumpen läcker vid axeln och vi behöver en tekniker så snart
som möjligt." och samma mening på engelska. Hos oss blev det 18 och 17 tokens
för själva meningen. Med kursens modell är skillnaden liten för vanliga ord.
Namn och långa sammansatta ord delas upp i fler bitar.

## Kapitel 3 – Chatt och systemprompt

**3.1** Lägg till under *Fakta om Nordverk*:

```
- Serviceavtal: Bas (ett servicebesök per år) och Plus (två besök per år och förtur vid akuta fel).
```

Nu kan assistenten svara på frågor om serviceavtalen.

**3.2** Utan historiken får modellen bara systemprompten och vet inte vad du
frågade. Den svarar ungefär "Hur kan jag hjälpa dig?". Den senaste frågan är
det sista meddelandet i listan:

```js
messages: [{ role: 'system', content: systemprompt }, meddelanden.value.at(-1)],
```

Då svarar den på frågan, men minns inget. Fråga *"Vad heter jag?"* efter att
du sagt ditt namn, så märks det.

**3.3** Lägg till i anropet:

```js
temperature: temperatur.value,
cache: { 'no-cache': true },
```

*Svar:* Med temperatur 0 blir namnen samma eller nästan samma varje gång. Med
temperatur 2 blir de olika varje gång, och ibland konstiga.

**3.4** Byt `...meddelanden.value` mot:

```js
messages: [{ role: 'system', content: systemprompt }, ...meddelanden.value.slice(-6)],
```

Efter fyra frågor till har ditt första meddelande försvunnit ur de sex
senaste. Assistenten vet inte längre vad du heter.

**3.5** Lägg till under chatten i `<template>`:

```html
<p>Antal meddelanden: {{ meddelanden.length }}</p>
```

## Kapitel 4 – Promptteknik

**4.1** Lägg till under *Regler* i systemprompten:

```
- Skriv vanlig text utan Markdown, utan stjärnor och utan formler.
```

**4.2** Båda prompterna sorterar oftast rätt. Gränsfallet är fakturan: den
blir ibland *reklamation* och ibland *övrigt*. Med few-shot bestämmer **du**
hur gränsfallen ska sorteras, genom att lägga till ett exempel:

```
Meddelande: Fakturan har fel belopp.
Kategori: övrigt
```

**4.3** 8 timmar × 5 dagar = 40 timmar i veckan. 2000 / 40 = **50 veckor**.
Modellen räknar oftast rätt båda gångerna. Svaret med steg är lättare att
kontrollera, eftersom du ser varje del av uträkningen.

**4.4** *Exempel:*

```
Du får ett kundmejl mellan <mejl> och </mejl>.
Gör om det till en punktlista med:
- Kund: kundens namn
- Pump: vilken pump det gäller
- Vill: vad kunden vill
Skriv bara punktlistan. Skriv "okänt" om något saknas.
```

*Svar:* Kund: Kim, Pump: NV-200, Vill: att en tekniker kommer i veckan.

**4.5** Lägg till i `<script setup>`:

```js
function granska() {
  fraga.value = 'Granska ditt förra svar. Finns det något du har hittat på eller inte kan veta säkert? Rätta i så fall svaret.'
  skicka()
}
```

och i `<template>`, bredvid **Ny konversation**:

```html
<button :disabled="laddar || meddelanden.length === 0" @click="granska">Granska svaret</button>
```

**4.6** Hos oss lyckades attacken ungefär en gång av tre utan de extra raderna.
Med raderna lyckades den inte på tre försök. Men det betyder inte att den
aldrig kan lyckas. En systemprompt är inget lås.

## Kapitel 5 – Strukturerad utdata

**5.1** *Exempel på resultat:* datum 2026-10-03, kund Brobyverken, pump NV900,
felkod –, åtgärd "Bytte två filter och fyllde på 5 liter olja.", reservdelar
"2 filter". Modellen förstod det korta datumet och att felkod saknas. Men
pumpen fick fel stavning (NV900 i stället för NV-900) och oljan kom inte med
bland reservdelarna. Det är sådant en människa behöver kontrollera.

**5.2** I `properties`:

```js
tekniker: { type: 'string' },
```

I `required`, lägg till `'tekniker'` sist i listan. I tabellen:

```html
<tr><th>Tekniker</th><td>{{ resultat.tekniker }}</td></tr>
```

**5.3** Lägg till efter kontrollen av datumet:

```js
if (data.felkod !== null && !/^E\d{2}$/.test(data.felkod)) {
  fel.value = `Kontrollera felkoden: ${data.felkod}`
}
```

**5.4** I `properties` och `required`:

```js
arProtokoll: { type: 'boolean' },
```

```js
required: ['arProtokoll', 'datum', 'kund', 'pump', 'felkod', 'atgard', 'reservdelar'],
```

Lägg till en rad i `instruktion`:

```
- arProtokoll: true om texten är ett serviceprotokoll, annars false.
```

Direkt efter `JSON.parse`:

```js
if (!data.arProtokoll) {
  fel.value = 'Texten ser inte ut som ett serviceprotokoll.'
  laddar.value = false
  return
}
```

`return` avslutar funktionen, så tabellen visas aldrig.
