# Modul 01 – Generativ AI

I den här modulen lär du dig vad generativ AI är och hur en språkmodell
fungerar. Du bygger en chatt åt det påhittade företaget **Nordverk AB** och
låter modellen läsa serviceprotokoll och svara med data som ett program kan
använda.

## Det här kan du efter modulen

- Förklara vad generativ AI och en språkmodell är, och vad tokens är
- Nämna språkmodellers begränsningar: hallucinationer, kunskapsgräns, bias och kontextfönster
- Anropa en språkmodell från en Vue-app utan att avslöja API-nyckeln
- Bygga en chatt med historik, systemprompt, temperatur och streaming
- Skriva bättre promptar med tydliga instruktioner, exempel och avgränsningar
- Förklara vad prompt injection är
- Få modellen att svara med JSON enligt ett schema och kontrollera svaret

## Upplägg

### Del 1 – Grunderna

| Kapitel | Vad du gör |
|---------|------------|
| [01 – Vad är generativ AI?](01-vad-ar-generativ-ai.md) | Lär dig hur en språkmodell fungerar och vad den inte klarar |
| [02 – Kom igång](02-kom-igang.md) | Lägger in nyckeln, startar [`nordverk-chatt`](nordverk-chatt/) och gör ditt första anrop |

### Del 2 – Bygg Nordverk-assistenten

| Kapitel | Vad du gör |
|---------|------------|
| [03 – Chatt och systemprompt](03-chatt-och-systemprompt.md) | Bygger en chatt som minns, jobbar på Nordverk och strömmar svaret |
| [04 – Promptteknik](04-promptteknik.md) | Provar knep för bättre promptar och ser hur en prompt kan luras |
| [05 – Strukturerad utdata](05-strukturerad-utdata.md) | Plockar ut fält ur ett serviceprotokoll och visar dem i en tabell |

All kod skriver du i Vue-projektet [`nordverk-chatt`](nordverk-chatt/). Om du
kör fast finns färdiga filer för varje kapitel i
[`nordverk-chatt/facit`](nordverk-chatt/facit/).

## Det här behöver du

- En API-nyckel som du får.
