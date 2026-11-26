# Modul 02 – Bygga med AI

I modul 01 lärde du dig att prata med en språkmodell. Nu bygger du
**Nordverk-assistenten**, ett verktyg för det påhittade pumpföretaget Nordverk
AB. Assistenten skriver rapporter från företagets dokument, svarar på frågor
utifrån manualerna, hämtar data med verktyg och hanterar kundärenden.

All data finns i [`data/nordverk`](../data/nordverk/). Företaget och datan är
påhittade.

## Det här kan du efter modulen

- Skapa en rapport från många dokument med map–reduce och en mall
- Förklara varför koden ska räkna och modellen skriva text
- Förklara vad RAG är och bygga en enkel RAG med nyckelordssökning
- Förklara vad embeddings och cosinuslikhet är och söka på betydelse
- Låta modellen välja funktioner som programmet kör (verktygsanrop)
- Bygga en enkel agent med maxantal steg, logg och mänskligt godkännande
- Förklara vad MCP är och bygga en MCP-server och en klient

## Upplägg

### Del 1 – Företagets dokument

| Kapitel | Vad du gör |
|---------|------------|
| [01 – Rapporter från befintliga rapporter](01-rapporter-fran-befintliga-rapporter.md) | Skapar en månadsrapport från 12 serviceprotokoll |
| [02 – RAG med sökning](02-rag-med-sokning.md) | Låter assistenten svara utifrån manualerna, med källor |
| [03 – RAG med embeddings](03-rag-med-embeddings.md) | Söker på betydelse i stället för exakta ord |

### Del 2 – AI som gör saker

| Kapitel | Vad du gör |
|---------|------------|
| [04 – Verktygsanrop](04-verktygsanrop.md) | Låter modellen be programmet köra funktioner |
| [05 – Agenter](05-agenter.md) | Bygger en agent som hanterar kundärenden |
| [06 – MCP](06-mcp.md) | Gör Nordverks verktyg tillgängliga för alla AI-program |
| [07 – Agenter i Microsoft 365](07-agenter-i-microsoft-365.md) | Ser hur färdiga agenter fungerar i Microsoft 365 |

## Projektet

Kapitel 1–5 bygger vidare på Vue-projektet
[`nordverk-assistent`](nordverk-assistent/). Varje kapitel lägger till en ny
komponent i `src/` och en rad i `src/App.vue`. Kapitel 6 använder Node-skripten i
[`kod/`](kod/), eftersom en MCP-server är ett eget program som läser filer från
disken.

## Det här behöver du

- [Node.js](https://nodejs.org/) version 20.12 eller senare och [VS Code](https://code.visualstudio.com/)
- Filen `.env` med din nyckel i repots rotmapp, från modul 01
- `npm install` körd i repots rot (för kapitel 6) och i `nordverk-assistent`
