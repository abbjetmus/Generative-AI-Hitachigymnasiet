# nordverk-assistent

Vue 3-projektet som du bygger vidare på i modul 02,
[Bygga med AI](../README.md). Det läser Nordverks påhittade testdata från
[`data/nordverk`](../../data/nordverk/) i repots rot.

## Starta

Du behöver filen `.env` med din nyckel i repots rotmapp (se modul 01). Öppna
sedan en terminal i den här mappen och skriv:

```
npm install
npm run dev
```

Öppna <http://localhost:5173> i webbläsaren.

| Fil | Innehåll |
|---|---|
| `src/App.vue` | Startsidan. Här lägger du till komponenterna från kapitlen. |
| `src/ai.js` | `chatta()` och `client`. Pratar med AI-tjänsten via `/ai` på utvecklingsservern. |
| `src/nordverk.js` | Läser in manualer, serviceprotokoll och kundärenden. |
| `src/protokoll.js` | Läser fälten i serviceprotokollen och räknar nyckeltal, utan AI. |

Nyckeln läggs på av utvecklingsservern (se `gemensam/ai-vite.js`). Den finns
aldrig i webbläsaren. Det fungerar bara med `npm run dev`.
