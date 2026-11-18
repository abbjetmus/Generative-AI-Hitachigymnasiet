# riskverkstad

Ett litet Vue 3-projekt med en naiv ärendeassistent för Nordverk AB
(påhittat företag). Det används i
[kapitel 6 – Risker, test och ansvar](../06-risker-test-och-ansvar.md).

## Starta

Du behöver `.env` med din nyckel i repots rotmapp (se `.env.example`).
Öppna sedan en terminal i den här mappen och skriv:

```
npm install
npm run dev
```

Öppna <http://localhost:5173> i webbläsaren.

All kod du skriver hamnar i `src/App.vue`. Exempelärendena ligger i
`src/arenden.js`. Den färdiga versionen efter kapitel 6 finns i
`facit/App-kapitel-6.vue`. Kopiera den till `src/App.vue` om du vill prova
den.
