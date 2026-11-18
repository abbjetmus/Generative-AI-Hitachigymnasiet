# startprojekt

Ett litet Vue 3-projekt för företagsprojektet. Du läser in testdatafiler (text,
Markdown, CSV eller JSON), väljer en instruktion och låter modellen skapa ett
resultat, till exempel en rapport. Det används i
[kapitel 2 – Bygga lösningen](../02-bygga-losningen.md).

## Starta

Gruppens nyckel ska ligga i `.env` i repots rotmapp. Öppna en terminal i den här
mappen och skriv:

```
npm install
npm run dev
```

Öppna <http://localhost:5173> i webbläsaren.

Koden finns i `src/App.vue`. Instruktionerna finns i `src/instruktioner.js`.

⚠️ Lägg aldrig företagets data i den här mappen. Välj filerna från en mapp utanför repot.
