# 5. Lokala modeller med Ollama

I det här kapitlet kör du en språkmodell på din egen dator, alternativ D
från kapitel 4. Då lämnar ingen data datorn. Programmet heter
[Ollama](https://ollama.com/). Det bästa: **kurskoden behöver inte ändras.**
Ollama har ett API som fungerar som OpenAI:s, precis som Aixia. Du byter
bara några rader i `.env`.

> Har du inte en tillräckligt stark dator kan du läsa kapitlet, prova en
> liten modell eller arbeta tillsammans med någon som har en sådan dator.
> Fråga läraren om skolans datorer får användas.

## Vad krävs av datorn?

En modell måste få plats i datorns minne. På en Mac med M-chip delar
processorn och grafikdelen på minnet. På en Windows-dator går det snabbast
om modellen får plats i grafikkortets minne (VRAM).

| Modell | Nedladdning | Ungefär vad som krävs | Kommentar |
|--------|-------------|------------------------|-----------|
| `gpt-oss:20b` | 14 GB | Minst 16 GB minne | Lillasyster till `gpt-oss-120b` som Aixia kör. Resonerar, bra på svenska |
| `qwen3:8b` | 5,2 GB | Ca 8 GB minne | Mindre och snabbare |
| `qwen3:4b` | 2,5 GB | Går på de flesta datorer | Märkbart sämre, men bra för att prova |
| `bge-m3` | 1,2 GB | Liten | Samma embeddingmodell som Aixia kör |

Storlekarna kommer från Ollamas modellsidor
([gpt-oss](https://ollama.com/library/gpt-oss),
[qwen3](https://ollama.com/library/qwen3),
[bge-m3](https://ollama.com/library/bge-m3)). Ollama säger att `gpt-oss:20b`
kan köras med så lite som 16 GB minne. Räkna med att det går långsamt om
modellen inte får plats i grafikminnet. Nya modeller kommer hela tiden, så
titta gärna i [Ollamas modellbibliotek](https://ollama.com/library).

## Installera Ollama

Ladda ner från [ollama.com/download](https://ollama.com/download).

- **macOS:** kräver macOS 14 (Sonoma) eller nyare. Öppna filen och dra
  Ollama till Program. Starta appen.
- **Windows:** kräver Windows 10 22H2 eller nyare. Kör `OllamaSetup.exe`.
  Det behövs inga administratörsrättigheter.
- **Linux:** kör `curl -fsSL https://ollama.com/install.sh | sh`

Efter installationen kör Ollama i bakgrunden och lyssnar på
`http://localhost:11434`. Kontrollera i en terminal:

```
ollama --version
```

## Ladda ner och prova en modell

```
ollama pull gpt-oss:20b
ollama pull bge-m3
ollama run gpt-oss:20b
```

`pull` laddar ner modellen. `run` startar en chatt i terminalen. Skriv
`/bye` för att avsluta. Några fler kommandon:

| Kommando | Gör |
|----------|-----|
| `ollama list` | Visar nedladdade modeller |
| `ollama ps` | Visar vilka modeller som är laddade och om de körs på grafikkortet (`100% GPU`) eller processorn |
| `ollama rm qwen3:4b` | Tar bort en modell |

> **Varning:** modeller vars namn slutar på `-cloud`, till exempel
> `gpt-oss:20b-cloud`, körs **inte** på din dator utan på Ollamas servrar.
> Då lämnar datan datorn. Använd inte dem för alternativ D.

## Kör kurskoden mot Ollama

Öppna `.env` i repots rotmapp. Kommentera bort Aixia-raderna med `#` och
lägg till Ollama-raderna (de finns redan i `.env.example`):

```
# Aixia
# AI_BASE_URL=https://llm.aiqu.ai/v1
# AI_API_KEY=sk-...
# AI_MODEL=gpt-oss-120b
# AI_EMBED_MODEL=bge-m3

# Ollama
AI_BASE_URL=http://localhost:11434/v1
AI_API_KEY=ollama
AI_MODEL=gpt-oss:20b
AI_EMBED_MODEL=bge-m3
```

Ollama bryr sig inte om nyckeln, men kurskoden kräver att det står något,
därför `ollama`. Nu går alla kurskodens anrop till din dator i stället för
till Aixia. Det fungerar eftersom Ollama har samma typ av API
([Ollamas dokumentation om OpenAI-kompatibilitet](https://docs.ollama.com/api/openai-compatibility)):
chatt, streaming, embeddings och `response_format`.

Kör ett Vue-projekt, till exempel [`riskverkstad`](riskverkstad/)? Starta om
`npm run dev` efter att du ändrat `.env`, eftersom servern läser filen när
den startar.

Byt tillbaka till Aixia genom att göra tvärtom.

## Jämför Aixia och Ollama

Skriptet [`kod/jamfor-modeller.js`](kod/jamfor-modeller.js) kör tre
uppgifter mot den modell som står i `.env`:

1. svara på en fråga utifrån ett utdrag ur en manual,
2. sammanfatta ett kundärende,
3. bedöma prioritet och svara med JSON enligt ett schema.

Varje uppgift har en kontroll som säger om svaret är godkänt. Skriptet
mäter också tiden och hur många **token per sekund** modellen skriver, och
provar att skapa en embedding. Kör det från repots rotmapp:

```
node 03-ai-losningar-for-foretag/kod/jamfor-modeller.js
```

Så här såg det ut mot Aixia när kapitlet skrevs:

```
Tjänst: https://llm.aiqu.ai/v1
Modell: gpt-oss-120b

✅ Svara utifrån manualen: 1.2 s, 148 token, 122 token/s
   Avlufta pumpen.

✅ Sammanfatta ett kundärende: 1.4 s, 190 token, 132 token/s
   Kunden rapporterar att NV‑400 på linje 2 har stannat två gånger ...

✅ Klassa prioritet som JSON: 1.2 s, 153 token, 129 token/s
   { "prioritet": "akut" }

Embedding med bge-m3: 1024 dimensioner, 0.1 s

Resultat: 3 av 3 godkända, total tid 3.8 s
```

Byt `.env` till Ollama och kör igen. Jämför:

- **Kvalitet:** blev lika många uppgifter godkända? Läs svaren. Är svenskan lika bra?
- **Hastighet:** hur många token per sekund? Hur lång tid tog den första
  frågan? (Första gången laddas modellen in i minnet, det tar extra tid.)
- **Fungerar allt?** Särskilt JSON-uppgiften och embeddings.

Tre uppgifter är för få för att säga säkert vilken modell som är bäst. I
kapitel 6 lär du dig skriva fler testfall.

## Två saker att känna till

**Kontextlängd.** Ollama ger modellen ett kortare "minne" (kontext) som
standard om datorn har lite grafikminne, till exempel 4 000 token på en
dator med under 24 GB
([Ollamas dokumentation](https://docs.ollama.com/context-length)). Skickar
du med långa dokument kan början klippas bort utan varning. Öka
kontextlängden i Ollama-appens inställningar, eller starta servern med
`OLLAMA_CONTEXT_LENGTH=16000 ollama serve`. Kontrollera med `ollama ps`.

**Bara din dator.** Ollama lyssnar som standard bara på din egen dator.
Ändra inte det (inställningen `OLLAMA_HOST`) om du inte vet vad du gör, för
då kan andra i nätverket använda din modell.

## Sammanfattning

- Ollama kör språkmodeller lokalt. Ingen data lämnar datorn, om du inte använder `-cloud`-modeller.
- `gpt-oss:20b` kräver minst 16 GB minne. Mindre modeller går fortare men blir sämre.
- Byt `AI_BASE_URL`, `AI_API_KEY` och `AI_MODEL` i `.env`, så kör samma kurskod mot Ollama.
- Jämför kvalitet och hastighet med samma uppgifter mot båda.

Nästa kapitel: [Risker, test och ansvar](06-risker-test-och-ansvar.md)
