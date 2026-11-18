# Modul 03 – AI-lösningar för företag

I modul 01 och 02 lärde du dig hur språkmodeller fungerar och hur man bygger
med dem. I den här modulen lär du dig **tänka** som den som tar fram en
AI-lösning till ett riktigt företag. Det handlar mindre om kod och mer om
frågor: Vad ska lösningen göra? Vilken data rör det? Vilka lagar gäller? Får
datan skickas till USA, eller måste modellen köras på företagets egen dator?

Modulen förbereder dig för modul 04, där du får ett riktigt case från ett
företag. Innan du bygger något där ska du kunna utreda vilken säkerhetsnivå
caset har och välja en lösning som passar.

> **Obs!** Säkerhetsnivåerna och driftalternativen i modulen är skolans
> förenklade modell. De hjälper dig att ställa rätt frågor, men de är inte
> juridisk rådgivning. I ett riktigt projekt är det alltid företaget, ofta
> tillsammans med jurist, dataskyddsombud eller säkerhetsansvarig, som
> fattar beslutet.

## Det här kan du efter modulen

- Gå från ett behov hos ett företag till en genomtänkt AI-lösning, steg för steg
- Avgöra när AI är rätt verktyg och när vanlig kod eller ingen teknik alls är bättre
- Klassa data efter konfidentialitet, riktighet och tillgänglighet och placera den på skolans säkerhetsnivåer 1–4
- Känna igen personuppgifter, känsliga personuppgifter och företagshemligheter
- Förklara grunderna i GDPR, EU:s AI-förordning och varför CLOUD Act diskuteras
- Jämföra driftalternativen A–D (global moln-API, EU-moln, svensk leverantör, lokal modell) och välja ett som passar säkerhetsnivån
- Köra en lokal språkmodell med Ollama och samma kurskod
- Hitta och minska risker som prompt injection, dataläckage och hallucinationer
- Testa en AI-lösning med testfall och dokumentera den

## Upplägg

### Del 1 – Från behov till lösning

| Kapitel | Vad du gör |
|---------|------------|
| [01 – Från behov till AI-lösning](01-fran-behov-till-ai-losning.md) | Följer processen från första mötet med företaget till drift, och lär dig vilka frågor du ska ställa |
| [02 – Data och säkerhetsnivåer](02-data-och-sakerhetsnivaer.md) | Klassar data och placerar den på säkerhetsnivå 1–4 |

### Del 2 – Regler och var modellen körs

| Kapitel | Vad du gör |
|---------|------------|
| [03 – Lagar och regler](03-lagar-och-regler.md) | Lär dig grunderna i GDPR, CLOUD Act, AI-förordningen, NIS2, säkerhetsskydd och exportkontroll |
| [04 – Var körs modellen?](04-var-kors-modellen.md) | Jämför driftalternativ A–D och väljer alternativ utifrån säkerhetsnivå |
| [05 – Lokala modeller med Ollama](05-lokala-modeller-med-ollama.md) | Kör en språkmodell på din egen dator och jämför med Aixia |

### Del 3 – Risker och test

| Kapitel | Vad du gör |
|---------|------------|
| [06 – Risker, test och ansvar](06-risker-test-och-ansvar.md) | Angriper en naiv assistent i Vue-projektet [`riskverkstad`](riskverkstad/), täpper till hålen och skriver testfall |

Lösningsförslag till övningarna finns i [losningar.md](losningar.md).
Försök själv först!

Om skolan eller företaget använder Microsoft 365 kan du också läsa
[kapitel 7 i modul 02 – Agenter i Microsoft 365](../02-bygga-med-ai/07-agenter-i-microsoft-365.md).
Där ser du hur samma tankar om säkerhetsnivåer används när man bygger agenter
i Microsoft 365 i stället för med egen kod.

## Det gemensamma exemplet

Exemplen bygger på **Nordverk AB**, ett **påhittat** företag i Västerås som
tillverkar och servar industripumpar. Datan finns i
[`data/nordverk/`](../data/nordverk/): manualer, serviceprotokoll och
kundärenden. Alla kunder, personer och serienummer är påhittade.

## Det här behöver du

- Modul 01 och 02 (du ska kunna anropa en språkmodell från kod och från Vue)
- Node.js och repots `.env` med nyckeln du fått av läraren
- För kapitel 5: en dator där du får installera program, gärna med minst 16 GB minne. Har du inte det kan du läsa kapitlet och göra övningarna med en mindre modell, eller arbeta tillsammans med någon som har en sådan dator.
- Kod i modulen:
  - [`riskverkstad/`](riskverkstad/) – Vue-projekt för kapitel 6
  - [`kod/jamfor-modeller.js`](kod/jamfor-modeller.js) – skript som jämför modeller i kapitel 5
