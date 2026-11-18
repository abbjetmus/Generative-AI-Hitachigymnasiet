# Exempelcase 1 – Månadsrapport från serviceprotokoll

> ⚠️ **Övningscase.** Företaget, kunden och all testdata är **påhittade**. Likheter
> med verkliga företag är oavsiktliga. Testdatan får användas fritt i kursen.

## Från företaget

*Tolvåker Fastighetsteknik AB* servar ventilationsaggregat (FTX) och värmepumpar
i bostadshus. Vår största kund är *Fastighets AB Kvarnbacken* med sex fastigheter.

Varje månad skriver vår driftchef en rapport till kunden och till vår egen
ledning. Driftchefen läser alla serviceprotokoll och driftlarm, räknar och
skriver en text. Det tar ungefär en halv arbetsdag. Protokollen skrivs av olika
tekniker och ser lite olika ut.

Vi vill ha ett verktyg som tar månadens protokoll och larmlista och skriver ett
**utkast** till månadsrapporten i samma form som de rapporter vi redan skickar.
Driftchefen läser alltid igenom och ändrar innan rapporten skickas.

## Data

| Fil | Innehåll |
|---|---|
| `testdata/serviceprotokoll/` | 6 serviceprotokoll, de flesta från oktober |
| `testdata/driftlarm-oktober.csv` | Driftlarm från styrsystemet (semikolon som avgränsare) |
| `testdata/larmkoder.md` | Vad larmkoderna betyder |
| `testdata/rapport-september.md` | Förra månadens rapport. Så här vill vi att rapporten ser ut |

Protokollen innehåller inga kundnamn eller personnamn. Teknikerna anges med
ett ID (T01–T04). Fastigheterna har namn som Björken 4 och Lönnen 2.

**Företagets egen bedömning:** Uppgifterna är interna. De är inte hemliga, men
vi vill inte att de publiceras.

## Krav

1. Rapporten ska följa samma rubriker som `rapport-september.md`.
2. Siffror (antal besök, arbetstid, antal larm) ska stämma med underlaget.
3. Rapporten ska ta upp problem som återkommer, även över månadsskiftet.
4. Om något saknas eller är oklart i underlaget ska rapporten säga det, inte gissa.
5. Det ska gå att använda verktyget utan att kunna programmera.

## Framgångskriterier

- Driftchefen behöver ändra högst några meningar innan rapporten kan skickas.
- Alla nyckeltal stämmer när vi kontrollräknar för hand.
- Det tar mindre än 15 minuter från färdiga protokoll till granskat utkast.

## Tips till gruppen

- Läs igenom protokollen själva först. Vad är svårt att få rätt?
- Låt en människa räkna nyckeltalen en gång. Det blir ert facit.
- Fundera på om siffrorna ska räknas av språkmodellen eller av vanlig kod.
