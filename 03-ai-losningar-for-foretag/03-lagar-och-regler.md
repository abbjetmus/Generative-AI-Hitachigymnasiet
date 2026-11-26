# 3. Lagar och regler

När en AI-lösning hanterar riktig data gäller riktiga lagar. Du behöver inte
kunna dem i detalj. Du behöver känna igen **när** en lag spelar roll, så att
du kan ställa rätt frågor till företaget. Länkarna leder till källorna om du
vill läsa mer.

> **Obs!** Kapitlet förenklar och är inte juridisk rådgivning. Läget gäller
> november 2026. Där det står **kontrollera aktuell status** kan något ha
> ändrats.

## Översikt

| Regel | Spelar roll när | Frågan du ställer |
|-------|-----------------|-------------------|
| GDPR | Det finns personuppgifter i datan | Får vi använda uppgifterna så här, och vart skickas de? |
| CLOUD Act | Leverantören är amerikansk | Kan amerikanska myndigheter begära ut datan? |
| AI-förordningen | Nästan alltid | Vilken riskklass? Måste användaren få veta att det är AI? |
| NIS2 | Företaget är viktigt för samhället | Passar AI-tjänsten in i företagets säkerhetsarbete? |
| Säkerhetsskydd | Verksamheten är viktig för Sveriges säkerhet | Får informationen lämna företaget alls? |
| Exportkontroll | Tekniken kan användas militärt | Räknas det som export att skicka den till en tjänst utomlands? |

## GDPR

[GDPR](https://eur-lex.europa.eu/legal-content/SV/TXT/?uri=CELEX:32016R0679)
är EU:s lag om personuppgifter. Den gäller så fort du skickar text med till
exempel ett namn till en språkmodell. I Sverige har
[IMY](https://www.imy.se/) tillsyn.

### Några grundregler i vardagliga ord

- **Ha ett skäl som lagen godkänner.** Lagen kallar det *rättslig grund*.
  Ett företag får till exempel använda en kunds uppgifter för att leverera
  det kunden har beställt.
- **Använd uppgifterna bara till det de samlades in för.** Kundärenden från
  supporten får inte plötsligt användas till reklam.
- **Använd så lite som möjligt.** Ta bort namn innan texten skickas till
  modellen (*uppgiftsminimering*).
- **Spara inte för länge.** Bestäm hur länge loggar med frågor och svar sparas.
- **Skydda uppgifterna** och kunna visa hur.
- **Låt inte en maskin ensam fatta viktiga beslut om människor.** Därför
  säger kursens tumregel att en människa fattar beslut som påverkar personer.

IMY och Digg har skrivit en vägledning om
[GDPR vid användning av generativ AI](https://www.imy.se/publikationer/gdpr-vid-anvandning-av-generativ-ai/).

### Avtal med leverantören

Företaget är **personuppgiftsansvarigt**. AI-leverantören som behandlar
uppgifterna åt företaget är **personuppgiftsbiträde**. Då ska det finnas ett
skriftligt **personuppgiftsbiträdesavtal** (PUB-avtal) som säger vad
leverantören får göra med uppgifterna
([IMY om PUB-avtal](https://www.imy.se/verksamhet/dataskydd/det-har-galler-enligt-gdpr/personuppgiftsansvariga-och-personuppgiftsbitraden/personuppgiftsbitradesavtal/)).

### Får datan skickas till USA?

Länder utanför EU/EES kallas **tredje land**. Personuppgifter får bara
skickas dit om det finns ett särskilt skydd.

- **2020** underkände EU-domstolen det dåvarande avtalet mellan EU och USA
  (domen kallas **Schrems II**). Skälet var att amerikanska
  underrättelsetjänster kunde komma åt uppgifterna
  ([IMY om Schrems II](https://www.imy.se/verksamhet/dataskydd/det-har-galler-enligt-gdpr/overforing-till-tredje-land/schrems-ii-domen-overforingar-till-tredje-land/)).
- **2023** kom ett nytt beslut, **EU–US Data Privacy Framework**. Det
  tillåter överföring till amerikanska företag som är **certifierade** enligt
  ramverket ([IMY om överföringar till USA](https://www.imy.se/verksamhet/dataskydd/det-har-galler-enligt-gdpr/overforing-till-tredje-land/overforingar-till-usa/)).
- **I dag (november 2026)** gäller ramverket fortfarande, men det är
  ifrågasatt. Ett försök att få det ogiltigförklarat förlorade 2025 men är
  överklagat till EU-domstolen, och organisationen
  [noyb](https://noyb.eu/en/us-supreme-court-just-blew-eu-us-data-transfers)
  bad 2026 EU-kommissionen att avveckla det. **Kontrollera aktuell status.**

Två tidigare avtal mellan EU och USA har redan underkänts. Därför vill många
företag inte bygga en AI-lösning som skickar personuppgifter till USA, även
om det är lagligt just nu.

## CLOUD Act

**CLOUD Act** är en amerikansk lag från 2018. Den säger att ett företag som
lyder under amerikansk lag måste lämna ut data till amerikanska myndigheter
när de begär det på rätt sätt, **även om datan ligger i ett annat land**
([lagtexten](https://www.law.cornell.edu/uscode/text/18/2713)).

Därför räcker det inte att fråga **var** datan ligger. Ett amerikanskt
molnföretag med datacenter i Sverige lyder ändå under amerikansk lag. Man
måste också fråga **vem** som äger leverantören.

Frågan är omdiskuterad. De amerikanska molnföretagen lovar i sina avtal att
bestrida begäranden de anser är felaktiga. Svenska leverantörer som Aixia
använder CLOUD Act som ett skäl att välja dem. Slutsatsen för dig: **ju
känsligare data, desto viktigare blir frågan om vem leverantören lyder
under.**

## AI-förordningen

[AI-förordningen](https://eur-lex.europa.eu/legal-content/SV/TXT/?uri=CELEX:32024R1689)
är EU:s lag om AI. Den delar in AI i riskklasser. Ju större risk för
människor, desto fler krav. Svensk information finns hos
[PTS](https://pts.se/ai/ai-forordningen/).

| Riskklass | Exempel | Vad som gäller |
|-----------|---------|----------------|
| **Förbjuden** | Känsloigenkänning av anställda eller elever, social poängsättning, att manipulera människor så att de skadas | Förbjudet |
| **Hög risk** | AI som sorterar jobbansökningar, sätter betyg, prövar lån, styr säkerheten i maskiner | Många krav: dokumentation, loggning, mänsklig tillsyn |
| **Transparens** | Chattbotar, AI-genererade bilder och texter | Användaren ska få veta att det är AI. Genererat innehåll ska märkas |
| **Minimal risk** | Stavningskontroll, spamfilter, de flesta interna verktyg | Inga särskilda krav |

Det här betyder för dina case:

- En **chattbot** som pratar med kunder ska tala om att den är en AI.
- En AI som **väljer ut jobbsökande** eller **bedömer elever** är högrisk.
  Säg till företaget att det kräver mycket mer än en skolprototyp.
- En AI som **sammanfattar interna rapporter** åt en människa som granskar
  dem är oftast minimal risk.
- Företag som använder AI ska se till att personalen förstår tekniken
  (*AI-kunnighet*).

### När gäller vad?

Förordningen gäller i steg. Tidplanen ändrades 2026
([förordning (EU) 2026/1744](https://eur-lex.europa.eu/eli/reg/2026/1744/oj)).

| Datum | Vad |
|-------|-----|
| 2 februari 2025 | Förbuden och kravet på AI-kunnighet |
| 2 augusti 2025 | Krav på dem som gör stora språkmodeller |
| 2 augusti 2026 | Transparenskraven, till exempel för chattbotar |
| 2 december 2027 | Högriskkraven för till exempel rekrytering och utbildning (flyttat från augusti 2026) |
| 2 augusti 2028 | Högriskkraven för AI inbyggd i produkter, till exempel maskiner |

**Kontrollera aktuell status** på [pts.se](https://pts.se/ai/ai-forordningen/).
Fler ändringar har diskuterats.

## NIS2 och cybersäkerhetslagen

Sedan 15 januari 2026 gäller
[cybersäkerhetslagen](https://data.riksdagen.se/dokument/sfs-2025-1506.html),
den svenska versionen av EU:s NIS2-direktiv. Den gäller medelstora och stora
företag i sektorer som energi, transport, vård och delar av industrin. De
måste arbeta systematiskt med IT-säkerhet och rapportera allvarliga
incidenter. En pumptillverkare som Nordverk kan omfattas om det är
tillräckligt stort. Då måste en ny AI-tjänst passa in i företagets
säkerhetsarbete. Fråga företaget om de omfattas.

## Säkerhetsskydd

Säkerhetsskydd gäller verksamheter som är viktiga för **Sveriges säkerhet**,
till exempel delar av försvarsindustrin och energiförsörjningen
([säkerhetsskyddslagen](https://lagen.nu/2018:585)). Sådan information får
inte lämnas till en leverantör utan särskild prövning, ibland tillsammans med
Säkerhetspolisen.

**Nämner företaget säkerhetsskydd är det nivå 4.** Använd bara påhittad
testdata och låt företaget bestämma om AI får användas alls.

## Exportkontroll

Viss teknik kan användas både civilt och militärt, till exempel vissa
pumpar, ventiler och styrsystem. Att föra ut sådan teknik ur EU kräver ofta
tillstånd från [Inspektionen för strategiska produkter (ISP)](https://isp.se/).
Export är inte bara att skicka en låda. Att **skicka ritningar eller
teknisk information elektroniskt** till ett land utanför EU kan också räknas
som export
([förordning (EU) 2021/821](https://eur-lex.europa.eu/legal-content/SV/TXT/?uri=CELEX:32021R0821)).
Att klistra in en sådan ritning i en AI-tjänst i USA kan alltså vara ett
problem. Det är nivå 4.

## Avtal

Även när ingen lag stoppar kan **avtal** göra det: sekretessavtal med
kunder, avtal om var kunddata får lagras, eller AI-leverantörens villkor om
datan får användas för att träna modeller. Fråga företaget.

## Sammanfattning

- GDPR: ha ett godkänt skäl, använd så lite data som möjligt, ha ett PUB-avtal och koll på om datan lämnar EU.
- Att skicka personuppgifter till USA är lagligt i dag via Data Privacy Framework, men ifrågasatt.
- CLOUD Act gör att det spelar roll vem som äger leverantören, inte bara var datan ligger.
- AI-förordningen: chattbotar ska säga att de är AI. Rekrytering och betygsättning är högrisk.
- Säkerhetsskydd och exportkontroll betyder nivå 4. Då bestämmer företaget.

Nästa kapitel: [Var körs modellen?](04-var-kors-modellen.md)
