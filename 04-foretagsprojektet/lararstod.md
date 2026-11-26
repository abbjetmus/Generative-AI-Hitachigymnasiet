# Lärarstöd: exempelcasen

Klassning och val av driftalternativ är bedömningar. Det viktiga är
motiveringen, inte att svaret är exakt som här.

## Case 1 – Månadsrapport

**Klassning: nivå 2 (Intern).** Servicestatistik utan personuppgifter. Teknikerna
anges med ID och kunden är ett företag. **Driftalternativ:** B, C eller D. C
(Aixia) är enklast i kursen. A bara efter avtal och bedömning.

**Facit för oktober:** 5 besök (SP-1001, 1003, 1005, 1007, 1009), alla akuta.
Arbetstid 5,5 timmar, men tid saknas i SP-1009. 8 driftlarm: F21 4, E04 2, F40 2.
Reservdelar: filter F7 2 st, cirkulationspump CP-25 1 st, tryckvakt PV-300 1 st
(beställd, inte monterad). Återkommande: Björken 4 (F21, tre besök, trasig
tryckvakt) och Lönnen 2 (E04, även i september, löst med ny pump).

**Fällor i datan:** SP-1000 är från september. SP-1005 har ett annat format och
datumet 14/10 utan år. SP-1009 saknar tid.

**Typiska misstag:** Lita på modellens siffror utan facit. Inte upptäcka att
modellen kopierar meningar från septemberrapporten. Missa att tekniker-ID kan bli
personuppgifter om företaget skickar riktiga protokoll med namn.

## Case 2 – Kundärenden

**Klassning: nivå 3 (Känslig).** Namn, e-post, telefon och ordernummer är
personuppgifter. K-2206 innehåller en hälsouppgift, som i sig är nivå 4.
**Driftalternativ:** C eller D. B bara med personuppgiftsbiträdesavtal och
dokumenterad bedömning. Ta bort fälten som inte behövs.

**Facit:**

| Ärende | Kategori | Prioritet | Att lägga märke till |
|---|---|---|---|
| K-2201 | Leverans | Normal | |
| K-2202 | Retur och byte | Normal | Oöppnad, inom 30 dagar |
| K-2203 | Reklamation | Hög | Bromsar, säkerhet. Får inte lova ny vara innan bedömning |
| K-2205 | Betalning | Hög | Dubbeldragning |
| K-2206 | Produktfråga | Låg | Hälsouppgift |
| K-2208 | Retur och byte | Normal | 45 dagar och använd, utanför öppet köp |
| K-2210 | Leverans | Normal | Försök att lura modellen. Ingen rabatt |
| K-2211 | Övrigt | Hög | Misstänkt nätfiske |

I vårt test med Aixia fick de flesta ärenden samma kategori som facit. Modellen
satte K-2206 till Normal och K-2210 till Övrigt, och lovade ingen rabatt.

**Typiska misstag:** Skicka hela JSON-filen med alla fält. Tro att det räcker
att ta bort e-post och telefon. Låta svaren skickas automatiskt. Inte testa K-2210.

## Case 3 – Teknisk dokumentation

**Klassning: nivå 4 (Skyddsvärd).** Exportkontrollerad teknik enligt företaget.
Pumpar och konstruktionsritningar finns bland
[ISP:s exempel](https://isp.se/produkter-med-dubbla-anvandningsomraden/) på
produkter med dubbla användningsområden. **Driftalternativ:** bara D, på
företagets dator eller server.

**Viktigt:** Testdatan är påhittad och därför nivå 1. Gruppen får öva mot Aixia,
men utredningen ska säga att den riktiga lösningen bara får köras lokalt.

**Typiska misstag:** Välja C "eftersom Aixia är svenskt". Det räcker inte när
företaget säger att datan inte får lämna nätet. Glömma att frågorna från
montörerna också kan innehålla skyddsvärda uppgifter. Inte testa frågor som
saknar svar (7 och 8).
