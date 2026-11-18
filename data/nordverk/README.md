# Testdata: Nordverk AB

**Nordverk AB finns inte.** Företaget, kunderna, personerna, pumparna,
serienumren, priserna och all text i den här mappen är påhittade för kursen.
Likheter med verkliga företag eller personer är oavsiktliga. Datan räknas som
säkerhetsnivå 1 (öppen) och får användas med alla driftalternativ.

I kursen är Nordverk ett företag i Västerås som tillverkar och servar
industripumpar. Du bygger stegvis *Nordverk-assistenten*, som läser datan här.

## Innehåll

| Fil eller mapp | Innehåll | Används i |
|---|---|---|
| [`manualer/`](manualer/) | Manualer för pumparna NV-200, NV-400 och NV-900: tekniska data, säkerhet, underhållsintervall, felkoder och reservdelar | RAG, verktyget som slår upp felkoder, MCP |
| [`serviceprotokoll/`](serviceprotokoll/) | 12 serviceprotokoll från oktober 2026: datum, kund, pump, serienummer, tekniker, felkod, åtgärder, reservdelar och arbetstid | Månadsrapporten, verktyget som hämtar servicehistorik, MCP |
| [`kundarenden.json`](kundarenden.json) | 10 kundärenden från oktober och november 2026: frågor, klagomål och beställningar | Klassning, verktygsanrop och agenter |

## Format

Varje serviceprotokoll börjar med en tabell med fasta fält (`| Fält | Värde |`).
Det gör att ett program kan läsa fälten utan AI. Resten av protokollet är fri
text som teknikern har skrivit. Reservdelarna står i en tabell med artikelnummer,
antal och à-pris i kronor. Tal skrivs på svenskt sätt, till exempel `5,5 h` och
`1 450` kr.

Ett av kundärendena innehåller ett försök att lura en AI-assistent (en så
kallad *prompt injection*). Det finns där med flit och används i kapitlet om
agenter.
