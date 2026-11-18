# Exempelcase 3 – Frågor om teknisk dokumentation (exportkontroll)

> ⚠️ **Övningscase.** Företaget, produkten och all testdata är **påhittade**.
> Märkningen "exportkontrollerad" i dokumenten är en del av övningen. Inga
> uppgifter i dokumenten är verklig teknik.

## Från företaget

*Norrhammar Flödesteknik AB* tillverkar högtryckspumpar för processindustrin.
Våra montörer och servicetekniker letar ofta i långa dokument efter
åtdragningsmoment, intervall och felsökning. Vi vill ha en assistent där de kan
ställa frågor och få svar med hänvisning till rätt dokument.

En del av vår teknik omfattas av **exportkontroll**. Pumpar och
konstruktionsunderlag är exempel på sådant som kan vara produkter med dubbla
användningsområden, alltså produkter som har civil användning men också kan
användas militärt. Vår exportkontrollansvarige har sagt att dokumenten **inte
får lämna företagets nät**, och inte heller får skickas till molntjänster.

## Data

| Fil | Innehåll |
|---|---|
| `testdata/dokument/` | Produktblad, montering och underhåll för pumpen HX-40 |
| `testdata/testfragor.md` | Åtta testfrågor med förväntade svar |

**Företagets egen bedömning:** Dokumenten om HX-40 K är exportkontrollerade.
Resten är interna.

## Krav

1. Dokumenten och frågorna får inte lämna företagets nät. Ingen molntjänst.
2. Svaren ska hänvisa till dokument-ID.
3. Finns svaret inte i dokumenten ska assistenten säga det och inte gissa.
4. Lösningen ska fungera på en vanlig arbetsdator eller en server i vår serverhall.

## Framgångskriterier

- Minst 5 av testfrågorna 1–6 får rätt svar med rätt källa.
- Frågorna 7 och 8 besvaras med att uppgiften saknas.
- Svarstiden är högst en minut per fråga.

## Tips till gruppen

- Läs om lokala modeller i [05-lokala-modeller-med-ollama.md](../../../03-ai-losningar-for-foretag/05-lokala-modeller-med-ollama.md).
- Testdatan är påhittad och därför i sig på nivå 1. Öva ändå som om den vore
  riktig. Kör helst mot en lokal modell. Om ingen dator i gruppen klarar det får ni
  testa mot Aixia, **bara** för att datan är påhittad, men skriv i utredningen att
  den riktiga lösningen bara får köras lokalt.
- Läs mer om exportkontroll hos [Inspektionen för strategiska produkter (ISP)](https://isp.se/produkter-med-dubbla-anvandningsomraden/).
