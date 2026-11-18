# Exempelcase 2 – Sortera kundärenden och föreslå svar

> ⚠️ **Övningscase.** Företaget och all testdata är **påhittade**, även namn,
> adresser, telefonnummer och e-postadresser. Likheter med verkliga personer eller
> företag är oavsiktliga.

## Från företaget

*Ljungtorp Cykel AB* säljer cykeldelar i en webbutik. Vi är fyra personer på
kundtjänst och får ungefär 150 mejl per dag. Mycket är samma frågor: var är mitt
paket, hur returnerar jag, passar den här delen min cykel.

Vi vill att varje nytt ärende automatiskt får en **kategori**, en **prioritet** och
ett **förslag på svar**. En person på kundtjänst läser förslaget, ändrar och
skickar. Inget får skickas till kunden utan att en människa har godkänt det.

## Data

| Fil | Innehåll |
|---|---|
| `testdata/arenden.json` | 8 kundärenden med namn, e-post, telefon, ordernummer och text |
| `testdata/kategorier.md` | Våra kategorier och regler för prioritet |
| `testdata/returpolicy.md` | Våra villkor för retur, byte och reklamation |
| `testdata/svarsmall.md` | Hur vi skriver till kunder |

**Företagets egen bedömning:** Ärendena innehåller kunduppgifter. De får inte
spridas. Vi vet inte om det är okej att skicka dem till en AI-tjänst. Det vill vi
att ni utreder.

## Krav

1. Varje ärende får en av våra kategorier och en prioritet enligt `kategorier.md`.
2. Svarsförslagen följer `returpolicy.md` och `svarsmall.md` och lovar aldrig
   något som inte står där.
3. Ärenden som gäller säkerhet, pengar eller misstänkt bedrägeri ska alltid få
   prioritet Hög.
4. Lösningen ska klara att en kund försöker lura systemet med text i sitt mejl.
5. Så lite personuppgifter som möjligt ska lämna vår miljö.

## Framgångskriterier

- Minst 7 av 8 testärenden får samma kategori som vår kundtjänst satte.
- Inget svarsförslag lovar något som strider mot våra villkor.
- Kundtjänst sparar tid: ett förslag går att skicka efter små ändringar.

## Tips till gruppen

- Sortera de 8 ärendena själva först. Det blir ert facit.
- Behöver modellen kundens namn, e-post och telefonnummer för att göra jobbet?
- Läs ärende K-2210 noga.
