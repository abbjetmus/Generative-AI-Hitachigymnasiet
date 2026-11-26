# 3. Demonstrationen

Sist visar ni lösningen för företaget. Demon är 10–15 minuter, och sedan får
företaget ställa frågor. Målet är inte att imponera, utan att företaget ska
förstå vad lösningen kan, vad den inte kan och vad som behövs för att gå vidare.

Planera demon i [mallen](mallar/demo.md).

## Upplägg

| Del | Tid | Vad ni visar |
|---|---|---|
| 1. Problemet | 1 min | Företagets problem med era ord. Visa att ni har förstått |
| 2. Utredningen | 2–3 min | Säkerhetsnivå, driftalternativ och varför. Viktigaste lagarna och riskerna |
| 3. Lösningen live | 4–5 min | Ett vanligt exempel från början till slut |
| 4. Testresultat | 2 min | Hur många rätt, vad gick fel, jämförelsen med en människa |
| 5. Begränsningar | 1–2 min | Vad lösningen inte klarar. Var en människa måste kontrollera |
| 6. Nästa steg | 1 min | Vad som behövs för att använda lösningen på riktigt |

Var ärlig om begränsningarna. Ett företag litar mer på en grupp som berättar vad
som inte fungerar än på en som säger att allt fungerar.

## Förberedelse

- Öva hela demon minst en gång med tidtagning.
- Bestäm vem som säger vad och vem som sköter datorn.
- Ha testdatan och instruktionen klara. Börja inte leta efter filer under demon.
- Kontrollera att nyckeln fungerar samma dag, och att ni kör mot rätt driftalternativ.
- Stäng chatt, mejl och flikar som inte ska synas på skärmen.
- Visa bara data som företaget har godkänt att ni visar, särskilt om fler än
  kontaktpersonen är med.

## Reservplan

Något krånglar nästan alltid. Förbered det här:

- **Nätet eller API:t fungerar inte:** ha ett sparat resultat (`resultat.md`) och
  skärmbilder eller en kort skärminspelning av en lyckad körning. Säg som det är:
  "Tjänsten svarar inte just nu, så vi visar en körning vi gjorde i går."
- **Modellen svarar långsamt:** starta körningen och berätta om testresultaten
  medan ni väntar.
- **Modellen svarar fel:** bra tillfälle att visa varför en människa granskar. Visa
  hur felet hade upptäckts.
- **Lokal modell:** kör den på datorn ni visar på, så behövs inget nät.

## Frågor att vänta sig

- Vad kostar det att använda lösningen?
- Var hamnar vår data? Sparas den hos AI-tjänsten?
- Hur ofta blir det fel? Hur märker vi det?
- Kan vi använda den här på riktigt från i morgon? Vad saknas?
- Vad händer om en anställd skriver in något känsligt?
- Fungerar den med våra riktiga data, och med mer data?

Vet ni inte svaret: säg det, och lova att återkomma.

## Efter demon

Skriv en kort reflektion i gruppens repo: vad företaget tyckte, vad ni skulle göra
annorlunda och vad ni lärde er. Radera företagets data enligt utredningen.
