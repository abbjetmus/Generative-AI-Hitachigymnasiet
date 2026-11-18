// Läser serviceprotokollen med vanlig kod, utan AI.
import { serviceprotokoll } from './nordverk.js'

// Gör om svenska tal som "5,5" och "1 450" till JavaScript-tal.
function tal(text) {
  return Number(text.replace(/\s/g, '').replace(',', '.'))
}

// Läser ett protokoll och returnerar ett objekt med fälten.
function lasProtokoll(text) {
  const falt = {}
  let delkostnad = 0

  for (const rad of text.split('\n')) {
    // "| Datum | 2026-10-01 |" blir ['', 'Datum', '2026-10-01', '']
    const celler = rad.split('|').map((cell) => cell.trim())

    // Rader med två celler: fält och värde
    if (celler.length === 4) {
      falt[celler[1]] = celler[2]
    }
    // Rader med fyra celler: reservdelar, "| 400-7070 | Kopplingsgummi | 1 | 540 |"
    if (celler.length === 6 && /^\d{3}-\d{4}$/.test(celler[1])) {
      delkostnad += tal(celler[3]) * tal(celler[4])
    }
  }

  return {
    nummer: text.match(/SP-\d{4}-\d{4}/)[0],
    datum: falt['Datum'],
    kund: falt['Kund'],
    pump: falt['Pump'],
    serienummer: falt['Serienummer'],
    tekniker: falt['Tekniker'],
    felkod: falt['Felkod'],
    timmar: tal(falt['Arbetstid'].replace('h', '')),
    status: falt['Status'],
    delkostnad,
    text, // hela texten, till AI:n
  }
}

export const protokoll = serviceprotokoll.map((p) => lasProtokoll(p.text))

// Räknar hur många gånger varje värde finns, t.ex. { 'NV-400': 5, 'NV-900': 4 }
function rakna(lista, falt) {
  const antal = {}
  for (const sak of lista) {
    antal[sak[falt]] = (antal[sak[falt]] ?? 0) + 1
  }
  return antal
}

// Nyckeltalen till månadsrapporten.
export function raknaNyckeltal(lista) {
  let timmar = 0
  let kostnad = 0
  for (const p of lista) {
    timmar += p.timmar
    kostnad += p.delkostnad
  }
  return {
    antalBesok: lista.length,
    perPumpmodell: rakna(lista, 'pump'),
    perFelkod: rakna(lista, 'felkod'),
    arbetstidTimmar: timmar,
    reservdelarKronor: kostnad,
    uppfoljningKravs: lista.filter((p) => p.status !== 'Avslutat').map((p) => p.nummer),
  }
}
