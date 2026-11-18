// Läser in Nordverks påhittade testdata från mappen data/nordverk i repots rot.
//
// import.meta.glob är en funktion i Vite. Den letar upp alla filer som matchar
// mönstret och läser in dem när appen startar. ?raw betyder "ge mig filen som
// vanlig text". Resultatet är ett objekt: { 'sökväg/till/fil.md': 'texten', ... }
const manualFiler = import.meta.glob('../../../data/nordverk/manualer/*.md', {
  query: '?raw',
  import: 'default',
  eager: true,
})
const protokollFiler = import.meta.glob('../../../data/nordverk/serviceprotokoll/*.md', {
  query: '?raw',
  import: 'default',
  eager: true,
})

// JSON-filer kan importeras direkt. Då blir de vanliga JavaScript-objekt.
import kundarendenFil from '../../../data/nordverk/kundarenden.json'

// Gör om { sökväg: text } till en lista [{ fil: 'nv-200.md', text: '...' }, ...]
function tillLista(filer) {
  return Object.entries(filer).map(([sokvag, text]) => ({
    fil: sokvag.split('/').pop(),
    text,
  }))
}

export const manualer = tillLista(manualFiler)
export const serviceprotokoll = tillLista(protokollFiler)
export const kundarenden = kundarendenFil
