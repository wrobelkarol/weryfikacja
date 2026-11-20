// Zapis i odczyt xlsx Geneteki na oryginalnym szablonie indeksy.xlsx (JSZip + ręczna edycja XML, bez SheetJS,
// żeby zostały style, komentarze w nagłówkach, słownik imion i arkusz Uwagi). Wypełniany jest tylko arkusz Zgony
// i tylko on jest zaznaczony (inaczej karty się grupują i zmiany trafiają też do Urodzeń).

import { KOLUMNY, LICZBOWE, doZapisu } from './konwersja.js';

const STYL_KOLUMNY = { A: 5 };                 // style kolumn z szablonu (<cols>): A = 5, reszta = 1
const STYL_DOMYSLNY = 1;

export function xmlEsc(s) {
  return String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
    // znaki sterujące są niedozwolone w XML
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
}

function xmlOdEsc(s) {
  return String(s).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(+n)).replace(/&#x([0-9a-f]+);/gi, (_, n) => String.fromCodePoint(parseInt(n, 16)))
    .replace(/&amp;/g, '&');
}

/** Ścieżki arkuszy w pliku: {nazwa arkusza: 'xl/worksheets/sheetN.xml'}. */
async function arkusze(zip) {
  const wb = await zip.file('xl/workbook.xml').async('string');
  const rels = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  const cele = {};
  for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
    const id = (m[0].match(/\bId="([^"]+)"/) || [])[1];
    const cel = (m[0].match(/\bTarget="([^"]+)"/) || [])[1];
    if (id && cel) cele[id] = cel.replace(/^\/?(xl\/)?/, 'xl/');
  }
  const wynik = {}, kolejnosc = [];
  for (const m of wb.matchAll(/<sheet\b[^>]*>/g)) {
    const nazwa = xmlOdEsc((m[0].match(/\bname="([^"]*)"/) || [])[1] || '');
    const id = (m[0].match(/\br:id="([^"]+)"/) || [])[1];
    wynik[nazwa] = cele[id];
    kolejnosc.push(nazwa);
  }
  return { sciezki: wynik, kolejnosc };
}

async function wspolneTeksty(zip) {
  const plik = zip.file('xl/sharedStrings.xml');
  if (!plik) return [];
  const xml = await plik.async('string');
  return [...xml.matchAll(/<si>([\s\S]*?)<\/si>/g)].map((m) =>
    [...m[1].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((t) => xmlOdEsc(t[1])).join(''));
}

function numerKolumny(litery) {
  let n = 0;
  for (const ch of litery) n = n * 26 + ch.charCodeAt(0) - 64;
  return n;
}

/** Odczyt arkusza: {numer wiersza: {litera kolumny: tekst}}. */
async function czytajArkusz(zip, sciezka, teksty) {
  const xml = await zip.file(sciezka).async('string');
  const wiersze = {};
  for (const r of xml.matchAll(/<row\b[^>]*?\br="(\d+)"[^>]*?(?:\/>|>([\s\S]*?)<\/row>)/g)) {
    const komorki = {};
    for (const c of (r[2] || '').matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
      const atr = c[1];
      const ref = (atr.match(/\br="([A-Z]+)\d+"/) || [])[1];
      const typ = (atr.match(/\bt="([^"]+)"/) || [])[1];
      const wnetrze = c[2] || '';
      let v = null;
      if (typ === 'inlineStr') v = [...wnetrze.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((t) => xmlOdEsc(t[1])).join('');
      else {
        const vv = (wnetrze.match(/<v>([\s\S]*?)<\/v>/) || [])[1];
        if (vv !== undefined) v = typ === 's' ? teksty[+vv] : xmlOdEsc(vv);
      }
      if (ref && v !== null && v !== '') komorki[ref] = v;
    }
    if (Object.keys(komorki).length) wiersze[+r[1]] = komorki;
  }
  return wiersze;
}

/** Wiersze arkusza „Słownik imion łacińskich” z szablonu: [mianownik, dopełniacz, biernik, skrót, polskie]. */
export async function czytajSlownikImion(JSZip, bufor) {
  const zip = await JSZip.loadAsync(bufor);
  const { sciezki } = await arkusze(zip);
  const nazwa = Object.keys(sciezki).find((n) => /słownik/i.test(n));
  if (!nazwa) return [];
  const w = await czytajArkusz(zip, sciezki[nazwa], await wspolneTeksty(zip));
  return Object.keys(w).map(Number).sort((a, b) => a - b).slice(1)
    .map((r) => ['A', 'B', 'C', 'D', 'E'].map((k) => w[r][k] || ''));
}

/** Wiersze danych arkusza Zgony (bez nagłówka i bez wierszy, w których jest tylko L.p.). */
export async function czytajZgony(JSZip, bufor) {
  const zip = await JSZip.loadAsync(bufor);
  const { sciezki } = await arkusze(zip);
  if (!sciezki.Zgony) throw new Error('W pliku nie ma arkusza „Zgony”.');
  const w = await czytajArkusz(zip, sciezki.Zgony, await wspolneTeksty(zip));
  return Object.keys(w).map(Number).sort((a, b) => a - b).filter((r) => r > 1)
    .map((r) => w[r])
    .filter((k) => Object.keys(k).some((kol) => kol !== 'A'))
    .map((k) => Object.fromEntries(KOLUMNY.map((kol) => [kol, k[kol] ?? ''])));
}

function komorka(kol, nrWiersza, v) {
  const t = doZapisu(kol, v);
  if (t === '') return '';
  const s = STYL_KOLUMNY[kol] ?? STYL_DOMYSLNY;
  const ref = `${kol}${nrWiersza}`;
  if (LICZBOWE.has(kol) && /^\d{1,9}$/.test(t)) return `<c r="${ref}" s="${s}"><v>${+t}</v></c>`;
  const zachowaj = /^\s|\s$/.test(t) ? ' xml:space="preserve"' : '';
  return `<c r="${ref}" s="${s}" t="inlineStr"><is><t${zachowaj}>${xmlEsc(t)}</t></is></c>`;
}

/**
 * Plik xlsx Geneteki: szablon + wiersze (obiekty {A..S}) w arkuszu Zgony. Kolumna A (L.p.) jest numerowana od 1.
 * Zwraca Uint8Array.
 */
export async function zbudujXlsx(JSZip, szablon, wiersze) {
  const zip = await JSZip.loadAsync(szablon);
  const { sciezki, kolejnosc } = await arkusze(zip);
  const sciezkaZgonow = sciezki.Zgony;
  if (!sciezkaZgonow) throw new Error('W szablonie nie ma arkusza „Zgony”.');

  // 1. tylko karta Zgony zaznaczona
  for (const [nazwa, sciezka] of Object.entries(sciezki)) {
    let xml = await zip.file(sciezka).async('string');
    xml = xml.replace(/<sheetView\b[^>]*?>/g, (tag) => {
      let t = tag.replace(/\s+tabSelected="[^"]*"/g, '');
      if (nazwa === 'Zgony') t = t.replace(/<sheetView\b/, '<sheetView tabSelected="1"');
      return t;
    });
    if (nazwa === 'Zgony') xml = wstawWiersze(xml, wiersze);
    zip.file(sciezka, xml);
  }
  let wb = await zip.file('xl/workbook.xml').async('string');
  const indeks = kolejnosc.indexOf('Zgony');
  wb = wb.replace(/<workbookView\b[^>]*?(\/?)>/, (tag, zamkniecie) => {
    let t = tag.replace(/\s+(activeTab|firstSheet)="[^"]*"/g, '');
    return t.replace(/\s*\/?>$/, ` activeTab="${indeks}"${zamkniecie ? '/' : ''}>`);
  });
  zip.file('xl/workbook.xml', wb);
  // usunięty łańcuch obliczeń, gdyby był (Excel odtworzy go sam)
  if (zip.file('xl/calcChain.xml')) zip.remove('xl/calcChain.xml');

  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
}

function wstawWiersze(xml, wiersze) {
  const m = xml.match(/<sheetData>([\s\S]*?)<\/sheetData>|<sheetData\s*\/>/);
  if (!m) throw new Error('Arkusz Zgony bez <sheetData>.');
  const tresc = m[1] || '';
  const naglowek = (tresc.match(/<row\b[^>]*\br="1"[\s\S]*?<\/row>/) || [''])[0];
  const wierszeXml = wiersze.map((w, i) => {
    const nr = i + 2;
    const kom = KOLUMNY.map((kol) => komorka(kol, nr, kol === 'A' ? String(i + 1) : w[kol])).join('');
    return `<row r="${nr}" spans="1:19">${kom}</row>`;
  }).join('');
  let wynik = xml.replace(m[0], `<sheetData>${naglowek}${wierszeXml}</sheetData>`);
  const ostatni = Math.max(1, wiersze.length + 1);
  wynik = wynik.replace(/<dimension ref="[^"]*"\s*\/>/, `<dimension ref="A1:S${ostatni}"/>`);
  // kursor na początku danych
  wynik = wynik.replace(/(<selection pane="bottomRight"[^>]*?)activeCell="[^"]*"([^>]*?)sqref="[^"]*"/, '$1activeCell="B2"$2sqref="B2"');
  return wynik;
}

