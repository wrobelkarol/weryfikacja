'use strict';
/* Weryfikacja odczytów: wycinki z folderu "pocięte" + odczyt AI (łacina i polski) do poprawy, eksport xlsx.
   Dane trzyma serwer (weryfikacja.pyw); poprawki zapisują się same do odczyt_zweryfikowany.json.
   Na stronie www dla pomocnika serwer zastępuje www.js (window.WWW): te same adresy /api/..., dane z GitHuba. */

const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => Array.from(r.querySelectorAll(s));
const enc = encodeURIComponent;
const WWW = window.WWW || null;

function el(tag, atrybuty = {}, ...dzieci) {
  const e = document.createElement(tag);
  for (const [k, v] of Object.entries(atrybuty)) {
    if (v === null || v === undefined || v === false) continue;
    if (k === 'class') e.className = v;
    else if (k === 'text') e.textContent = v;
    else if (k.startsWith('on')) e.addEventListener(k.slice(2), v);
    else if (k in e && typeof v !== 'string') e[k] = v;
    else e.setAttribute(k, v === true ? '' : v);
  }
  for (const d of dzieci.flat()) if (d !== null && d !== undefined && d !== false) e.append(d);
  return e;
}

const API = {
  async get(url) {
    if (WWW) return WWW.get(url);
    const r = await fetch(url);
    const d = await r.json().catch(() => ({ blad: `błąd ${r.status}` }));
    if (!r.ok) throw new Error(d.blad || `błąd ${r.status}`);
    return d;
  },
  async post(url, cialo) {
    if (WWW) return WWW.post(url, cialo);
    const r = await fetch(url, {
      method: 'POST', body: JSON.stringify(cialo),
      headers: { 'Content-Type': 'application/json', 'X-Weryfikacja': '1' },
    });
    if (!r.ok) {
      const d = await r.json().catch(() => ({}));
      throw new Error(d.blad || `błąd ${r.status}`);
    }
    return r;
  },
};

let POLA = [], TYPY = {}, RODZAJE = {};
let POLA_KSIAG = { zgony: [], urodzenia: [] };
/* Księga urodzeń ma te same klucze dat i dziecka co zgony (dzien_zgonu = urodzenie, dzien_pogrzebu = chrzest,
   zm_* = dziecko), więc reguły dat działają tak samo - zmieniają się tylko słowa w komunikatach. */
let URODZENIA = false;
const SLOWA_ZGONY = { zdarzenia: 'zgonu', przed: 'zgonem', po: 'zgonie', drugie: 'pogrzeb', drugiego: 'pogrzebu' };
const SLOWA_URODZENIA = { zdarzenia: 'urodzenia', przed: 'urodzeniem', po: 'urodzeniu', drugie: 'chrzest', drugiego: 'chrztu' };
let S = SLOWA_ZGONY;

/* Format Geneteki (zgony): pod łaciną każdy akt jest od razu przeliczony na kolumny Geneteki A–S z polskiego
   odczytu AI – tym samym kodem co w programie „genetyka weryfikacja” (konwersja.js, xlsx.js). Ręczne poprawki
   kolumn zapisują się przy wycinku (pole „geneteka”), a „Pobierz xlsx genetyka” zapisuje plik dla Geneteki
   w folderze rocznika. Księga urodzeń zostaje po staremu (łacina i polski obok siebie, dawny xlsx). */
let K = null, XL = null, SZABLON = null, SLOWNIK = null;
let GENETEKA = false;
const NAZWY_PARAFII = { Przytkowice: 'Przytkowice-Z-przenajswietszej-trojcy' };
const SZEROKIE_GEN = new Set(['M']);
const SREDNIE_GEN = new Set(['L', 'H', 'I']);
// pola łaciny widoczne zawsze; pozostałe puste są schowane (przycisk „Pokaż puste pola”)
const ZAWSZE_WIDOCZNE = new Set(['nr', 'miesiac', 'dzien_zgonu', 'dzien_pogrzebu', 'dom', 'katolik', 'plec', 'wiek',
  'przyczyna', 'typ', 'zm_imie', 'zm_nazwisko', 'tekst']);

function ustawKsiege(ksiega) {
  URODZENIA = /^urodz/i.test(ksiega || '');
  GENETEKA = !!K && !URODZENIA;
  S = URODZENIA ? SLOWA_URODZENIA : SLOWA_ZGONY;
  POLA = URODZENIA && POLA_KSIAG.urodzenia.length ? POLA_KSIAG.urodzenia : POLA_KSIAG.zgony;
  for (const k of Object.keys(RODZAJE_POL)) delete RODZAJE_POL[k];
  for (const [k, , r] of POLA) RODZAJE_POL[k] = r;
}
const st = {
  sciezka: null, opis: null, wycinki: [], mapa: new Map(), flagi: new Map(), aiZmieniony: null,
  brudne: false, zapisTimer: null, flagiTimer: null, otwarte: true, filtr: 'wszystkie', drzewo: null, edycja: null,
  linki: [],             // linki do skanów: [{od, do, url}] (pokazLinkiSkanow)
  kolumna: 'lacina',     // kolumna, w której zostaje strzałka w dół/górę po przejściu przez pole wspólne
  gen: new Map(),        // wycinek -> {wiersz (kolumny A–S), flagi, podpowiedzi} liczone z odczytu (przeliczGeneteke)
  pokazPuste: false,     // puste pola łaciny widoczne
};
const RODZAJE_POL = {};

// ---------------------------------------------------------------- pomocnicze: daty, wiek, imiona

const MIESIACE = {
  januarius: 1, januarii: 1, jan: 1, styczen: 1, 'styczeń': 1, februarius: 2, februarii: 2, feb: 2, febr: 2, luty: 2,
  martius: 3, martii: 3, mart: 3, marzec: 3, aprilis: 4, apr: 4, kwiecien: 4, 'kwiecień': 4, majus: 5, maius: 5,
  maji: 5, maj: 5, junius: 6, iunius: 6, junii: 6, jun: 6, czerwiec: 6, julius: 7, iulius: 7, julii: 7, jul: 7,
  lipiec: 7, augustus: 8, augusti: 8, aug: 8, sierpien: 8, 'sierpień': 8, september: 9, septembris: 9, sept: 9,
  '7ber': 9, wrzesien: 9, 'wrzesień': 9, october: 10, octobris: 10, oct: 10, '8ber': 10, pazdziernik: 10,
  'październik': 10, november: 11, novembris: 11, nov: 11, '9ber': 11, listopad: 11, december: 12,
  decembris: 12, dec: 12, '10ber': 12, xber: 12, grudzien: 12, 'grudzień': 12,
};

function miesiacZ(v) {
  const s = (v || '').toLowerCase().replace(/[^a-ząćęłńóśźż0-9]/g, '');
  if (!s) return null;
  if (/^\d+$/.test(s)) { const n = +s; return n >= 1 && n <= 12 ? n : null; }
  if (MIESIACE[s]) return MIESIACE[s];
  for (const [nazwa, n] of Object.entries(MIESIACE)) if (s.length >= 3 && /^[a-z]/.test(nazwa) && nazwa.startsWith(s.slice(0, 3))) return n;
  return null;
}

function liczba(v) {
  const m = /^\s*(\d+)\.?\s*$/.exec(v || '');
  return m ? +m[1] : null;
}

function dzienMiesiac(v) {   // '28' -> [28, null]; '28/2' -> [28, 2]
  const m = /^\s*(\d{1,2})\s*(?:\/\s*(\d{1,2}))?\s*\.?\s*$/.exec(v || '');
  return m ? [+m[1], m[2] ? +m[2] : null] : [null, null];
}

function data(rok, miesiac, dzien) {
  const d = new Date(rok, miesiac - 1, dzien);
  return d.getFullYear() === rok && d.getMonth() === miesiac - 1 && d.getDate() === dzien ? d : null;
}

function wiekWLatach(v) {
  let s = (v || '').toLowerCase().replace(/,/g, '.').trim();
  if (!s || /^[—–-]+$/.test(s)) return null;
  if (/\bn\.|ur\.|natus|nata|\b1[789]\d\d\b/.test(s)) return null;      // data urodzenia zamiast wieku
  s = s.replace(/½/g, ' 1/2').replace(/¼/g, ' 1/4').replace(/¾/g, ' 3/4').replace(/⅓/g, ' 1/3').replace(/⅔/g, ' 2/3');
  const re = /(\d+(?:\.\d+)?)(?:\s+(\d+)\s*\/\s*(\d+))?\s*\.?\s*([a-ząćęłńóśźż]*)/g;
  let suma = 0, ile = 0, m;
  while ((m = re.exec(s))) {
    let w = parseFloat(m[1]);
    if (m[2] && +m[3]) w += m[2] / m[3];
    const j = m[4] || '';
    const mnoznik = /^(mies|mens|m$)/.test(j) ? 1 / 12 : /^(tyg|hebd)/.test(j) ? 7 / 365.25
      : /^(dn|dzie|dier|dies|die|d$)/.test(j) ? 1 / 365.25 : 1;
    suma += w * mnoznik;
    ile++;
  }
  return ile ? suma : null;
}

// <wiek-polski> – ta sama zasada co w programie „genetyka weryfikacja” (konwersja.js, funkcja wiek):
// wiek bez ułamków zwykłych, zaokrąglony do całości albo do ,5 (połówki w górę), z polską odmianą
// („40½ lat” -> „40,5 roku”, „1 10/12 roku” -> „2 lata”, „4 lata 3 miesiące” -> „4,5 roku”).
// Data urodzenia zamiast wieku, „—” i nietypowe zapisy zostają bez zmian.
const WIEK_ULAMKI = { '½': 0.5, '¼': 0.25, '¾': 0.75, '⅓': 1 / 3, '⅔': 2 / 3, '⅙': 1 / 6, '⅚': 5 / 6, '⅛': 0.125 };
const WIEK_JEDNOSTKI = [
  { re: /^(rok|roku|lat|lata|latek|r$|ann)/i, formy: ['rok', 'lata', 'lat', 'roku'], dni: 365.25 },
  { re: /^(miesiąc|miesiąca|miesiące|miesięcy|mies|m-c|mens|m$)/i, formy: ['miesiąc', 'miesiące', 'miesięcy', 'miesiąca'], dni: 30.4375 },
  { re: /^(tydzień|tygodnie|tygodni|tygodnia|tydz|tyg|hebd)/i, formy: ['tydzień', 'tygodnie', 'tygodni', 'tygodnia'], dni: 7 },
  { re: /^(dzień|dni|dnia|dzie|dier|dies|die|d$)/i, formy: ['dzień', 'dni', 'dni', 'dnia'], dni: 1 },
  { re: /^(godzina|godziny|godzin|godz|hor)/i, formy: ['godzina', 'godziny', 'godzin', 'godziny'], dni: 1 / 24 },
];
const WIEK_RE = /(?:(?:(\d+)\s+)?(\d+)\s*\/\s*(\d+)|(\d+(?:\.\d+)?)?\s*([½¼¾⅓⅔⅙⅚⅛])|(\d+(?:\.\d+)?))\s*\.?\s*([a-ząćęłńóśźż-]+)?\.?/gi;

function wiekPolski(zrodlo) {
  const surowy = String(zrodlo ?? '').trim();
  if (!surowy || /^[—–\-?.]+$/.test(surowy)) return zrodlo ?? '';
  if (/\bn\.|\bur\.|urodz|natus|nata\b|\b1[5-9]\d\d\b/i.test(surowy)) return zrodlo;
  const t = surowy.replace(/\(.*?\)/g, ' ').replace(/(\d),(\d)/g, '$1.$2').replace(/\s+/g, ' ').trim();
  const czesci = [];
  for (const m of t.matchAll(WIEK_RE)) {
    if (!m[0].trim()) continue;
    let x;
    if (m[2] !== undefined) { if (!+m[3]) continue; x = (m[1] ? +m[1] : 0) + +m[2] / +m[3]; }
    else if (m[5]) x = (m[4] ? +m[4] : 0) + WIEK_ULAMKI[m[5]];
    else x = +m[6];
    const j = m[7] ? WIEK_JEDNOSTKI.find((jj) => jj.re.test(m[7])) : null;
    if (m[7] && !j) return zrodlo;                      // nieznana jednostka – zostaw, jak jest
    czesci.push({ x, j });
  }
  if (!czesci.length) return zrodlo;
  const jednostki = czesci.map((c) => c.j).filter(Boolean);
  const glowna = jednostki.length ? WIEK_JEDNOSTKI.find((jj) => jednostki.includes(jj)) : null;
  const x = czesci.length === 1 ? czesci[0].x : czesci.reduce((s, c) => s + c.x * (c.j || glowna).dni, 0) / glowna.dni;
  const y = Math.floor(x * 2 + 0.5 + 1e-9) / 2;
  const liczbaTekst = (Math.round(y * 100) / 100).toString().replace('.', ',');
  if (!glowna) return liczbaTekst;
  const f = glowna.formy;
  let forma;
  if (!Number.isInteger(y)) forma = f[3];
  else if (y === 1) forma = f[0];
  else { const j = y % 10, d = y % 100; forma = j >= 2 && j <= 4 && !(d >= 12 && d <= 14) ? f[1] : f[2]; }
  return `${liczbaTekst} ${forma}`;
}
// </wiek-polski>

const IMIONA_MESKIE_NA_A =['kuba', 'barnaba', 'bonawentura', 'kosma', 'jarema', 'zawisza', 'sasza', 'bonaventura'];
const IMIONA_ZENSKIE_BEZ_A = ['agnes', 'elisabeth', 'elizabeth', 'beatrix', 'gertrudis', 'hedvigis', 'ruth', 'rachel', 'dorothea'];

function plecImienia(imie) {
  const s = ((imie || '').trim().split(/\s+/)[0] || '').toLowerCase();
  if (s.length < 3 || /[\[\]]/.test(s)) return null;
  if (IMIONA_MESKIE_NA_A.includes(s)) return 'M';
  if (IMIONA_ZENSKIE_BEZ_A.includes(s)) return 'K';
  if (/^puell|^duae$|^bliźniaczk/.test(s)) return 'K';
  if (/^puer|^bliźniacy$/.test(s)) return 'M';
  if (/^(duo|gemell|bliźnięt)/.test(s)) return null;        // bliźnięta: chłopiec i dziewczynka albo nie wiadomo
  return s.endsWith('a') || s.endsWith('ae') ? 'K' : 'M';
}

function rdzenNazwiska(n) {
  let k = (n || '').toLowerCase().replace(/[^a-ząćęłńóśźż]/g, '');
  for (const [z, na] of [['ska', 'ski'], ['cka', 'cki'], ['dzka', 'dzki'], ['ówna', ''], ['owna', ''], ['owa', '']]) {
    if (k.endsWith(z)) { k = k.slice(0, -z.length) + na; break; }
  }
  return k;
}

// ---------------------------------------------------------------- dane wycinków

function tekst(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'boolean') return v ? 'tak' : 'nie';
  return String(v).trim();
}

function normalizujWspolne(klucz, v) {
  if (!v) return '';
  const s = v.toLowerCase();
  if (klucz === 'katolik') {
    if (['tak', 'true', '1', 'catholica', 'katolik', 'katolicka', 'katolickie'].includes(s)) return 'tak';
    if (['nie', 'false', '0', 'alia', 'aut alia', 'inne', 'inna'].includes(s)) return 'nie';
  }
  if (klucz === 'slubne') {
    if (/^(tak|true|1|legitim|ślubn|slubn|prawe)/.test(s)) return 'tak';
    if (/^(nie|false|0|illegitim|nieślubn|nieslubn|nieprawe)/.test(s)) return 'nie';
  }
  if (klucz === 'plec') {
    if (/^(m|masc|męż|mez|mezczyzna|mężczyzna)/.test(s)) return 'M';
    if (/^(k|f|fem|kob)/.test(s)) return 'K';
  }
  if (klucz === 'typ') {
    if (TYPY[v]) return v;
    const kod = Object.keys(TYPY).find((t) => TYPY[t] === s || t === s.replace(/\s+/g, '_'));
    if (kod) return kod;
  }
  if (klucz === 'miesiac' || klucz === 'miesiac_pogrzebu') {           // "Januarius.", "styczeń", "1" -> "1"
    const n = miesiacZ(v);
    if (n) return String(n);
  }
  return v;
}

const NAZWISKA = ['zm_nazwisko', 'zm_rodowe', 'oj_nazwisko', 'ma_rodowe', 'mz_nazwisko', 'mz_rodowe'];
// pola z kilkoma osobami (imiona po łacinie i po polsku są inne, nazwiska te same): poprawione słowa
// przechodzą na drugą stronę i do tekstu aktu tak jak w „Innych informacjach z aktu”
const POLA_OSOB = ['chrzestni', 'oj_rodzice', 'ma_rodzice', 'akuszerka'];

const NAZWY_MIESIECY = ['styczeń', 'luty', 'marzec', 'kwiecień', 'maj', 'czerwiec', 'lipiec', 'sierpień', 'wrzesień',
  'październik', 'listopad', 'grudzień'];

/* Miesiąc pogrzebu: wpisany w polu (albo w dniu pogrzebu jako „3/2”); gdy go nie ma, ten sam co zgonu,
   a gdy dzień pogrzebu jest wcześniejszy niż dzień zgonu - następny (zgon 31 stycznia, pogrzeb 3 lutego). */
function miesiacPogrzebu(w) {
  const L = w.lacina;
  const [dp, mp] = dzienMiesiac(L.dzien_pogrzebu);
  const jawny = miesiacZ(L.miesiac_pogrzebu) || mp;
  if (jawny) return { n: jawny, opis: 'wpisany' };
  const [dz, mz] = dzienMiesiac(L.dzien_zgonu);
  const m = mz || miesiacZ(L.miesiac);
  if (!m) return null;
  return dz !== null && dp !== null && dp < dz ? { n: (m % 12) + 1, opis: `następny po ${S.po}` } : { n: m, opis: `jak ${S.zdarzenia}` };
}

/* Miesiąc dla aktów, przy których nie jest napisany: z daty "28/2" albo z najbliższego wcześniejszego
   aktu (w księdze nazwa miesiąca stoi tylko nad pierwszym aktem miesiąca). Taki miesiąc jest
   oznaczony jako automatyczny - zmienia się sam, gdy poprawisz miesiąc wcześniej. */
const TYPY_DZIECI = ['dziecko', 'dziecko_nieslubne', 'martwe_urodzenie'];

function formaZenska(n) {          // córka: Rudecki -> Rudecka, Ciepły -> Ciepła; Rzepka, Morek bez zmian
  return /(sk|ck|dzk)i$/.test(n) || /[^aeiouyąęó]y$/.test(n) ? n.slice(0, -1) + 'a' : n;
}

/* Nazwisko dziecka: w akcie zwykle nie stoi przy imieniu, więc bierzemy je od ojca (dziecko nieślubne -
   od matki), u dziewczynki w formie żeńskiej. Zmienia się samo, gdy poprawisz nazwisko rodzica;
   nazwisko wpisane albo poprawione ręcznie zostaje bez zmian. */
function uzupelnijNazwiskaDzieci() {
  for (const w of st.wycinki) {
    w.nazwiskoAuto = false;
    if (w.rodzaj !== 'akt' || !(URODZENIA || TYPY_DZIECI.includes(w.lacina.typ))) continue;
    if (w.edytowane.has('lacina.zm_nazwisko') || w.edytowane.has('polski.zm_nazwisko')) continue;
    const z = (d) => {
      const bezOjca = !d.oj_imie && !d.oj_nazwisko;      // ojca nie ma w akcie (np. martwo urodzone dziecko nieślubne)
      const nieslubne = URODZENIA ? w.lacina.slubne === 'nie' : w.lacina.typ === 'dziecko_nieslubne';
      const n = nieslubne || bezOjca ? (d.ma_rodowe || d.oj_nazwisko) : d.oj_nazwisko;
      return n && w.lacina.plec === 'K' && n === d.oj_nazwisko ? formaZenska(n) : (n || '');
    };
    const lat = z(w.lacina), pol = z(w.polski) || lat;
    if (!lat && !pol) continue;          // brak nazwiska rodzica: zostaje to, co odczytało AI
    w.lacina.zm_nazwisko = lat || pol;
    w.polski.zm_nazwisko = pol;
    w.nazwiskoAuto = true;
  }
}

function uzupelnijMiesiace() {
  let ostatni = null;
  for (const w of st.wycinki) {
    if (w.rodzaj !== 'akt') continue;
    const [, zDaty] = dzienMiesiac(w.lacina.dzien_zgonu);
    if (w.lacina.miesiac && !w.miesiacAuto) {
      ostatni = miesiacZ(w.lacina.miesiac) || ostatni;
      continue;
    }
    const m = zDaty || ostatni;
    w.miesiacAuto = !!m;
    w.lacina.miesiac = w.polski.miesiac = m ? String(m) : '';
    if (zDaty) ostatni = zDaty;
  }
}

function pusteWartosci() {
  return Object.fromEntries(POLA.map(([k]) => [k, '']));
}

/* Uwagi bez pogrzebu (od 13.11.2026), bez stanów zmarłego, męża, ojca i matki (od 24.11.2026) i bez przyczyny zgonu (od 25.11.2026) – także uwagi
   poprawione wcześniej ręcznie; reszta tekstu (np. dopisane informacje z aktu) zostaje. */
function bezPogrzebu(s) {
  return String(s ?? '')
    .replace(/(^|, )pogrzeb: \d{1,2}\.\d{1,2}\.\d{4}(?=, |$)/, '')
    .replace(/(^|, )stan(?: męża| ojca| matki)?: [^,]*(?=, |$)/g, '')
    .replace(/(^|, )przyczyna zgonu: [^,]*(?:, (?![a-ząćęłńóśźż ]+: )[^,]*)*(?=, |$)/g, '')
    .replace(/^, /, '');
}

function normalizuj(wpis, plik, wycinek) {
  wpis = wpis || {};
  const w = {
    wycinek, plik,
    rodzaj: RODZAJE[wpis.rodzaj] ? wpis.rodzaj : 'akt',
    lacina: pusteWartosci(), polski: pusteWartosci(),
    niepewne: Array.isArray(wpis.niepewne) ? wpis.niepewne.map(String) : [],
    uwagi: tekst(wpis.uwagi), sprawdzony: !!wpis.sprawdzony,
    edytowane: new Set(Array.isArray(wpis.edytowane_pola) ? wpis.edytowane_pola : []),  // np. 'polski.zm_imie'
    zmieniono: +wpis.zmieniono || 0,     // czas ostatniej poprawki: przy łączeniu z poprawkami ze strony www wygrywa nowsza
    ai: null,                            // odczyt AI do porównania (zapamietajAi)
    miesiacAuto: false,
    // ręczne poprawki kolumn Geneteki: {wiersze: [{poprawki: {kolumna: {wartosc, auto}}}]} (drugi wiersz = bliźnięta)
    geneteka: wpis.geneteka && Array.isArray(wpis.geneteka.wiersze) ? JSON.parse(JSON.stringify(wpis.geneteka)) : null,
  };
  // od 13.11.2026 pogrzebu nie ma w uwagach Geneteki – znika też z uwag poprawionych wcześniej ręcznie
  for (const r of (w.geneteka && w.geneteka.wiersze) || []) {
    const p = r && r.poprawki && r.poprawki.M;
    if (p) { p.wartosc = bezPogrzebu(p.wartosc); p.auto = bezPogrzebu(p.auto); }
  }
  for (const [k, , rodzaj] of POLA) {
    let l = tekst(wpis.lacina && wpis.lacina[k]);
    let p = tekst(wpis.polski && wpis.polski[k]);
    if (k === 'miesiac' && wpis.miesiac_auto) l = p = '';     // policzy się od nowa (uzupelnijMiesiace)
    if (rodzaj !== 'tekst' && rodzaj !== 'dlugi') {
      l = p = normalizujWspolne(k, l || p);       // jedna wartość dla obu języków (liczby, płeć, rodzaj aktu)
    }
    w.lacina[k] = l;
    w.polski[k] = p;
  }
  return w;
}

function daneDoZapisu() {
  return {
    sciezka: st.sciezka,
    dane: {
      opis: st.opis, ai_zmieniony: st.aiZmieniony, linki_skanow: st.linki,
      wycinki: st.wycinki.map((w) => ({
        wycinek: w.wycinek, plik: w.plik, rodzaj: w.rodzaj, lacina: w.lacina, polski: w.polski,
        niepewne: w.niepewne, uwagi: w.uwagi, sprawdzony: w.sprawdzony, miesiac_auto: w.miesiacAuto,
        edytowany: w.edytowane.size > 0, edytowane_pola: [...w.edytowane], zmieniono: w.zmieniono,
        geneteka: w.geneteka || undefined,
      })),
      link_jednostki: st.link || undefined,
      linki_zmieniono: st.linkiZmieniono || undefined,    // przy łączeniu z poprawkami ze strony wygrywają nowsze linki
    },
  };
}

/* Wpis z Twoimi poprawkami (z) + aktualny odczyt AI (a), pole po polu: pola, które sam wpisałeś
   albo zmieniłeś, zostają; pozostałe zawsze biorą wartość z odczytu AI (np. gdy AI skończyło
   czytać już po tym, jak coś wpisałeś). */
function polacz(z, a, plik, wycinek) {
  // wiek po polsku zawsze zaokrąglony (wiekPolski) – w odczycie AI i w zapisie tak samo,
  // więc sprawdzone akty nie tracą znacznika „sprawdzony”; łacina zostaje dosłowna
  if (!z) {
    const w = normalizuj(a, plik, wycinek);
    w.polski.wiek = wiekPolski(w.polski.wiek);
    if (a) zapamietajAi(w, w);
    return { w, zmiany: 0, zdjete: false };
  }
  const w = normalizuj(z, plik, wycinek);
  w.polski.wiek = wiekPolski(w.polski.wiek);
  if (!a) return { w, zmiany: 0, zdjete: false };
  const ai = normalizuj(a, plik, wycinek);
  ai.polski.wiek = wiekPolski(ai.polski.wiek);
  zapamietajAi(w, ai);
  // zapis z pierwszej wersji programu nie pamiętał, które pola zmieniłeś - wtedy Twoje są wszystkie niepuste
  const staryZapis = !Array.isArray(z.edytowane_pola) && !!z.edytowany;
  const twoje = (id, wartosc) => {
    if (w.edytowane.has(id)) return true;
    if (staryZapis && wartosc !== '') { w.edytowane.add(id); return true; }
    return false;
  };
  let zmiany = 0;
  for (const jezyk of ['lacina', 'polski']) {
    for (const [k] of POLA) {
      if (twoje(`${jezyk}.${k}`, w[jezyk][k]) || w[jezyk][k] === ai[jezyk][k]) continue;
      // nazwisko dziecka liczy program (uzupelnijNazwiskaDzieci) - to nie jest nowy odczyt AI
      if (k === 'zm_nazwisko' && (URODZENIA || TYPY_DZIECI.includes(w.lacina.typ))) continue;
      w[jezyk][k] = ai[jezyk][k];
      zmiany++;
    }
  }
  if (w.rodzaj !== ai.rodzaj && !twoje('rodzaj', staryZapis ? w.rodzaj : '')) { w.rodzaj = ai.rodzaj; zmiany++; }
  if (w.uwagi !== ai.uwagi && !twoje('uwagi', w.uwagi)) { w.uwagi = ai.uwagi; zmiany++; }
  w.niepewne = ai.niepewne;
  const zdjete = zmiany > 0 && w.sprawdzony;
  if (zdjete) w.sprawdzony = false;          // sprawdzałeś inne wartości niż te, które są teraz
  return { w, zmiany, zdjete };
}

/* Odczyt AI zostaje obok poprawek, żeby przy polu poprawionym przez człowieka pokazać ✓ (i co odczytało AI). */
function zapamietajAi(w, ai) {
  w.ai = { lacina: { ...ai.lacina }, polski: { ...ai.polski }, uwagi: ai.uwagi };
}

/* Pole poprawione przez człowieka: ktoś je zmienił i jest teraz inne niż w odczycie AI
   (wpisane i cofnięte do tego samego nie liczy się; bez odczytu AI liczy się każde wpisane). */
function poprawione(w, klucz, jezyk) {
  if (klucz === 'uwagi') return w.edytowane.has('uwagi') && (!w.ai || w.uwagi.trim() !== (w.ai.uwagi || '').trim());
  const j = jezyk === 'oba' ? 'lacina' : jezyk;
  if (!w.edytowane.has(`${j}.${klucz}`)) return false;
  return !w.ai || (w[j][klucz] || '').trim() !== (w.ai[j][klucz] || '').trim();
}

function opisPoprawki(w, klucz, jezyk) {
  if (!w.ai) return 'Poprawione (wpisane ręcznie, bez odczytu AI)';
  const j = jezyk === 'oba' ? 'lacina' : jezyk;
  let ai = klucz === 'uwagi' ? w.ai.uwagi : w.ai[j][klucz];
  if (klucz === 'typ' && TYPY[ai]) ai = TYPY[ai];
  if (klucz === 'miesiac' && ai) ai = `${ai} – ${NAZWY_MIESIECY[+ai - 1] || ai}`;
  return ai ? `Poprawione – AI odczytało: „${ai}”` : 'Poprawione – AI zostawiło to pole puste';
}

function liczPoprawione(w) {
  let n = poprawione(w, 'uwagi') ? 1 : 0;
  for (const [k, , rodzaj] of POLA) {
    if (rodzaj === 'tekst' || rodzaj === 'dlugi') n += poprawione(w, k, 'lacina') + poprawione(w, k, 'polski');
    else n += poprawione(w, k, 'oba');
  }
  return n + poprawkiGeneteki(w);
}

// ---------------------------------------------------------------- format Geneteki: wartości i podświetlenia

function przeliczGeneteke() {
  st.gen = new Map();
  if (!GENETEKA) return;
  const wyniki = K.konwertujRocznik({ wycinki: st.wycinki, linki_skanow: st.linki },
    { rok: st.opis && st.opis.rok, slownik: SLOWNIK, miejscowosc: (st.opis && st.opis.parafia) || 'Przytkowice',
      link: st.link || undefined, skany: st.skany });
  for (const r of wyniki) st.gen.set(r.wycinek, r);
}

/** Wiersze Geneteki aktu: [{nr, wartosci, auto, poprawki, flagi (kolumna -> [powody]), zmiany}] – jak w „genetyka weryfikacja”. */
function wierszeGeneteki(w) {
  const wynik = GENETEKA && w.rodzaj === 'akt' ? st.gen.get(w.wycinek) : null;
  if (!wynik) return [];
  const lista = w.geneteka && w.geneteka.wiersze.length ? w.geneteka.wiersze : [{ poprawki: {} }];
  return lista.map((r, nr) => {
    const poprawki = r.poprawki || {};
    const wartosci = {}, flagi = {}, zmiany = {}, aktywne = {};
    for (const k of K.KOLUMNY) {
      // poprawka równa obecnej wartości automatycznej (np. link wpisany ręcznie, potem ten sam w „Link do skanów”) już nic nie zmienia
      let p = poprawki[k] && poprawki[k].wartosc !== wynik.wiersz[k] ? poprawki[k] : null;
      // nr skanu doszedł później (Tartak, zakresy) – dopisuje się do uwag poprawionych ręcznie, bez pytania
      const nrSkanu = k === 'M' && p && !/nr skanu:/.test(p.wartosc)
        && wynik.wiersz.M.startsWith(`${p.auto}, nr skanu: `) ? wynik.wiersz.M.slice(p.auto.length) : '';
      if (nrSkanu) p = { wartosc: p.wartosc + nrSkanu, auto: wynik.wiersz.M };
      wartosci[k] = p ? p.wartosc : wynik.wiersz[k];
      if (p) aktywne[k] = p;
      const powody = [];
      if (!p) powody.push(...(wynik.flagi[k] || []));
      if (p && p.auto !== wynik.wiersz[k]) {
        zmiany[k] = p;
        powody.push(`odczyt się zmienił: automatycznie było „${p.auto || '(puste)'}”, teraz „${wynik.wiersz[k] || '(puste)'}” – zostaje Twoja poprawka`);
      }
      if (k !== 'A' && K.WYMAGANE.includes(k) && !String(wartosci[k] || '').trim() && !powody.some((x) => /wymagane/.test(x))) {
        powody.push('pole wymagane przez Genetekę jest puste');
      }
      if (k !== 'L' && p && /[\\()'"*;]|\.\.\.|…|\.,|[½¼¾]/.test(p.wartosc)) powody.push('znaki niedozwolone w Genetece (zostaną usunięte przy zapisie)');
      if (powody.length) flagi[k] = powody;
    }
    if (nr > 0) wartosci.A = '';
    return { nr, wartosci, auto: wynik.wiersz, poprawki: aktywne, flagi, zmiany };
  });
}

/** Podświetlenia Geneteki aktu jako teksty „H nazwisko: powód”. */
function flagiGeneteki(w) {
  const wynik = [];
  for (const r of wierszeGeneteki(w)) {
    for (const [k, powody] of Object.entries(r.flagi)) for (const p of powody) wynik.push(`${k} ${K.NAGLOWKI[k]}: ${p}`);
  }
  return wynik;
}

function poprawkiGeneteki(w) {
  if (!w.geneteka) return 0;
  const wiersze = wierszeGeneteki(w);
  return wiersze.reduce((n, r) => n + Object.keys(r.poprawki).length, 0) + Math.max(0, wiersze.length - 1);
}

function maFlagi(w) {
  return st.flagi.has(w.wycinek) || flagiGeneteki(w).length > 0;
}

// ---------------------------------------------------------------- podejrzane miejsca (reguły)

function obliczFlagi() {
  const flagi = new Map();
  const rok = parseInt(st.opis && st.opis.rok, 10) || null;
  let poprzNr = null, poprzData = null, miesiac = null;
  for (const w of st.wycinki) {
    if (w.rodzaj !== 'akt') continue;
    const f = {};
    const dodaj = (pole, t) => { (f[pole] = f[pole] || []).push(t); };
    const L = w.lacina, P = w.polski;

    const nr = liczba(L.nr);
    if (!L.nr) dodaj('nr', 'brak numeru aktu');
    else if (nr === null) dodaj('nr', 'numer aktu to nie liczba');
    else if (poprzNr !== null && nr !== poprzNr + 1) dodaj('nr', `numeracja: po akcie nr ${poprzNr} jest nr ${nr}`);
    if (nr !== null) poprzNr = nr;

    const m = miesiacZ(L.miesiac);
    const poprzMiesiac = miesiac;
    if (m) miesiac = m;
    else if (L.miesiac) dodaj('miesiac', 'nie rozpoznaję miesiąca (wpisz numer 1–12)');
    if (!L.miesiac) dodaj('miesiac', 'brak miesiąca – nie ma go na wycinku ani we wcześniejszych aktach');
    if (m && poprzMiesiac && m < poprzMiesiac) dodaj('miesiac', `miesiąc ${m} po miesiącu ${poprzMiesiac} z wcześniejszego aktu`);
    const [dz, mz] = dzienMiesiac(L.dzien_zgonu);
    if (mz && m && mz !== m && !w.miesiacAuto) dodaj('miesiac', `data ${S.zdarzenia} zapisana jako ${L.dzien_zgonu}, a miesiąc ${m}`);
    const [dp, mpZapis] = dzienMiesiac(L.dzien_pogrzebu);
    const mpPole = miesiacZ(L.miesiac_pogrzebu);
    const mp = mpPole || mpZapis;          // miesiąc pogrzebu (chrztu): z pola albo z zapisu dnia „3/2”
    if (L.dzien_zgonu && dz === null) dodaj('dzien_zgonu', `dzień ${S.zdarzenia} to nie liczba`);
    if (L.dzien_pogrzebu && dp === null) dodaj('dzien_pogrzebu', `dzień ${S.drugiego} to nie liczba`);
    const mies = mz || miesiac;
    let zgon = null;
    if (dz !== null && rok && mies) {
      zgon = data(rok, mies, dz);
      if (!zgon) dodaj('dzien_zgonu', `niemożliwa data ${S.zdarzenia}`);
    }
    if (zgon && dp !== null) {
      let pm, pr;
      if (mp) { pm = mp; pr = rok + (mp < mies ? 1 : 0); }
      else if (dp >= dz) { pm = mies; pr = rok; }
      else { pm = (mies % 12) + 1; pr = rok + (mies === 12 ? 1 : 0); }
      const pogrzeb = data(pr, pm, dp);
      if (!pogrzeb) dodaj('dzien_pogrzebu', `niemożliwa data ${S.drugiego}`);
      else {
        const dni = Math.round((pogrzeb - zgon) / 864e5);
        if (dni > (URODZENIA ? 14 : 5) || dni < 0) {     // chrzest bywa tego samego dnia, pogrzeb nie
          dodaj('dzien_pogrzebu', dni < 0 ? `${S.drugie} przed ${S.przed}` : `${S.drugie} ${dni} dni po ${S.po}`);
          if (mpPole) dodaj('miesiac_pogrzebu', `miesiąc ${S.drugiego} ${mpPole}: ${S.drugie} ${dni} dni po ${S.po}`);
        } else if (dni === 0 && !URODZENIA) dodaj('dzien_pogrzebu', 'pogrzeb tego samego dnia co zgon');
      }
    }
    if (zgon && poprzData && zgon < poprzData) dodaj('dzien_zgonu', `data ${S.zdarzenia} wcześniejsza niż w poprzednim akcie`);
    if (zgon) poprzData = zgon;

    const typ = L.typ, plec = L.plec;
    if (['zona', 'wdowa'].includes(typ) && plec && plec !== 'K') {
      dodaj('plec', `rodzaj aktu „${TYPY[typ]}”, a płeć ${plec}`);
      dodaj('typ', `rodzaj aktu „${TYPY[typ]}”, a płeć ${plec}`);
    }
    if (['maz', 'wdowiec'].includes(typ) && plec && plec !== 'M') {
      dodaj('plec', `rodzaj aktu „${TYPY[typ]}”, a płeć ${plec}`);
      dodaj('typ', `rodzaj aktu „${TYPY[typ]}”, a płeć ${plec}`);
    }
    const imie = P.zm_imie || L.zm_imie;
    const pi = plecImienia(imie);
    if (pi && plec && pi !== plec) {
      const t = `imię „${imie}” wygląda na ${pi === 'K' ? 'żeńskie' : 'męskie'}, a płeć ${plec}`;
      dodaj('plec', t);
      dodaj('zm_imie', t);
    }
    if (L.katolik === 'nie') dodaj('katolik', 'zaznaczone wyznanie inne niż katolickie – sprawdź na skanie');

    const lata = wiekWLatach(L.wiek) ?? wiekWLatach(P.wiek);
    if (typ === 'martwe_urodzenie' && L.wiek && !/^[—–-]+$/.test(L.wiek)) dodaj('wiek', 'martwe urodzenie, a podany wiek');
    if (['zona', 'wdowa', 'maz', 'wdowiec'].includes(typ) && lata !== null && lata < 14) {
      dodaj('wiek', `osoba w małżeństwie, a wiek ${L.wiek || P.wiek}`);
    }
    if (typ === 'dziecko' && L.zm_nazwisko && L.oj_nazwisko && rdzenNazwiska(L.zm_nazwisko) !== rdzenNazwiska(L.oj_nazwisko)) {
      dodaj('zm_nazwisko', `nazwisko dziecka inne niż ojca (${L.oj_nazwisko})`);
    }
    if (['zona', 'wdowa'].includes(typ) && L.zm_nazwisko && L.mz_nazwisko &&
        rdzenNazwiska(L.zm_nazwisko) !== rdzenNazwiska(L.mz_nazwisko)) {
      dodaj('zm_nazwisko', `nazwisko zmarłej inne niż męża (${L.mz_nazwisko})`);
    }
    for (const k of ['zm_nazwisko', 'zm_rodowe', 'oj_nazwisko', 'ma_rodowe', 'mz_nazwisko', 'mz_rodowe']) {
      if (L[k] && P[k] && L[k] !== P[k]) {
        dodaj(k, GENETEKA ? `w tłumaczeniu (format Geneteki) jest „${P[k]}” – nazwisk się nie tłumaczy; popraw łacinę albo kolumnę Geneteki`
          : 'nazwisko po polsku różni się od łacińskiego (nazwisk się nie tłumaczy)');
      }
    }
    for (const [k] of POLA) {
      if (/\[\s*łac\./i.test(P[k])) dodaj(k, 'AI zostawiło część po łacinie ([łac. …]) – do przetłumaczenia');
    }
    if (Object.keys(f).length) flagi.set(w.wycinek, f);
  }
  st.flagi = flagi;
}

// ---------------------------------------------------------------- drzewo folderów

const IKONY = { nie_pociete: '–', pociete: '○', ai_w_toku: '◐', ai: '●', w_trakcie: '●', gotowe: '✓', gotowe_zmiany: '✓' };
const OPISY_STANU = {
  nie_pociete: 'skany nie są jeszcze pocięte (Tartak)',
  pociete: 'pocięte, czeka na odczyt AI',
  ai_w_toku: 'AI jeszcze czyta (albo przerwało)',
  ai: 'odczyt AI gotowy – do weryfikacji',
  w_trakcie: 'w trakcie weryfikacji',
  gotowe: 'gotowe – pobrano xlsx',
  gotowe_zmiany: 'pobrano xlsx, ale potem były zmiany – pobierz jeszcze raz',
};

function rozwiniete() {
  try { return JSON.parse(localStorage.getItem('rozwiniete') || 'null'); } catch (e) { return null; }
}
function zapiszRozwiniete(zbior) {
  try { localStorage.setItem('rozwiniete', JSON.stringify([...zbior])); } catch (e) { /* bez znaczenia */ }
}

async function odswiezDrzewo() {
  try {
    st.drzewo = await API.get('/api/drzewo');
  } catch (e) {
    $('#drzewo').replaceChildren(el('p', { class: 'komunikat blad', text: e.message }));
    return;
  }
  rysujDrzewo();
}

function rysujDrzewo() {
  const zapamietane = rozwiniete();
  const otwarte = new Set(zapamietane || []);
  const domyslnie = zapamietane === null;
  const lista = el('ul');
  const wezel = (w, poziom) => {
    const li = el('li');
    const maDzieci = w.dzieci.length > 0;
    const otwarty = domyslnie ? poziom < 2 : otwarte.has(w.sciezka);
    const info = w.info || {};
    let dopisek = '';
    if (info.wycinkow) dopisek = info.odczytanych !== undefined && w.stan === 'ai_w_toku' ? `${info.odczytanych}/${info.wycinkow}` : `${info.wycinkow}`;
    const wiersz = el('div', {
      class: 'wezel' + (w.stan ? ' rocznik' : ' folder') + (st.opis && w.sciezka === st.opis.rocznik ? ' aktywny' : ''),
      title: w.stan ? `${w.nazwa}: ${OPISY_STANU[w.stan]}` : w.sciezka,
    },
    el('span', { class: 'strzalka', text: maDzieci ? (otwarty ? '▾' : '▸') : '' }),
    el('span', { class: 'nazwa', text: w.nazwa }),
    w.stan ? el('span', { class: 'dopisek', text: dopisek }) : null,
    w.stan ? el('span', { class: `ikona stan-${w.stan}`, text: IKONY[w.stan] + (w.stan === 'gotowe_zmiany' ? '*' : '') }) : null);
    wiersz.dataset.sciezka = w.sciezka;
    li.append(wiersz);
    let podlista = null;
    if (maDzieci) {
      podlista = el('ul', { class: otwarty ? '' : 'ukryta' });
      for (const d of w.dzieci) podlista.append(wezel(d, poziom + 1));
      li.append(podlista);
    }
    const przelacz = () => {
      if (!podlista) return;
      podlista.classList.toggle('ukryta');
      const teraz = !podlista.classList.contains('ukryta');
      wiersz.querySelector('.strzalka').textContent = teraz ? '▾' : '▸';
      const zbior = new Set(rozwiniete() || []);
      if (teraz) zbior.add(w.sciezka); else zbior.delete(w.sciezka);
      zapiszRozwiniete(zbior);
    };
    wiersz.querySelector('.strzalka').addEventListener('click', (e) => { e.stopPropagation(); przelacz(); });
    wiersz.addEventListener('click', () => {
      if (!w.stan) return przelacz();
      if (w.stan === 'nie_pociete') return toast(`${w.nazwa}: najpierw potnij te skany w Tartaku.`);
      otworzFolder(w.sciezka);
    });
    return li;
  };
  for (const d of st.drzewo.dzieci) lista.append(wezel(d, 0));
  if (!st.drzewo.dzieci.length) lista.append(el('li', { class: 'pusto', text: `Brak roczników w ${st.drzewo.sciezka}` }));
  $('#drzewo').replaceChildren(lista);
  if (domyslnie) {
    const zbior = new Set();
    const zbierz = (w, poziom) => { if (poziom < 2 && w.dzieci.length) zbior.add(w.sciezka); w.dzieci.forEach((d) => zbierz(d, poziom + 1)); };
    st.drzewo.dzieci.forEach((d) => zbierz(d, 0));
    zapiszRozwiniete(zbior);
  }
}

function ustawStanWDrzewie(stan) {
  if (!st.opis) return;
  const wiersz = $$('.wezel').find((w) => w.dataset.sciezka === st.opis.rocznik);
  const ikona = wiersz && wiersz.querySelector('.ikona');
  if (!ikona) return;
  ikona.className = `ikona stan-${stan}`;
  ikona.textContent = IKONY[stan] + (stan === 'gotowe_zmiany' ? '*' : '');
  wiersz.title = `${wiersz.querySelector('.nazwa').textContent}: ${OPISY_STANU[stan]}`;
}

// ---------------------------------------------------------------- otwieranie folderu

async function otworzFolder(sciezka) {
  if (st.brudne) await zapiszTeraz();
  let d;
  try {
    d = await API.get('/api/folder?sciezka=' + enc(sciezka));
  } catch (e) {
    return toast('Nie udało się otworzyć folderu: ' + e.message);
  }
  const tenSam = st.sciezka === d.sciezka;      // "Odśwież" tego samego rocznika: strona zostaje w miejscu
  st.sciezka = d.sciezka;
  st.opis = d.opis;
  ustawKsiege(d.opis && d.opis.ksiega);
  st.aiZmieniony = d.ai_zmieniony;
  st.linki = ((d.zweryfikowany && d.zweryfikowany.linki_skanow) || []).filter((l) => l && l.url);
  try { localStorage.setItem('ostatni', d.sciezka); } catch (e) { /* bez znaczenia */ }

  const ai = new Map(((d.ai && d.ai.wycinki) || []).map((w) => [String(w.wycinek), w]));
  const zw = new Map(((d.zweryfikowany && d.zweryfikowany.wycinki) || []).map((w) => [String(w.wycinek), w]));
  let uzupelnione = 0, zdjete = 0;
  st.wycinki = d.wycinki.map((o) => {
    const z = zw.get(o.wycinek);
    const wynik = polacz(z, ai.get(o.wycinek), o.plik, o.wycinek);
    if (z && wynik.zmiany) uzupelnione++;
    if (wynik.zdjete) zdjete++;
    wynik.w.zmieniony = o.zmieniony;
    return wynik.w;
  });
  st.mapa = new Map(st.wycinki.map((w) => [w.wycinek, w]));
  st.link = (d.zweryfikowany && d.zweryfikowany.link_jednostki) || '';
  st.linkiZmieniono = (d.zweryfikowany && d.zweryfikowany.linki_zmieniono) || 0;
  st.skany = d.skany || null;          // wycinek -> zdjęcie strony (Tartak, pocięte/skany.json)
  /* Rocznik pobrany już starym sposobem (dawny xlsx Łacina/Polski, sprawdzany potem w programie „Weryfikacja
     Geneteka”) zostaje po staremu; format Geneteki od razu tutaj mają roczniki jeszcze niepobrane. */
  const zapis = d.zweryfikowany || {};
  st.stary = !URODZENIA && !!(d.geneteka_poprawki || (zapis.pobrano && !zapis.xlsx_geneteki));
  if (st.stary) GENETEKA = false;

  const kom = [];
  const nazwa = [d.opis.parafia, d.opis.ksiega, d.opis.rok].filter(Boolean).join(' · ') || d.sciezka;
  if (d.uwaga) kom.push(['uwaga', d.uwaga]);
  if (d.ai_blad) {
    kom.push(['blad', `Plik odczyt_ai.json jest uszkodzony (${d.ai_blad}). Poproś AI w Antigravity o poprawienie pliku (musi być poprawnym JSON), potem kliknij Odśwież.`, true]);
  } else if (!d.ai && d.wycinki.length) {
    kom.push(['', WWW ? 'Odczytu AI tego rocznika jeszcze nie ma. Zajrzyj później (Odśwież) – pola możesz też wypełniać ręcznie.'
      : 'Ten folder nie ma jeszcze odczytu AI. Kliknij „Prompt dla AI”, skopiuj go i wklej w Antigravity. Gdy AI skończy, kliknij „Odśwież”. Pola możesz też wypełniać ręcznie.', true]);
  } else if (d.ai && !d.ai.ukonczono) {
    kom.push(['uwaga', `AI jeszcze nie skończyło (albo przerwało): odczytane ${ai.size} z ${d.wycinki.length} wycinków. Kliknij „Odśwież”, żeby wczytać nowe.`, true]);
  }
  const bezObrazu = [...ai.keys()].filter((k) => !st.mapa.has(k));
  if (bezObrazu.length) kom.push(['uwaga', `W odczycie AI są wpisy bez obrazu: ${bezObrazu.join(', ')}. Numeracja wycinków mogła się zmienić (ponowne cięcie w Tartaku?).`]);
  if (uzupelnione) kom.push(['', `Wczytano nowy odczyt AI do ${uzupelnione} wycinków. Pola, które wpisałeś albo poprawiłeś sam, zostały bez zmian.`]);
  if (zdjete) kom.push(['uwaga', `W ${zdjete} wycinkach oznaczonych jako sprawdzone AI zmieniło wartości – oznaczenie „sprawdzone” zdjęte, przejrzyj je jeszcze raz.`]);
  if (st.stary) {
    kom.push(['', 'Ten rocznik był już pobrany starym sposobem, więc jest pokazany po staremu (łacina i polski obok siebie, '
      + '„Pobierz xlsx” w dawnym formacie). Format Geneteki sprawdzasz dla niego w programie „Weryfikacja Geneteka”.']);
  }
  pokazKomunikaty(kom);

  $('#naglowek-folder').textContent = nazwa;
  $('#naglowek-folder').title = d.sciezka;
  $('#pusto').classList.add('ukryta');
  $('#narzedzia').classList.toggle('ukryta', !st.wycinki.length);
  $('#b-pobierz').disabled = !st.wycinki.length;
  $('#b-pobierz').textContent = GENETEKA ? 'Pobierz xlsx genetyka' : 'Pobierz xlsx';
  $('#b-pobierz').title = GENETEKA ? 'Zapisz plik xlsx dla Geneteki w folderze rocznika' : 'Pobierz plik xlsx z arkuszami Łacina i Polski';
  $('#b-scal').disabled = !GENETEKA;
  // na stronie www dawny xlsx (Łacina/Polski) się nie robi – robi go program na komputerze
  $('#b-pobierz').classList.toggle('ukryta', !!WWW && !GENETEKA);
  $('#b-linki').textContent = GENETEKA ? 'Link do skanów' : 'Linki do skanów';
  $('#b-linki').title = GENETEKA ? 'Link do jednostki w Skanotece (kolumna „plik” Geneteki)'
    : 'Zakresy wycinków i linki do zdjęć stron księgi (kolumna „Skan” w xlsx)';
  uzupelnijMiesiace();
  uzupelnijNazwiskaDzieci();
  obliczFlagi();
  przeliczGeneteke();
  rysujKarty();
  $$('.wezel').forEach((w) => w.classList.toggle('aktywny', w.dataset.sciezka === d.opis.rocznik));
  if (!tenSam) {
    // powrót do rozpoczętego rocznika: od razu pierwszy niesprawdzony wycinek
    const pierwszy = st.wycinki.some((w) => w.sprawdzony) && st.wycinki.find((w) => !w.sprawdzony);
    const karta = pierwszy && document.getElementById('w-' + pierwszy.wycinek);
    if (karta) requestAnimationFrame(() => karta.scrollIntoView({ block: 'start' }));
    else window.scrollTo(0, 0);
  }
  if (uzupelnione) zaplanujZapis();
}

function pokazKomunikaty(lista) {
  const box = $('#komunikaty');
  box.replaceChildren(...lista.map(([rodzaj, tekstK, zPrzyciskami]) => el('div', { class: `komunikat ${rodzaj}` },
    el('span', { text: tekstK }),
    zPrzyciskami ? el('div', { class: 'przyciski' },
      el('button', { class: 'tylko-lokalnie', text: 'Prompt dla AI', onclick: pokazPrompt }),
      el('button', { text: 'Odśwież', onclick: odswiez })) : null)));
}

// ---------------------------------------------------------------- karty wycinków

let obserwator = null;

function rysujKarty() {
  if (obserwator) obserwator.disconnect();
  obserwator = new IntersectionObserver((wpisy) => {
    for (const wpis of wpisy) {
      if (!wpis.isIntersecting) continue;
      const karta = wpis.target;
      if (!karta.dataset.gotowa) rysujPola(karta);
      obserwator.unobserve(karta);
    }
  }, { rootMargin: '1500px 0px' });
  const karty = st.wycinki.map(karta);
  $('#karty').replaceChildren(...karty);
  karty.forEach((k) => obserwator.observe(k));
  odswiezLiczniki();
  zastosujFiltr();
}

function opisKarty(w) {
  const L = w.lacina, P = w.polski;
  if (w.rodzaj !== 'akt') return RODZAJE[w.rodzaj];
  const imie = P.zm_imie || L.zm_imie;
  const nazwisko = L.zm_nazwisko || L.zm_rodowe || L.oj_nazwisko || '';
  return [L.nr ? `akt nr ${L.nr}` : 'akt', [imie, nazwisko].filter(Boolean).join(' ')].filter(Boolean).join(' · ');
}

function karta(w) {
  const s = el('section', { class: 'karta' + (w.sprawdzony ? ' gotowa' : ''), id: 'w-' + w.wycinek });
  s.dataset.wycinek = w.wycinek;
  const rodzaj = el('select', { class: 'rodzaj', title: 'Rodzaj wycinka' },
    Object.entries(RODZAJE).map(([k, n]) => el('option', { value: k, text: n })));
  rodzaj.value = w.rodzaj;
  rodzaj.addEventListener('change', () => {
    w.rodzaj = rodzaj.value;
    w.edytowane.add('rodzaj');
    rysujPola(s);
    poZmianie(s, w);
  });
  // ten sam przycisk u góry i na dole karty (żeby po sprawdzeniu pól nie przewijać do góry);
  // dolny (i Enter) po oznaczeniu przechodzi do następnego wycinka
  const sprawdzone = [0, 1].map(() => el('button', { class: 'b-sprawdzone', type: 'button' }));
  const pokazSprawdzone = () => sprawdzone.forEach((b, i) => {
    b.classList.toggle('wcisniety', w.sprawdzony);
    if (i === 0) b.textContent = w.sprawdzony ? '✓ Sprawdzone' : 'Oznacz jako sprawdzone';
    else {
      b.replaceChildren(w.sprawdzony ? 'Sprawdzony ✓ (kliknij, żeby cofnąć)' : 'Sprawdzone ✓', el('kbd', { text: 'Enter' }));
      b.title = 'Oznacz jako sprawdzony i przejdź do następnego wycinka (Enter)';
    }
  });
  pokazSprawdzone();
  s.ustawSprawdzony = (wartosc, dalej) => {
    if (w.sprawdzony !== wartosc) {
      w.sprawdzony = wartosc;
      w.zmieniono = Date.now();
      zaplanujZapis();
    }
    pokazSprawdzone();
    s.classList.toggle('gotowa', w.sprawdzony);
    // sprawdzony wycinek zwija pola; cofnięcie oznaczenia znów je pokazuje
    szczegoly.open = w.sprawdzony ? false : st.otwarte;
    zaznaczFlagiKarty(s, w);
    odswiezLiczniki();
    zastosujFiltr(true);
    if (w.sprawdzony && dalej) {
      const nast = sasiedniaKarta(s, 1);
      if (nast) pokazKarte(nast);
      else {
        s.scrollIntoView({ block: 'start' });
        const brak = st.wycinki.filter((x) => !x.sprawdzony).length;
        toast(brak ? `To ostatni wycinek. Niesprawdzonych zostało: ${brak}.` : 'Wszystkie wycinki sprawdzone – możesz pobrać xlsx.');
      }
    }
  };
  sprawdzone[0].addEventListener('click', () => s.ustawSprawdzony(!w.sprawdzony, false));
  sprawdzone[1].addEventListener('click', () => s.ustawSprawdzony(!w.sprawdzony, true));
  const poprzedni = el('button', { type: 'button', text: '← Poprzedni skan', title: 'Poprzedni wycinek',
    onclick: () => { const k = sasiedniaKarta(s, -1); if (k) pokazKarte(k); } });
  const nastepny = el('button', { type: 'button', text: 'Następny skan →', title: 'Następny wycinek (bez oznaczania)',
    onclick: () => { const k = sasiedniaKarta(s, 1); if (k) pokazKarte(k); } });
  const img = el('img', { loading: 'lazy', alt: `wycinek ${w.wycinek}` });
  if (WWW) WWW.obraz(img, st.sciezka, w.plik);       // zaszyfrowany na GitHubie - odszyfrowuje www.js
  else img.src = `/api/obraz?sciezka=${enc(st.sciezka)}&plik=${enc(w.plik)}&v=${w.zmieniony}`;
  const obraz = el('div', { class: 'obraz', title: 'Kliknij, żeby powiększyć' }, img);
  img.addEventListener('click', () => obraz.classList.toggle('duzy'));
  const szczegoly = el('details', { open: st.otwarte && !w.sprawdzony },
    el('summary', { text: GENETEKA ? 'Łacina i format Geneteki (możesz poprawić każde pole)' : 'Wszystkie pola (możesz poprawić każde)' }),
    el('div', { class: 'cialo' }));
  szczegoly.addEventListener('toggle', () => { if (szczegoly.open) $$('textarea', szczegoly).forEach(dopasujWysokosc); });
  s.append(
    el('h2', {},
      el('span', { text: `Wycinek ${w.wycinek}` }),
      el('span', { class: 'opis', text: opisKarty(w) }),
      el('a', { class: 'link-skanu' + (linkSkanu(w.wycinek) ? '' : ' ukryta'), href: linkSkanu(w.wycinek) || '#',
        target: '_blank', rel: 'noopener', text: 'skan ↗', title: 'Zdjęcie strony księgi (Linki do skanów)' }),
      el('span', { class: 'znaczek ukryta' }),
      el('span', { class: 'znaczek ok poprawki ukryta', title: 'Pola poprawione względem odczytu AI – przy każdym z nich jest ✓' }),
      el('span', { class: 'prawo' }, rodzaj, sprawdzone[0])),
    obraz, szczegoly,
    el('div', { class: 'stopka' }, poprzedni, sprawdzone[1], nastepny));
  zaznaczFlagiKarty(s, w);
  return s;
}

/** Następna (kierunek 1) albo poprzednia (-1) widoczna karta. */
function sasiedniaKarta(karta, kierunek) {
  let k = karta;
  do k = kierunek > 0 ? k.nextElementSibling : k.previousElementSibling;
  while (k && k.classList.contains('ukryta'));
  return k;
}

/* Przejście do wycinka: karta od góry, a gdy cała się nie mieści – przewinięta tak, żeby było widać cały
   „Format Geneteki” (wycinek skanu zostaje przyklejony u góry, łacina chowa się pod nim). */
function pokazKarte(k) {
  if (!k.dataset.gotowa) rysujPola(k);
  k.scrollIntoView({ block: 'start' });
  requestAnimationFrame(() => {
    const g = $('.sekcja-geneteki', k);
    if (!g || !g.offsetParent) return;
    const r = g.getBoundingClientRect();
    const nadGeneteka = (parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--wys-naglowka')) || 53)
      + $('.obraz', k).offsetHeight + 8;
    const o = Math.min(r.bottom - window.innerHeight + 12, r.top - nadGeneteka);
    if (o > 0) window.scrollBy(0, o);
  });
}

/** Karta, w której jesteś: z kursorem w polu albo pierwsza widoczna pod nagłówkiem. */
function biezacaKarta() {
  const a = document.activeElement && document.activeElement.closest && document.activeElement.closest('.karta');
  if (a) return a;
  const gora = $('header').offsetHeight + 40;
  return $$('.karta').find((k) => !k.classList.contains('ukryta') && k.getBoundingClientRect().bottom > gora) || null;
}

function pole(w, klucz, jezyk, dlugi, podpowiedz) {
  const wartosc = klucz === 'uwagi' ? w.uwagi : w[jezyk === 'oba' ? 'lacina' : jezyk][klucz];
  const e = dlugi ? el('textarea', { rows: 1 }) : el('input', { type: 'text' });
  e.value = wartosc || '';
  if (podpowiedz) e.placeholder = podpowiedz;
  e.dataset.pole = klucz;
  e.dataset.jezyk = jezyk;
  e.spellcheck = false;
  return e;
}

function wybor(w, klucz, opcje) {
  const wartosc = w.lacina[klucz] || '';
  const s = el('select', {}, opcje.map(([v, n]) => el('option', { value: v, text: n })));
  if (wartosc && !opcje.some(([v]) => v === wartosc)) s.append(el('option', { value: wartosc, text: wartosc }));
  s.value = wartosc;
  s.dataset.pole = klucz;
  s.dataset.jezyk = 'oba';
  return s;
}

function rysujPola(karta) {
  const w = st.mapa.get(karta.dataset.wycinek);
  const cialo = $('.cialo', karta);
  if (GENETEKA) {
    // łacina (oryginał) w polach obok siebie, a pod nią ten sam akt w formacie Geneteki
    cialo.replaceChildren(sekcjaLaciny(w), ...(w.rodzaj === 'akt' ? [sekcjaGeneteki(w)] : []));
    karta.dataset.gotowa = '1';
    zaznaczFlagiKarty(karta, w);
    requestAnimationFrame(() => $$('textarea', cialo).forEach(dopasujWysokosc));
    return;
  }
  const wiersze = [];
  const wiersz = (etykieta, klucz, ...komorki) => {
    const tr = el('tr', {}, el('th', { class: 'etykieta', text: etykieta }), ...komorki);
    tr.dataset.pole = klucz;
    return tr;
  };
  // komórka z polem i miejscem na ✓ po prawej (pole poprawione przez człowieka)
  const kom = (e, colspan) => el('td', colspan ? { colspan } : {},
    el('div', { class: 'z-ptaszkiem' }, e, el('span', { class: 'ptaszek', text: '✓', 'aria-hidden': 'true' })));
  const oba = (e) => kom(e, '2');
  if (w.rodzaj === 'akt') {
    for (const [k, etykieta, rodzaj] of POLA) {
      if (rodzaj === 'wspolne') wiersze.push(wiersz(etykieta, k, oba(pole(w, k, 'oba', false))));
      else if (rodzaj === 'tak_nie') wiersze.push(wiersz(etykieta, k, oba(wybor(w, k, [['', '—'], ['tak', 'tak'], ['nie', 'nie']]))));
      else if (rodzaj === 'plec') wiersze.push(wiersz(etykieta, k, oba(wybor(w, k, [['', '—'], ['M', 'M – mężczyzna'], ['K', 'K – kobieta']]))));
      else if (rodzaj === 'typ') wiersze.push(wiersz(etykieta, k, oba(wybor(w, k, [['', '—'], ...Object.entries(TYPY)]))));
      else if (rodzaj === 'miesiac') {
        wiersze.push(wiersz(etykieta, k, oba(wybor(w, k, [['', '—'],
          ...NAZWY_MIESIECY.map((n, i) => [String(i + 1), `${i + 1} – ${n}`])]))));
      } else {
        const dlugi = rodzaj === 'dlugi';
        wiersze.push(wiersz(etykieta, k, kom(pole(w, k, 'lacina', dlugi)), kom(pole(w, k, 'polski', dlugi))));
      }
    }
  } else {
    wiersze.push(wiersz('Tekst', 'tekst', kom(pole(w, 'tekst', 'lacina', true)), kom(pole(w, 'tekst', 'polski', true))));
    if (w.rodzaj === 'poswiadczenie') {
      wiersze.push(wiersz('Ksiądz', 'ksiadz', kom(pole(w, 'ksiadz', 'lacina', false)), kom(pole(w, 'ksiadz', 'polski', false))));
    }
  }
  wiersze.push(wiersz('Uwagi', 'uwagi', oba(pole(w, 'uwagi', 'oba', true, 'problem z wycinkiem, Twoje notatki…'))));
  const tabela = el('table', { class: 'pola' },
    el('thead', {}, el('tr', {}, el('th', { class: 'etykieta', text: 'Pole' }), el('th', { text: 'Łacina' }), el('th', { text: 'Polski' }))),
    el('tbody', {}, wiersze));
  cialo.replaceChildren(tabela);
  karta.dataset.gotowa = '1';
  zaznaczFlagiKarty(karta, w);
  requestAnimationFrame(() => $$('textarea', cialo).forEach(dopasujWysokosc));
}

function dopasujWysokosc(t) {
  if (!t.offsetParent) return;
  t.style.height = 'auto';
  t.style.height = (t.scrollHeight + 2) + 'px';
}

/** Szerokość pola łaciny według długości tekstu (pola stoją obok siebie jak w „genetyka weryfikacja”). */
function dopasujSzerokosc(e) {
  if (e.tagName === 'INPUT') e.size = Math.max(5, Math.min(42, (e.value || e.placeholder || '').length + 2));
}

// ---------------------------------------------------------------- karta: łacina

function sekcjaLaciny(w) {
  const krotkie = el('div', { class: 'siatka' }), dlugie = el('div', { class: 'dlugie' });
  const f = st.flagi.get(w.wycinek) || {};
  let pustych = 0;
  const dodaj = (k, etykieta, kontrolka, dlugi) => {
    const pusty = !String(kontrolka.value || '').trim();
    const schowany = pusty && !ZAWSZE_WIDOCZNE.has(k) && !f[k] && !w.niepewne.includes(k);
    if (schowany) pustych++;
    dopasujSzerokosc(kontrolka);
    (dlugi ? dlugie : krotkie).append(el('div', { class: 'pole' + (schowany ? ' pusty' : '') },
      el('span', { class: 'etykieta', text: etykieta }),
      el('div', { class: 'z-ptaszkiem' }, kontrolka, el('span', { class: 'ptaszek', text: '✓', 'aria-hidden': 'true' }))));
  };
  const pola = w.rodzaj === 'akt' ? POLA
    : [['tekst', 'Tekst', 'dlugi'], ...(w.rodzaj === 'poswiadczenie' ? [['ksiadz', 'Ksiądz', 'tekst']] : [])];
  for (const [k, etykieta, rodzaj] of pola) {
    if (rodzaj === 'wspolne') dodaj(k, etykieta, pole(w, k, 'oba', false));
    else if (rodzaj === 'tak_nie') dodaj(k, etykieta, wybor(w, k, [['', '—'], ['tak', 'tak'], ['nie', 'nie']]));
    else if (rodzaj === 'plec') dodaj(k, etykieta, wybor(w, k, [['', '—'], ['M', 'M – mężczyzna'], ['K', 'K – kobieta']]));
    else if (rodzaj === 'typ') dodaj(k, etykieta, wybor(w, k, [['', '—'], ...Object.entries(TYPY)]));
    else if (rodzaj === 'miesiac') {
      dodaj(k, etykieta, wybor(w, k, [['', '—'], ...NAZWY_MIESIECY.map((n, i) => [String(i + 1), `${i + 1} – ${n}`])]));
    } else dodaj(k, etykieta, pole(w, k, 'lacina', rodzaj === 'dlugi'), rodzaj === 'dlugi');
  }
  dodaj('uwagi', 'Uwagi (Twoje notatki do wycinka)', pole(w, 'uwagi', 'oba', true, 'problem z wycinkiem, Twoje notatki…'), true);
  const przycisk = pustych ? el('button', { type: 'button', class: 'b-puste', onclick: przelaczPuste }) : null;
  if (przycisk) {
    przycisk.dataset.ile = pustych;
    przycisk.textContent = st.pokazPuste ? 'Ukryj puste pola' : `Pokaż puste pola (${pustych})`;
  }
  return el('div', { class: 'sekcja-karty sekcja-laciny' },
    el('h3', {}, el('span', { text: w.rodzaj === 'akt' ? 'Łacina' : 'Tekst wycinka' }),
      el('span', { class: 'opis', text: '(odczyt AI – oryginał; popraw, jeśli źle odczytany)' }),
      przycisk ? el('span', { class: 'prawo' }, przycisk) : null),
    krotkie, dlugie);
}

function przelaczPuste() {
  st.pokazPuste = !st.pokazPuste;
  try { localStorage.setItem('pokazPuste', st.pokazPuste ? '1' : '0'); } catch (e) { /* bez znaczenia */ }
  $('#karty').classList.toggle('pokaz-puste', st.pokazPuste);
  for (const b of $$('.b-puste')) b.textContent = st.pokazPuste ? 'Ukryj puste pola' : `Pokaż puste pola (${b.dataset.ile})`;
}

// ---------------------------------------------------------------- karta: format Geneteki

function sekcjaGeneteki(w) {
  const sekcja = el('div', { class: 'sekcja-karty sekcja-geneteki' },
    el('h3', {}, el('span', { text: 'Format Geneteki' }),
      el('span', { class: 'prawo' }, el('button', { type: 'button', text: 'Duplikuj wiersz (bliźnięta)',
        title: 'Dodaj drugi wiersz z tym samym numerem aktu (np. bliźnięta w jednym akcie)', onclick: () => duplikujWiersz(w) }))),
    el('p', { class: 'podpowiedz' }, el('span', { class: 'probka do-sprawdzenia' }), ' do sprawdzenia (najedź na pole) ',
      el('span', { class: 'ptaszek widoczny', text: '✓' }), ' zmieniono ręcznie (kliknij ✓, żeby przywrócić wartość automatyczną)'),
    el('div', { class: 'wiersze-geneteki' }));
  rysujWierszeGeneteki(sekcja, w);
  return sekcja;
}

function rysujWierszeGeneteki(sekcja, w) {
  const wiersze = wierszeGeneteki(w);
  const wynik = st.gen.get(w.wycinek);
  $('.wiersze-geneteki', sekcja).replaceChildren(...wiersze.map((r) => rysujWierszGeneteki(w, wynik, r, wiersze.length)));
  requestAnimationFrame(() => $$('textarea', sekcja).forEach(dopasujWysokosc));
}

function rysujWierszGeneteki(w, wynik, r, ile) {
  const blok = el('div', { class: 'wiersz-geneteki', 'data-wiersz': String(r.nr) });
  if (ile > 1) {
    blok.append(el('h4', {}, `Wiersz ${r.nr + 1}`, r.nr > 0
      ? el('button', { type: 'button', text: 'Usuń ten wiersz', onclick: () => usunWiersz(w, r.nr) })
      : el('span', { class: 'opis', text: '(główny)' })));
  }
  const siatka = el('div', { class: 'siatka-geneteki' });
  for (const k of K.KOLUMNY) {
    const etykieta = el('span', { class: 'etykieta' }, el('span', { class: 'litera', text: k }), ' ',
      K.WYMAGANE.includes(k) ? el('b', { text: K.NAGLOWKI[k] }) : K.NAGLOWKI[k]);
    let e;
    if (k === 'A') {
      e = el('input', { type: 'text', value: r.nr > 0 ? '(kolejny)' : 'automatycznie', readOnly: true, class: 'tylko-odczyt',
        tabIndex: -1, title: 'L.p. jest numerowane automatycznie przy zapisie xlsx' });
    } else if (SZEROKIE_GEN.has(k)) {
      e = el('textarea', { rows: 2, spellcheck: false });
      e.value = r.wartosci[k] ?? '';
    } else {
      e = el('input', { type: 'text', spellcheck: false });
      e.value = r.wartosci[k] ?? '';
    }
    e.dataset.kol = k;
    e.setAttribute('aria-label', `${k} ${K.NAGLOWKI[k]}`);
    const ptaszek = el('span', { class: 'ptaszek', text: '✓', role: 'button', tabIndex: -1 });
    ptaszek.addEventListener('click', () => przywrocGeneteke(w, r.nr, k));
    if (k !== 'A') {
      e.addEventListener('input', () => {
        if (e.tagName === 'TEXTAREA') dopasujWysokosc(e);
        zmienPoleGeneteki(w, r.nr, k, e.value);
      });
    }
    siatka.append(el('div', { class: `pole${SZEROKIE_GEN.has(k) ? ' szerokie' : SREDNIE_GEN.has(k) ? ' sredni' : ''}` },
      etykieta, el('span', { class: 'z-ptaszkiem' }, e, ptaszek)));
  }
  blok.append(siatka, el('ul', { class: 'powody ukryta' }));
  if (r.nr === 0 && wynik.podpowiedzi.length) {
    blok.append(el('div', { class: 'dopiski' }, wynik.podpowiedzi.map((t) => el('div', { class: 'dopisek' },
      el('b', { text: 'Inne informacje z aktu:' }), el('span', { text: t }),
      el('button', { type: 'button', text: 'Dopisz do uwag', title: 'Dopisuje tekst na końcu kolumny „uwagi” – potem zostaw tylko to, co jest w oryginale',
        onclick: () => dopiszDoUwag(w, r.nr, t) })))));
  }
  odswiezStanWierszaGeneteki(blok, w, r);
  return blok;
}

/** Kolory, ptaszki i lista powodów „do sprawdzenia” bez przerysowywania pól (kursor zostaje w polu). */
function odswiezStanWierszaGeneteki(blok, w, r) {
  for (const k of K.KOLUMNY) {
    const e = blok.querySelector(`[data-kol="${k}"]`);
    if (!e) continue;
    const opak = e.closest('.pole');
    const flagi = r.flagi[k];
    e.classList.toggle('do-sprawdzenia', !!flagi);
    opak.classList.toggle('ma-flage', !!flagi);
    e.title = flagi ? flagi.join('\n') : '';
    const p = r.poprawki[k];
    const ptaszek = opak.querySelector('.ptaszek');
    ptaszek.classList.toggle('widoczny', !!p);
    ptaszek.title = p ? `Zmieniono ręcznie. Automatycznie: „${r.auto[k] || '(puste)'}”. Kliknij, żeby przywrócić.` : '';
  }
  const pozycje = [];
  for (const k of K.KOLUMNY) {
    for (const powod of r.flagi[k] || []) {
      const li = el('li', {}, el('b', { text: `${k} (${K.NAGLOWKI[k]}): ` }), powod);
      if (r.zmiany[k] && powod.startsWith('odczyt się zmienił')) {
        li.append(el('button', { type: 'button', text: 'Zostaw moją', onclick: () => potwierdzPoprawke(w, r.nr, k) }),
          el('button', { type: 'button', text: 'Weź nową', onclick: () => przywrocGeneteke(w, r.nr, k) }));
      }
      pozycje.push(li);
    }
  }
  const ul = blok.querySelector('.powody');
  ul.replaceChildren(...pozycje);
  ul.classList.toggle('ukryta', !pozycje.length);
}

/** Po przeliczeniu: nowe wartości automatyczne (poza poprawionymi ręcznie i polem z kursorem) i podświetlenia. */
function odswiezGeneteke(karta, w) {
  const sekcja = $('.sekcja-geneteki', karta);
  if (!sekcja) return;
  const wiersze = wierszeGeneteki(w);
  const bloki = $$('.wiersz-geneteki', sekcja);
  if (bloki.length !== wiersze.length) return rysujWierszeGeneteki(sekcja, w);
  wiersze.forEach((r, i) => {
    for (const k of K.KOLUMNY) {
      if (k === 'A' || r.poprawki[k]) continue;
      const e = bloki[i].querySelector(`[data-kol="${k}"]`);
      const v = r.wartosci[k] ?? '';
      if (e && e !== document.activeElement && e.value !== v) {
        e.value = v;
        if (e.tagName === 'TEXTAREA') dopasujWysokosc(e);
      }
    }
    odswiezStanWierszaGeneteki(bloki[i], w, r);
  });
}

function kartaWycinka(w) {
  return document.getElementById('w-' + w.wycinek);
}

function poZmianieGeneteki(w, przerysuj = false) {
  // jeden wiersz bez poprawek = nic do zapamiętania
  if (w.geneteka && w.geneteka.wiersze.length === 1 && !Object.keys(w.geneteka.wiersze[0].poprawki || {}).length) w.geneteka = null;
  w.zmieniono = Date.now();
  zaplanujZapis();
  const karta = kartaWycinka(w);
  if (!karta) return;
  if (przerysuj) rysujWierszeGeneteki($('.sekcja-geneteki', karta), w);
  else odswiezGeneteke(karta, w);
  zaznaczFlagiKarty(karta, w);
  odswiezLiczniki();
}

function wierszGeneteki(w, nr) {
  if (!w.geneteka || !w.geneteka.wiersze.length) w.geneteka = { wiersze: [{ poprawki: {} }] };
  const r = w.geneteka.wiersze[nr] || (w.geneteka.wiersze[nr] = { poprawki: {} });
  r.poprawki = r.poprawki || {};
  return r;
}

function zmienPoleGeneteki(w, nr, k, wartosc) {
  const auto = st.gen.get(w.wycinek).wiersz[k];
  const r = wierszGeneteki(w, nr);
  if (wartosc === auto) delete r.poprawki[k];
  else r.poprawki[k] = { wartosc, auto };
  poZmianieGeneteki(w);
}

function przywrocGeneteke(w, nr, k) {
  const r = w.geneteka && w.geneteka.wiersze[nr];
  if (!r || !r.poprawki || !r.poprawki[k]) return;
  delete r.poprawki[k];
  poZmianieGeneteki(w, true);
}

function potwierdzPoprawke(w, nr, k) {
  const r = w.geneteka && w.geneteka.wiersze[nr];
  if (r && r.poprawki && r.poprawki[k]) r.poprawki[k].auto = st.gen.get(w.wycinek).wiersz[k];
  poZmianieGeneteki(w);
}

function dopiszDoUwag(w, nr, tekstD) {
  const karta = kartaWycinka(w);
  const e = karta && karta.querySelector(`.wiersz-geneteki[data-wiersz="${nr}"] textarea[data-kol="M"]`);
  if (!e) return;
  const dopisek = K.czysc(K.ulamkiNaDziesietne(tekstD)).replace(/\.\s*$/, '');
  e.value = e.value ? `${e.value.replace(/[\s,]+$/, '')}, ${dopisek}` : dopisek;
  e.dispatchEvent(new Event('input'));
  e.focus();
}

function duplikujWiersz(w) {
  const pierwszy = wierszGeneteki(w, 0);
  // drugi wiersz zaczyna jako kopia pierwszego (z jego poprawkami)
  w.geneteka.wiersze.push({ poprawki: JSON.parse(JSON.stringify(pierwszy.poprawki)) });
  poZmianieGeneteki(w, true);
  toast('Dodano drugi wiersz z tym samym numerem aktu – popraw imię i uwagi.');
  const karta = kartaWycinka(w);
  const nowy = karta && karta.querySelector(`.wiersz-geneteki[data-wiersz="${w.geneteka.wiersze.length - 1}"]`);
  if (nowy) nowy.scrollIntoView({ behavior: 'smooth', block: 'center' });
}

function usunWiersz(w, nr) {
  okno('Usunąć wiersz?', [el('p', { text: `Wiersz ${nr + 1} tego aktu (format Geneteki) zostanie usunięty.` })], [
    el('button', { class: 'glowny', text: 'Usuń', onclick: () => {
      zamknijOkno();
      if (!w.geneteka) return;
      w.geneteka.wiersze.splice(nr, 1);
      poZmianieGeneteki(w, true);
    } }),
  ]);
}

/** Przeliczenie Geneteki po zmianie odczytu (łaciny, dat, linków) we wszystkich narysowanych kartach. */
function odswiezWszystkieKarty() {
  przeliczGeneteke();
  for (const k of $$('.karta')) {
    const w = st.mapa.get(k.dataset.wycinek);
    if (k.dataset.gotowa) odswiezGeneteke(k, w);
    zaznaczFlagiKarty(k, w);
  }
  odswiezLiczniki();
}

function zaznaczFlagiKarty(karta, w) {
  const f = st.flagi.get(w.wycinek) || {};
  const fg = flagiGeneteki(w);
  const kolumnyGen = new Set(fg.map((t) => t.split(' ')[0]));
  const ile = Object.keys(f).length + kolumnyGen.size;
  karta.classList.toggle('ma-flagi', ile > 0 && !w.sprawdzony);
  const znaczek = $('.znaczek', karta);
  if (znaczek) {
    znaczek.classList.toggle('ukryta', ile === 0);
    znaczek.textContent = `do sprawdzenia: ${ile}`;
    znaczek.title = [...Object.values(f).flat(), ...fg.map((t) => `Geneteka ${t}`)].filter((v, i, a) => a.indexOf(v) === i).join('\n');
  }
  const ilePoprawek = liczPoprawione(w);
  const poprawki = $('.poprawki', karta);
  if (poprawki) {
    poprawki.classList.toggle('ukryta', ilePoprawek === 0);
    poprawki.textContent = `✓ poprawione: ${ilePoprawek}`;
  }
  for (const e of $$('[data-pole]', karta)) {
    if (e.tagName === 'TR') {
      e.classList.toggle('wiersz-podejrzany', !!f[e.dataset.pole]);
      continue;
    }
    const komunikaty = f[e.dataset.pole];
    e.classList.toggle('podejrzane', !!komunikaty);
    const opak = e.closest('.pole');
    if (opak) opak.classList.toggle('ma-flage', !!komunikaty);
    e.classList.toggle('niepewne', w.niepewne.includes(e.dataset.pole));
    const tytuly = [];
    if (komunikaty) tytuly.push(...komunikaty);
    if (w.niepewne.includes(e.dataset.pole)) tytuly.push('AI zaznaczyło, że nie jest pewne tego odczytu');
    if (e.dataset.pole === 'miesiac') {
      // miesiąc uzupełniany automatycznie zmienia się, gdy poprawisz miesiąc we wcześniejszym akcie
      if (e.value !== w.lacina.miesiac && document.activeElement !== e) e.value = w.lacina.miesiac;
      e.classList.toggle('auto', w.miesiacAuto);
      if (w.miesiacAuto) tytuly.push('uzupełnione automatycznie z wcześniejszego aktu (albo z daty „dzień/miesiąc”) – zmień, jeśli trzeba');
    }
    if (e.dataset.pole === 'zm_nazwisko' && e.tagName === 'INPUT') {
      // nazwisko dziecka liczone z nazwiska rodzica odświeża się, gdy poprawisz rodzica
      const v = w[e.dataset.jezyk].zm_nazwisko;
      if (w.nazwiskoAuto && e.value !== v && document.activeElement !== e) e.value = v;
      e.classList.toggle('auto', w.nazwiskoAuto);
      if (w.nazwiskoAuto) tytuly.push('uzupełnione automatycznie z nazwiska ojca (u dziecka nieślubnego – matki); zmienia się samo, gdy poprawisz rodzica, albo wpisz ręcznie');
    }
    if (e.dataset.pole === 'miesiac_pogrzebu' && e.options) {
      // puste pole = miesiąc liczony automatycznie; napis pokazuje, jaki wyjdzie (wpisz tylko, gdy inny)
      const ef = miesiacPogrzebu(w);
      e.options[0].textContent = ef && ef.opis !== 'wpisany'
        ? `— ${ef.opis}: ${NAZWY_MIESIECY[ef.n - 1]}` : `— tak jak ${S.zdarzenia}`;
      tytuly.push(URODZENIA
        ? 'Zostaw puste, gdy chrzest był w tym samym miesiącu co urodzenie. Wybierz miesiąc, gdy był w innym (np. urodzenie 31 stycznia, chrzest 2 lutego).'
        : 'Zostaw puste, gdy pogrzeb był w tym samym miesiącu co zgon. Wybierz miesiąc, gdy był w innym (np. zgon 31 stycznia, pogrzeb 3 lutego).');
    }
    e.title = tytuly.join('\n');
    const ptaszek = e.nextElementSibling;
    if (ptaszek && ptaszek.classList.contains('ptaszek')) {
      const zmienione = poprawione(w, e.dataset.pole, e.dataset.jezyk);
      ptaszek.classList.toggle('widoczny', zmienione);
      ptaszek.title = zmienione ? opisPoprawki(w, e.dataset.pole, e.dataset.jezyk) : '';
    }
  }
}

// ---------------------------------------------------------------- edycja i zapis

function poZmianie(karta, w) {
  w.zmieniono = Date.now();
  $('.opis', karta).textContent = opisKarty(w);
  zaplanujZapis();
  clearTimeout(st.flagiTimer);
  st.flagiTimer = setTimeout(() => {
    uzupelnijMiesiace();
    uzupelnijNazwiskaDzieci();
    obliczFlagi();
    odswiezWszystkieKarty();          // przelicza też format Geneteki (zmiana łaciny, dat, numeru domu…)
    zastosujFiltr(true);
  }, 300);
}

/* Poprawka słowa w polu (np. nazwisko rodowe matki) poprawia to samo słowo w „Tekście aktu (całym)”.
   Zamiana liczy się zawsze od stanu z chwili wejścia w pole, więc pisanie litera po literze
   (także wyczyszczenie pola i wpisanie od nowa) nie psuje tekstu. Podmieniane są tylko całe słowa. */
function zapamietajEdycje(e) {
  const t = e.target;
  const karta = t.closest && t.closest('.karta');
  if (!karta || !t.dataset.pole || !['lacina', 'polski'].includes(t.dataset.jezyk) ||
      (RODZAJE_POL[t.dataset.pole] !== 'tekst' && t.dataset.pole !== 'dodatkowe')) {
    st.edycja = null;
    return;
  }
  const w = st.mapa.get(karta.dataset.wycinek);
  const k = t.dataset.pole;
  st.edycja = {
    el: t,
    klucz: k,
    stare: { lacina: t.dataset.jezyk === 'lacina' ? w.lacina[k] : null, polski: w.polski[k] },
    pola: { lacina: w.lacina[k], polski: w.polski[k] },
    tekst: { lacina: w.lacina.tekst, polski: w.polski.tekst },
    dodatkowe: { lacina: w.lacina.dodatkowe, polski: w.polski.dodatkowe },
  };
}

/* Słowa zmienione między starą a nową wartością pola: wspólny początek i koniec zostają, zmieniona
   środkowa część -> [[stare, nowe]]. Samo dopisanie albo usunięcie słów nie ma czego zamieniać. */
function zmienioneSlowa(stare, nowe) {
  const slowa = (s) => (s || '').match(/[\p{L}\p{N}'’-]+/gu) || [];
  const s = slowa(stare), n = slowa(nowe);
  let p = 0;
  while (p < s.length && p < n.length && s[p] === n[p]) p++;
  let k = 0;
  while (k < s.length - p && k < n.length - p && s[s.length - 1 - k] === n[n.length - 1 - k]) k++;
  const ss = s.slice(p, s.length - k), nn = n.slice(p, n.length - k);
  return ss.length && nn.length ? [[ss.join(' '), nn.join(' ')]] : [];
}

/* "Inne informacje z aktu" (i pola osób, np. chrzestni): poprawione słowa (np. nazwisko, miejscowość, przezwisko)
   zmieniają się też w tekście aktu i po drugiej stronie (łacina/polski) - tam, gdzie stoi to samo słowo,
   także odmienione. */
function poprawZInnychInformacji(w, karta, edycja, jezyk, nowe) {
  const drugi = jezyk === 'lacina' ? 'polski' : 'lacina';
  const zmiany = zmienioneSlowa(edycja.pola[jezyk], nowe);
  const cele = [[jezyk, 'tekst', edycja.tekst[jezyk]], [drugi, edycja.klucz, edycja.pola[drugi]],
    [drugi, 'tekst', edycja.tekst[drugi]]];
  if (edycja.klucz !== 'dodatkowe') {
    cele.push([jezyk, 'dodatkowe', edycja.dodatkowe[jezyk]], [drugi, 'dodatkowe', edycja.dodatkowe[drugi]]);
  }
  for (const [j, pole, baza] of cele) {
    if (!baza) continue;
    const wynik = zmiany.reduce((t, [a, b]) => zamienSlowo(t, a, b), baza);   // zawsze od stanu z chwili wejścia w pole
    ustawPoprawione(w, karta, j, pole, baza, wynik);
  }
}

function ustawPoprawione(w, karta, j, pole, baza, wynik) {
  if (wynik === w[j][pole]) return;
  w[j][pole] = wynik;
  if (wynik !== baza) w.edytowane.add(`${j}.${pole}`);
  const e = karta.querySelector(`[data-pole="${pole}"][data-jezyk="${j}"]`);
  if (e) {
    e.value = wynik;
    if (e.tagName === 'TEXTAREA') dopasujWysokosc(e);
  }
}

const wRegex = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
const SAMOGLOSKA_NA_KONCU = /[aeiouyąęó]$/i;

/* Słowo z pola także w odmienionej formie: w tekście po polsku jest "syn Jana", "po Kazimierzu Skrzyńskim",
   a w polu "Jan", "Skrzyński". Rdzeń = słowo bez końcowej samogłoski, do niego dowolna krótka końcówka;
   w poprawionym słowie zostaje ta sama końcówka ("Mojawy" -> "Morawy" po poprawce "Mojawa" -> "Morawa"). */
function zamienOdmiane(tekstZ, stare, nowe) {
  const rdzen = (s) => (SAMOGLOSKA_NA_KONCU.test(s) && s.length > 3 ? s.slice(0, -1) : s);
  const rs = rdzen(stare);
  const rn = SAMOGLOSKA_NA_KONCU.test(stare) && stare.length > 3 ? rdzen(nowe) : nowe;
  if (rs.length < 3 || rs === rn) {           // krótkie słowo albo poprawiona sama końcówka: tylko ten sam zapis
    return tekstZ.replace(new RegExp(`(?<![\\p{L}\\p{N}])${wRegex(stare)}(?![\\p{L}\\p{N}])`, 'gu'), () => nowe);
  }
  // to samo słowo co w polu -> nowe słowo w całości; forma odmieniona -> nowy rdzeń + jej końcówka
  let wynik = tekstZ.replace(new RegExp(`(?<![\\p{L}\\p{N}])${wRegex(rs)}(\\p{Ll}{0,4})(?![\\p{L}\\p{N}])`, 'gu'),
    (slowo, koncowka) => (slowo === stare ? nowe : rn + koncowka));
  // ruchome "e": Marek -> Marku, Paweł -> Pawła, Gnojek -> Gnojka
  const bezE = (s) => (/e[kcł]$/.test(s) && s.length > 3 ? s.slice(0, -2) + s.slice(-1) : null);
  if (bezE(stare)) {
    wynik = wynik.replace(new RegExp(`(?<![\\p{L}\\p{N}])${wRegex(bezE(stare))}(\\p{Ll}{1,4})(?![\\p{L}\\p{N}])`, 'gu'),
      (_, koncowka) => (bezE(nowe) || nowe) + koncowka);
  }
  return wynik;
}

function zamienSlowo(tekstZ, stare, nowe) {
  stare = (stare || '').trim();
  nowe = (nowe || '').trim();
  if (!tekstZ || !stare || !nowe) return tekstZ;     // puste pole (np. w trakcie pisania od nowa) nie psuje tekstu
  const s = stare.split(/\s+/), n = nowe.split(/\s+/);
  if (s.length !== n.length) {                        // inna liczba słów: tylko dokładnie ten sam zapis
    return tekstZ.replace(new RegExp(`(?<![\\p{L}\\p{N}])${wRegex(stare)}(?![\\p{L}\\p{N}])`, 'gu'), () => nowe);
  }
  return s.reduce((t, slowo, i) => (slowo === n[i] ? t : zamienOdmiane(t, slowo, n[i])), tekstZ);
}

/* Poprawka pola zmienia to samo słowo w „Tekście aktu (całym)”, a poprawka nazwiska (zmarłego, rodziców,
   małżonka) także w „Innych informacjach z aktu”. Liczone od stanu z chwili wejścia w pole, więc powrót
   do starej wartości przywraca też tekst. */
function poprawTekstAktu(w, karta, edycja, nowe, jezyki) {
  const pola = NAZWISKA.includes(edycja.klucz) ? ['tekst', 'dodatkowe'] : ['tekst'];
  for (const j of jezyki) {
    const stare = edycja.stare[j];
    if (!stare) continue;
    for (const pole of pola) {
      const baza = edycja[pole][j];
      if (baza) ustawPoprawione(w, karta, j, pole, baza, zamienSlowo(baza, stare, nowe));
    }
  }
}

function obsluzZmiane(e) {
  const t = e.target;
  if (!t.dataset || !t.dataset.pole || t.classList.contains('rodzaj')) return;
  const karta = t.closest('.karta');
  if (!karta) return;
  const w = st.mapa.get(karta.dataset.wycinek);
  const klucz = t.dataset.pole, jezyk = t.dataset.jezyk;
  if (klucz === 'uwagi') {
    w.uwagi = t.value;
    w.edytowane.add('uwagi');
  } else if (jezyk === 'oba') {
    w.lacina[klucz] = t.value;
    w.polski[klucz] = t.value;
    w.edytowane.add(`lacina.${klucz}`).add(`polski.${klucz}`);
  } else {
    const poprzednia = w[jezyk][klucz];
    // wpisany wiek po polsku zaokrągla się po wyjściu z pola („1½ roku” -> „1,5 roku”)
    if (e.type === 'change' && jezyk === 'polski' && klucz === 'wiek') t.value = wiekPolski(t.value);
    w[jezyk][klucz] = t.value;
    w.edytowane.add(`${jezyk}.${klucz}`);
    const jezyki = [jezyk];
    // nazwisk się nie tłumaczy: poprawka po łacinie idzie też na polską stronę,
    // chyba że polską wersję ustawiłeś wcześniej ręcznie na coś innego
    if (jezyk === 'lacina' && NAZWISKA.includes(klucz) &&
        (!w.edytowane.has(`polski.${klucz}`) || w.polski[klucz] === poprzednia)) {
      w.polski[klucz] = t.value;
      w.edytowane.add(`polski.${klucz}`);
      const polskie = karta.querySelector(`[data-pole="${klucz}"][data-jezyk="polski"]`);
      if (polskie) polskie.value = t.value;
      jezyki.push('polski');
    }
    if (st.edycja && st.edycja.el === t) {
      if (klucz === 'dodatkowe' || POLA_OSOB.includes(klucz)) poprawZInnychInformacji(w, karta, st.edycja, jezyk, t.value);
      else poprawTekstAktu(w, karta, st.edycja, t.value, jezyki);
    }
  }
  if (klucz === 'miesiac') w.miesiacAuto = false;
  if (t.tagName === 'TEXTAREA') dopasujWysokosc(t);
  else if (t.closest('.siatka')) dopasujSzerokosc(t);
  poZmianie(karta, w);
}

/* Strzałka w dół/górę w polu przechodzi do pola w wierszu niżej/wyżej, w tej samej kolumnie (łacina albo
   polski). W polu wielowierszowym dopiero z końca/początku tekstu. Na listach wyboru strzałka też przechodzi
   dalej, a nie zmienia wartości (zmiana: myszką, Alt+↓ albo wpisaniem pierwszej litery). */
function strzalki(e) {
  if (!['ArrowDown', 'ArrowUp'].includes(e.key) || e.altKey || e.ctrlKey || e.shiftKey || e.metaKey) return;
  const t = e.target;
  if (!t.dataset || !t.dataset.pole || !t.closest('table.pola')) return;
  const wDol = e.key === 'ArrowDown';
  if (t.tagName === 'TEXTAREA' && (wDol ? t.selectionEnd < t.value.length : t.selectionStart > 0)) return;
  if (t.dataset.jezyk !== 'oba') st.kolumna = t.dataset.jezyk;
  const kolumna = t.dataset.jezyk === 'oba' ? st.kolumna : t.dataset.jezyk;
  for (let r = t.closest('tr'); (r = wDol ? r.nextElementSibling : r.previousElementSibling);) {
    const cel = r.querySelector(`[data-pole][data-jezyk="${kolumna}"]`) || r.querySelector('[data-pole]');
    if (!cel) continue;
    e.preventDefault();
    cel.focus();
    if (cel.tagName !== 'SELECT') cel.setSelectionRange(cel.value.length, cel.value.length);
    return;
  }
  if (t.tagName === 'SELECT') e.preventDefault();      // ostatnie pole karty: lista zostaje bez zmian
}

function zaplanujZapis() {
  st.brudne = true;
  ustawStanZapisu('Niezapisane zmiany…');
  clearTimeout(st.zapisTimer);
  st.zapisTimer = setTimeout(zapiszTeraz, 800);
}

async function zapiszTeraz() {
  clearTimeout(st.zapisTimer);
  if (!st.brudne || !st.sciezka) return;
  st.brudne = false;
  try {
    const r = await API.post('/api/zapisz', daneDoZapisu());
    const d = await r.json();
    ustawStanZapisu(d.komunikat || 'Zapisano ' + d.zapisano.slice(11), !!d.ostrzezenie);
    ustawStanWDrzewie(d.pobrano ? (d.zmiany_po_pobraniu === false ? 'gotowe' : 'gotowe_zmiany') : 'w_trakcie');
    if (d.nowsze) toast(`Pomocnik poprawił na stronie www ${d.nowsze} z tych wycinków – kliknij „Odśwież”, żeby zobaczyć jego wersję.`);
  } catch (e) {
    st.brudne = true;
    ustawStanZapisu('Błąd zapisu! ' + e.message, true);
  }
}

function ustawStanZapisu(tekstS, blad = false) {
  const e = $('#stan-zapisu');
  e.textContent = tekstS;
  e.classList.toggle('blad', blad);
}

function odswiezLiczniki() {
  const n = st.wycinki.length;
  const akty = st.wycinki.filter((w) => w.rodzaj === 'akt').length;
  const sprawdzone = st.wycinki.filter((w) => w.sprawdzony).length;
  const zFlagami = st.wycinki.filter((w) => !w.sprawdzony && maFlagi(w)).length;
  $('#licznik').replaceChildren(
    'Wycinków: ', el('b', { text: String(n) }), ` (aktów ${akty}) · sprawdzone: `, el('b', { text: `${sprawdzone}/${n}` }),
    ' · z podświetleniami: ', el('b', { text: String(zFlagami) }));
  $('#pasek').style.width = n ? (100 * sprawdzone / n) + '%' : '0';
}

function zastosujFiltr(tylkoKlasy = false) {
  for (const k of $$('.karta')) {
    const w = st.mapa.get(k.dataset.wycinek);
    let widoczna = true;
    if (st.filtr === 'niesprawdzone') widoczna = !w.sprawdzony;
    if (st.filtr === 'flagi') widoczna = maFlagi(w) || w.niepewne.length > 0;
    if (st.filtr === 'poprawione') widoczna = liczPoprawione(w) > 0;
    // karta, którą właśnie edytujesz, nie znika spod kursora
    if (tylkoKlasy && !widoczna && k.contains(document.activeElement)) continue;
    k.classList.toggle('ukryta', !widoczna);
  }
}

// ---------------------------------------------------------------- pobieranie xlsx

async function pobierz() {
  if (!st.sciezka) return;
  if (GENETEKA) return pobierzGeneteke();
  clearTimeout(st.zapisTimer);
  const przycisk = $('#b-pobierz');
  przycisk.disabled = true;
  try {
    const r = await API.post('/api/xlsx', daneDoZapisu());
    const blob = await r.blob();
    const nazwa = decodeURIComponent(r.headers.get('X-Nazwa-Pliku') || 'weryfikacja.xlsx');
    const a = el('a', { href: URL.createObjectURL(blob), download: nazwa });
    document.body.append(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 20000);
    st.brudne = false;
    ustawStanZapisu('Zapisano i pobrano');
    ustawStanWDrzewie('gotowe');
    const niesprawdzone = st.wycinki.filter((w) => !w.sprawdzony).length;
    toast(`Pobrano ${nazwa}` + (niesprawdzone ? ` (nieoznaczonych jako sprawdzone: ${niesprawdzone})` : ''));
    pokazPropozycjeSlowniczka();
  } catch (e) {
    toast('Nie udało się zrobić pliku xlsx: ' + e.message);
  } finally {
    przycisk.disabled = false;
  }
}

/** Okno z pytaniem: true po kliknięciu przycisku „tak”, false po „Zamknij”. */
function zapytaj(tytul, cialo, tak) {
  return new Promise((rozwiaz) => {
    okno(tytul, cialo, [el('button', { class: 'glowny', text: tak, onclick: () => { zamknijOkno(); rozwiaz(true); } })]);
    $('#okno-przyciski').lastElementChild.addEventListener('click', () => rozwiaz(false));
  });
}

function naBase64(bajty) {
  let s = '';
  for (let i = 0; i < bajty.length; i += 0x8000) s += String.fromCharCode.apply(null, bajty.subarray(i, i + 0x8000));
  return btoa(s);
}

function nazwaXlsxGeneteki(parafia, lata) {
  return `${NAZWY_PARAFII[parafia] || `${parafia || 'parafia'}-Z`}-${lata}.xlsx`;
}

/** Wiersze dla Geneteki całego rocznika, w kolejności wycinków (bliźnięta: kilka wierszy jednego aktu). */
function wierszeRocznika() {
  const wiersze = [];
  for (const wycinek of st.gen.keys()) for (const r of wierszeGeneteki(st.mapa.get(wycinek))) wiersze.push(r.wartosci);
  return wiersze;
}

/* „Pobierz xlsx genetyka”: plik dla Geneteki (szablon indeksy.xlsx, arkusz Zgony) zapisany w folderze rocznika,
   np. Przytkowice/Zgony/1890/Przytkowice-Z-przenajswietszej-trojcy-1890.xlsx. */
async function pobierzGeneteke() {
  clearTimeout(st.zapisTimer);
  przeliczGeneteke();
  const akty = [...st.gen.keys()].map((k) => st.mapa.get(k));
  const niesprawdzone = akty.filter((w) => !w.sprawdzony).map((w) => w.lacina.nr || w.wycinek);
  const puste = [];
  for (const w of akty) {
    for (const r of wierszeGeneteki(w)) {
      const brak = K.WYMAGANE.filter((k) => !String(r.wartosci[k] || '').trim());
      if (brak.length) puste.push(`akt ${r.wartosci.C || '?'}: ${brak.map((k) => K.NAGLOWKI[k]).join(', ')}`);
    }
  }
  const ostrzezenia = [];
  if (niesprawdzone.length) {
    ostrzezenia.push(`Niesprawdzone akty (${niesprawdzone.length}): ${niesprawdzone.slice(0, 30).join(', ')}${niesprawdzone.length > 30 ? '…' : ''}`);
  }
  if (puste.length) ostrzezenia.push(`Puste pola wymagane: ${puste.slice(0, 10).join('; ')}${puste.length > 10 ? '…' : ''}`);
  if (akty.length && !akty.some((w) => wierszeGeneteki(w).some((r) => r.wartosci.L))) {
    ostrzezenia.push('Brak linku do jednostki w Skanotece (kolumna „plik”) – wklej go w „Link do skanów”.');
  }
  if (!akty.length) return toast('W tym roczniku nie ma aktów do zapisania.');
  const nazwa = nazwaXlsxGeneteki(st.opis.parafia, st.opis.rok);
  if (ostrzezenia.length && !await zapytaj('Zapisać xlsx mimo uwag?', [el('ul', {}, ostrzezenia.map((o) => el('li', { text: o })))],
    'Zapisz mimo to')) return;
  const przycisk = $('#b-pobierz');
  przycisk.disabled = true;
  try {
    const plik = await XL.zbudujXlsx(window.JSZip, SZABLON, wierszeRocznika());
    const d = await (await API.post('/api/xlsx_geneteki', { ...daneDoZapisu(), nazwa, xlsx: naBase64(plik) })).json();
    st.brudne = false;
    ustawStanZapisu(WWW ? 'Zapisano, xlsx pobrany' : 'Zapisano i zapisano xlsx');
    ustawStanWDrzewie('gotowe');
    toast(d.komunikat || `Zapisano ${nazwa} w folderze ${d.folder}`);
    if (!WWW) pokazPropozycjeSlowniczka();         // słowniczek (prompt dla AI) jest na komputerze
  } catch (e) {
    toast('Nie udało się zapisać pliku xlsx: ' + e.message);
  } finally {
    przycisk.disabled = false;
  }
}

/* „Scal lata”: zapisane xlsx Geneteki kilku roczników w jeden plik (rocznik pod rocznikiem, ciągła L.p.)
   w folderze księgi, np. Przytkowice/Zgony/Przytkowice-Z-przenajswietszej-trojcy-1890-1896.xlsx. */
async function scalLata() {
  if (!st.sciezka || !GENETEKA) return toast('Otwórz najpierw rocznik zgonów.');
  await zapiszTeraz();
  let d;
  try {
    d = await API.get('/api/geneteka_lata?sciezka=' + enc(st.sciezka));
  } catch (e) {
    return toast('Nie udało się odczytać roczników: ' + e.message);
  }
  const pozycje = d.lata.map((r) => {
    const nazwa = nazwaXlsxGeneteki(d.parafia, r.rok);
    const jest = r.pliki.some((p) => p.normalize('NFC') === nazwa.normalize('NFC'));
    const info = [];
    if (!jest) info.push('brak xlsx Geneteki');
    else {
      info.push(`sprawdzone ${r.sprawdzone}/${r.akty}`);
      if (r.zmiany_po_pobraniu) info.push('zmiany po zapisie xlsx – zapisz ponownie');
    }
    return { r, nazwa, jest, pelny: r.akty > 0 && r.sprawdzone >= r.akty && !r.zmiany_po_pobraniu, info: info.join(' · ') };
  });
  const dostepne = pozycje.filter((p) => p.jest);
  if (dostepne.length < 2) {
    return okno('Scal lata', [el('p', { text: 'Do scalenia potrzebne są co najmniej dwa roczniki z zapisanym xlsx Geneteki '
      + `(przycisk „Pobierz xlsx genetyka”). Teraz: ${dostepne.length}.` })], []);
  }
  const lista = el('ul', { class: 'lata' }, pozycje.map((p) => el('li', {},
    el('input', { type: 'checkbox', id: `rok-${p.r.rok}`, checked: p.jest, disabled: !p.jest, 'data-rok': String(p.r.rok) }),
    el('label', { for: `rok-${p.r.rok}`, text: String(p.r.rok) }),
    el('span', { class: `info${!p.jest || !p.pelny ? ' uwaga' : ''}`, text: p.info }))));
  if (!await zapytaj('Scal lata – jeden plik dla Geneteki', [
    el('p', { text: 'Zaznacz roczniki. W pliku wynikowym będą jeden pod drugim, bez powtarzania nagłówków i bez pustych wierszy, z ciągłą L.p.' }),
    lista], 'Dalej')) return;
  const lata = $$('input:checked', lista).map((c) => +c.dataset.rok).sort((a, b) => a - b);
  if (lata.length < 2) return toast('Zaznacz co najmniej dwa roczniki.');
  const wybrane = pozycje.filter((p) => lata.includes(p.r.rok));
  const ostrzezenia = [];
  const dziury = [];
  for (let y = lata[0]; y <= lata[lata.length - 1]; y++) if (!lata.includes(y)) dziury.push(y);
  if (dziury.length) {
    ostrzezenia.push(`W zakresie ${lata[0]}–${lata[lata.length - 1]} brakuje lat: ${dziury.join(', ')} (nazwa pliku i tak poda cały zakres).`);
  }
  const niepelne = wybrane.filter((p) => !p.pelny).map((p) => `${p.r.rok} (${p.info})`);
  if (niepelne.length) ostrzezenia.push(`Nie w pełni sprawdzone albo zmienione po zapisie xlsx: ${niepelne.join(', ')}.`);
  const nazwa = nazwaXlsxGeneteki(d.parafia, `${lata[0]}-${lata[lata.length - 1]}`);
  if (!await zapytaj('Scal lata', [
    el('p', {}, 'Powstanie plik ', el('b', { text: nazwa }), ` w folderze ${d.folder}.`),
    ostrzezenia.length ? el('ul', {}, ostrzezenia.map((o) => el('li', { text: o }))) : null,
  ].filter(Boolean), ostrzezenia.length ? 'Scal mimo to' : 'Scal')) return;
  try {
    const wszystkie = [];
    for (const p of wybrane) {
      const sep = p.r.sciezka.includes('\\') ? '\\' : '/';
      const adres = '/api/plik_xlsx?sciezka=' + enc(p.r.sciezka + sep + p.nazwa);
      let bufor;
      if (WWW) bufor = await WWW.binarny(adres);
      else {
        const r = await fetch(adres);
        if (!r.ok) throw new Error(`nie udało się odczytać ${p.nazwa} (błąd ${r.status})`);
        bufor = await r.arrayBuffer();
      }
      wszystkie.push(...await XL.czytajZgony(window.JSZip, bufor));
    }
    const plik = await XL.zbudujXlsx(window.JSZip, SZABLON, wszystkie);
    const z = await (await API.post('/api/xlsx_scalone', { folder: d.folder, nazwa, xlsx: naBase64(plik) })).json();
    toast(z.komunikat || `Zapisano ${nazwa} (${wszystkie.length} wierszy) w folderze ${z.folder}`);
  } catch (e) {
    toast('Nie udało się scalić lat: ' + e.message);
  }
}

/* Po pobraniu xlsx: propozycje do słowniczka (prompt dla AI) z aktów oznaczonych jako sprawdzone.
   Zaznaczone serwer dopisuje do slowniczek.md, odznaczone zapamiętuje i już ich nie proponuje. */
/* Linki do skanów: zakresy wycinków (od–do) z linkiem do zdjęcia strony księgi, np. w Skanotece.
   Zapisują się w odczyt_zweryfikowany.json rocznika, a w pobieranym xlsx trafiają do kolumny "Skan". */
function linkSkanu(wycinek) {
  const n = parseInt(wycinek, 10);
  const l = (st.linki || []).find((x) => n >= x.od && n <= x.do);
  if (l) return l.url;
  if (!GENETEKA) return null;
  // bez zakresów: strona księgi z Tartaka (z którego zdjęcia jest wycinek) w jednostce z „Link do skanów”
  const jednostka = st.link || K.linkRocznika({ linki_skanow: st.linki });
  const m = jednostka.match(/id(\d+)-sy(\d+)-kt(\d+)/);
  const plik = st.skany && st.skany[wycinek];
  if (m && plik) return `https://skanoteka.genealodzy.pl/index.php?op=pg&id=${m[1]}&se=&sy=${m[2]}&kt=${m[3]}&plik=${enc(plik)}`;
  return jednostka || null;
}

/** Tabela „od wycinka – do wycinka – link do strony księgi” (w oknie linków). odczytaj() -> lista albo null (błąd). */
function tabelaZakresow() {
  const tbody = el('tbody');
  const dodajWiersz = (l = {}) => {
    const od = el('input', { type: 'number', min: 1, value: l.od ?? '' });
    const doo = el('input', { type: 'number', min: 1, value: l.do ?? '' });
    const url = el('input', { type: 'url', placeholder: 'https://skanoteka.genealodzy.pl/…', value: l.url || '' });
    const tr = el('tr', {}, el('td', {}, od), el('td', {}, doo), el('td', {}, url),
      el('td', {}, el('button', { text: '×', title: 'Usuń wiersz', onclick: () => tr.remove() })));
    tr.pola = { od, doo, url };
    tbody.append(tr);
  };
  const ostatni = st.wycinki.length ? parseInt(st.wycinki[st.wycinki.length - 1].wycinek, 10) : '';
  (st.linki.length ? st.linki : [{ od: 1, do: ostatni }]).forEach(dodajWiersz);
  const wezly = [
    el('table', { class: 'linki-skanow' },
      el('thead', {}, el('tr', {}, el('th', { text: 'Od wycinka' }), el('th', { text: 'Do wycinka' }),
        el('th', { text: 'Link do skanu' }), el('th'))),
      tbody),
    el('div', {}, el('button', { text: '+ Dodaj wiersz', onclick: () => dodajWiersz({}) })),
  ];
  const odczytaj = () => {
    const linki = [];
    for (const tr of $$('tr', tbody)) {
      const od = parseInt(tr.pola.od.value, 10), doo = parseInt(tr.pola.doo.value, 10), url = tr.pola.url.value.trim();
      if (!url) continue;                     // wiersz bez linku (np. sam podpowiedziany zakres) się nie liczy
      if (!(od >= 1) || !(doo >= od)) { toast(`Zły zakres wycinków: ${tr.pola.od.value || '?'}–${tr.pola.doo.value || '?'}`); return null; }
      if (!/^https?:\/\//i.test(url)) { toast(`Wiersz ${od}–${doo}: link musi zaczynać się od http:// albo https://`); return null; }
      linki.push({ od, do: doo, url });
    }
    return linki.sort((a, b) => a.od - b.od);
  };
  return { wezly, odczytaj };
}

/* Format Geneteki: „Link do skanów” = link do jednostki w Skanotece (np. https://skanoteka.genealodzy.pl/id5557-sy5500-kt10)
   – trafia do kolumny „plik” wszystkich aktów (zapis: link_jednostki w odczyt_zweryfikowany.json). Numer skanu
   w uwagach bierze się sam z Tartaka (pocięte/skany.json); zakresy wycinków są potrzebne tylko bez niego. */
function pokazLinkJednostki() {
  const wejscie = el('input', { type: 'url', placeholder: 'https://skanoteka.genealodzy.pl/id5557-sy5500-kt10' });
  wejscie.value = st.link || K.linkRocznika({ linki_skanow: st.linki }) || '';
  const zakresy = tabelaZakresow();
  const zTartaka = st.wycinki.filter((w) => st.skany && st.skany[w.wycinek]).length;
  const zapisz = () => {
    const v = wejscie.value.trim();
    if (v && !/^https?:\/\//i.test(v)) return toast('Link musi zaczynać się od http:// albo https://');
    const linki = zakresy.odczytaj();
    if (!linki) return;
    const jednostka = K.linkJednostki(v);
    if (v && !jednostka) toast('To nie wygląda na link do jednostki w Skanotece (…/id…-sy…-kt…) – zapisuję, jak jest.');
    st.link = jednostka || v;
    st.linki = linki;
    st.linkiZmieniono = Date.now();
    zaplanujZapis();
    odswiezLinkiSkanow();
    odswiezWszystkieKarty();
    zamknijOkno();
    toast('Zapisano link do skanów.' + (linki.length ? ` Zakresów: ${linki.length}.` : ''));
  };
  wejscie.addEventListener('keydown', (e) => { if (e.key === 'Enter') zapisz(); });
  let opisTartaka;
  if (zTartaka && zTartaka === st.wycinki.length) {
    opisTartaka = `Numer skanu w uwagach bierze się sam z Tartaka (wiadomo, z którego zdjęcia jest każdy z ${zTartaka} wycinków) – zakresy nie są potrzebne.`;
  } else if (zTartaka) {
    opisTartaka = `Tartak zapisał, z którego zdjęcia jest ${zTartaka} z ${st.wycinki.length} wycinków – dla reszty numer skanu weźmie się z zakresów poniżej.`;
  } else {
    opisTartaka = 'Tartak nie zapisał, z którego zdjęcia jest który wycinek – numer skanu w uwagach weźmie się z zakresów poniżej (jeśli je wpiszesz).';
  }
  okno('Link do skanów', [
    el('p', { text: 'Wklej link do jednostki w Skanotece (rocznika). Trafi do kolumny „plik” w formacie Geneteki wszystkich aktów tego rocznika.' }),
    wejscie,
    el('p', { text: opisTartaka }),
    el('details', { open: zTartaka < st.wycinki.length || st.linki.length > 0 },
      el('summary', { text: 'Zakresy wycinków → link do strony księgi (nieobowiązkowe)' }),
      el('p', { text: 'Np. wycinki 1–12 → link do zdjęcia strony 002 w Skanotece. Numer skanu bierze się z końcówki linku (plik=002.JPG).' }),
      ...zakresy.wezly),
  ], [el('button', { class: 'glowny', text: 'Zapisz', onclick: zapisz })], { trwale: true });
  setTimeout(() => { wejscie.focus(); wejscie.select(); }, 50);
}

function odswiezLinkiSkanow() {
  for (const k of $$('.karta')) {
    const a = $('.link-skanu', k), url = linkSkanu(k.dataset.wycinek);
    a.classList.toggle('ukryta', !url);
    a.href = url || '#';
  }
}

function pokazLinkiSkanow() {
  if (GENETEKA) return pokazLinkJednostki();
  const zakresy = tabelaZakresow();
  okno('Linki do skanów', [
    el('p', { text: 'Wpisz zakres wycinków i link do zdjęcia strony księgi, z której pochodzą (np. wycinki 1–12 → strona 2 '
      + 'w Skanotece). Link pokaże się przy wycinku i trafi do kolumny „Skan” w pobieranym xlsx.' }),
    ...zakresy.wezly,
  ], [
    el('button', { class: 'glowny', text: 'Zapisz', onclick: () => {
      const linki = zakresy.odczytaj();
      if (!linki) return;
      st.linki = linki;
      st.linkiZmieniono = Date.now();
      zaplanujZapis();
      odswiezLinkiSkanow();
      odswiezWszystkieKarty();
      zamknijOkno();
      const bez = st.wycinki.filter((w) => !linkSkanu(w.wycinek)).length;
      toast(`Zapisano linki do skanów (${linki.length}).` + (bez ? ` Bez linku: ${bez} wycinków.` : ''));
    } }),
  ], { trwale: true });
}

async function pokazPropozycjeSlowniczka() {
  let d;
  try {
    d = await API.get('/api/slowniczek?sciezka=' + enc(st.sciezka));
  } catch (e) {
    return toast('Nie udało się przygotować propozycji do słowniczka: ' + e.message);
  }
  if (!d.propozycje.length) return;
  const pozycje = d.propozycje.map((p) => {
    const pole = el('input', { type: 'checkbox', checked: true });
    const ile = p.wycinki.length;
    return { p, pole, wiersz: el('label', { class: 'propozycja' }, pole,
      el('span', {}, p.opis, el('small', { text: ` – ${ile === 1 ? 'wycinek' : 'wycinki'} ${p.wycinki.join(', ')}` }))) };
  });
  okno('Propozycje do słowniczka', [
    el('p', { text: 'Z aktów oznaczonych jako sprawdzone. Zaznaczone dopiszę do słowniczka, który AI dostaje w prompcie '
      + 'przy następnych rocznikach. Odznaczonych nie będę już proponować.' }),
    el('div', { class: 'propozycje' }, pozycje.map((x) => x.wiersz)),
  ], [
    el('button', { class: 'glowny', text: 'Dopisz zaznaczone', onclick: async () => {
      try {
        const r = await API.post('/api/slowniczek', {
          przyjete: pozycje.filter((x) => x.pole.checked).map((x) => x.p),
          odrzucone: pozycje.filter((x) => !x.pole.checked).map((x) => x.p.klucz),
        });
        zamknijOkno();
        toast(`Dopisano do słowniczka: ${(await r.json()).dopisano}.`);
      } catch (e) {
        toast('Nie udało się dopisać do słowniczka: ' + e.message);
      }
    } }),
  ]);
}

// ---------------------------------------------------------------- okna: prompt, wgraj

// trwale: okno nie znika po kliknieciu obok ani Esc - tylko przyciskiem (zeby nie stracic wpisanych linkow)
function okno(tytul, cialo, przyciski, { trwale = false } = {}) {
  $('#okno').dataset.trwale = trwale ? '1' : '';
  $('#okno-tytul').textContent = tytul;
  $('#okno-cialo').replaceChildren(...cialo);
  const zamknij = el('button', { text: 'Zamknij', onclick: zamknijOkno });
  $('#okno-przyciski').replaceChildren(...przyciski, zamknij);
  $('#okno').classList.remove('ukryta');
}
function zamknijOkno() { $('#okno').classList.add('ukryta'); }

async function pokazPrompt() {
  let d;
  try {
    d = await API.get('/api/prompt' + (st.sciezka ? '?sciezka=' + enc(st.sciezka) : ''));
  } catch (e) {
    return toast('Nie udało się zbudować promptu: ' + e.message);
  }
  const pole = el('textarea', { readonly: true });
  pole.value = d.prompt;
  const opis = st.sciezka
    ? `Prompt dla folderu ${st.sciezka}. Skopiuj i wklej w Antigravity. Gdy AI skończy, kliknij „Odśwież”.`
    : 'Nie wybrano rocznika – na górze promptu wklej ścieżkę folderu „pocięte”. Łatwiej: kliknij rocznik po lewej i otwórz prompt jeszcze raz.';
  okno('Prompt dla Antigravity', [el('p', { text: opis }), pole], [
    el('button', { class: 'glowny', text: 'Kopiuj', onclick: async () => {
      try { await navigator.clipboard.writeText(pole.value); toast('Skopiowano prompt.'); }
      catch (e) { pole.select(); document.execCommand('copy'); toast('Skopiowano prompt.'); }
    } }),
  ]);
}

function pokazWgraj() {
  const wejscie = el('input', { type: 'text', placeholder: 'np. C:\\Users\\...\\SKANY MAM DOŚĆ\\Przytkowice\\Zgony\\1890\\pocięte' });
  const otworz = () => { const s = wejscie.value.trim().replace(/^"|"$/g, ''); if (s) { zamknijOkno(); otworzFolder(s); } };
  wejscie.addEventListener('keydown', (e) => { if (e.key === 'Enter') otworz(); });
  okno('Wgraj skany', [
    el('p', { text: 'Roczniki z „SKANY MAM DOŚĆ” otwierasz kliknięciem w drzewie po lewej. Tutaj możesz wkleić ścieżkę dowolnego folderu z wycinkami (albo rocznika, w którym jest folder „pocięte”).' }),
    wejscie,
  ], [el('button', { class: 'glowny', text: 'Otwórz', onclick: otworz })]);
  setTimeout(() => wejscie.focus(), 50);
}

async function odswiez() {
  await zapiszTeraz();
  await odswiezDrzewo();
  if (st.sciezka) await otworzFolder(st.sciezka);
}

// ---------------------------------------------------------------- strona www dla pomocnika

async function wyslijNaStrone() {
  await zapiszTeraz();
  let s;
  try {
    s = await API.get('/api/strona');
  } catch (e) {
    return toast('Nie udało się: ' + e.message);
  }
  if (s.token) wyslij();
  else pokazToken(s);
}

async function wyslij() {
  const b = $('#b-wyslij');
  b.disabled = true;
  b.textContent = 'Wysyłanie…';
  try {
    const d = await (await API.post('/api/wyslij', {})).json();
    toast(d.opis);
    if (d.poprawek) await odswiez();
  } catch (e) {
    toast('Nie udało się wysłać na stronę: ' + e.message);
  } finally {
    b.disabled = false;
    b.textContent = 'Wyślij na stronę';
  }
}

/* Token GitHuba pozwala stronie www zapisywać poprawki pomocnika (trafia na stronę zaszyfrowany). */
function pokazToken(s) {
  const wejscie = el('input', { type: 'password', placeholder: 'github_pat_…', autocomplete: 'off' });
  const adres = 'https://github.com/settings/personal-access-tokens/new?name=Weryfikacja+strona'
    + '&description=Zapis+poprawek+ze+strony+weryfikacji&target_name=wrobelkarol&contents=write';
  const zapisz = async () => {
    try {
      await API.post('/api/token', { token: wejscie.value });
    } catch (e) {
      return toast(e.message);
    }
    zamknijOkno();
    wyslij();
  };
  wejscie.addEventListener('keydown', (e) => { if (e.key === 'Enter') zapisz(); });
  okno('Token do zapisywania poprawek', [
    el('p', { text: `Strona dla pomocnika: ${s.adres} (login ${s.login}). Żeby poprawki pomocnika wracały do Ciebie, strona potrzebuje tokenu GitHuba – wystarczy raz:` }),
    el('ol', {},
      el('li', {}, el('a', { href: adres, target: '_blank', rel: 'noopener', text: 'Otwórz tworzenie tokenu na GitHubie' })),
      el('li', { text: 'Repository access: „Only select repositories” → weryfikacja.' }),
      el('li', { text: 'Permissions → Repository permissions → Contents: „Read and write”.' }),
      el('li', { text: 'Expiration: najdłuższa możliwa (albo „No expiration”).' }),
      el('li', { text: '„Generate token”, skopiuj go (zaczyna się od github_pat_) i wklej poniżej.' })),
    wejscie,
  ], [
    el('button', { class: 'glowny', text: 'Zapisz i wyślij', onclick: zapisz }),
    el('button', { text: 'Wyślij bez tokenu', onclick: () => { zamknijOkno(); wyslij(); } }),
  ]);
  setTimeout(() => wejscie.focus(), 50);
}

let toastTimer = null;
function toast(tekstT) {
  const t = $('#toast');
  t.textContent = tekstT;
  t.classList.remove('ukryta');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.add('ukryta'), 5000);
}

// ---------------------------------------------------------------- start

async function start() {
  const d = await API.get('/api/start');
  POLA_KSIAG = { zgony: d.pola, urodzenia: d.pola_urodzenia || [] };
  TYPY = d.typy;
  RODZAJE = d.rodzaje;
  try {
    // format Geneteki: konwersja i zapis xlsx (te same pliki co w programie „genetyka weryfikacja”)
    [K, XL] = await Promise.all([import('./konwersja.js'), import('./xlsx.js')]);
    const r = await fetch('indeksy.xlsx');
    if (!r.ok) throw new Error(`brak szablonu indeksy.xlsx (błąd ${r.status})`);
    SZABLON = await r.arrayBuffer();
    SLOWNIK = K.zbudujSlownik(await XL.czytajSlownikImion(window.JSZip, SZABLON));
  } catch (e) {
    K = XL = null;
    toast('Nie udało się wczytać formatu Geneteki (' + e.message + ') – karty pokazują łacinę i polski jak dawniej.');
  }
  ustawKsiege(null);
  try { st.otwarte = localStorage.getItem('zwiniete') !== '1'; } catch (e) { /* bez znaczenia */ }
  try { st.pokazPuste = localStorage.getItem('pokazPuste') === '1'; } catch (e) { /* bez znaczenia */ }
  $('#karty').classList.toggle('pokaz-puste', st.pokazPuste);
  $('#b-zwin').textContent = st.otwarte ? 'Zwiń wszystkie pola' : 'Rozwiń wszystkie pola';

  $('#karty').addEventListener('input', obsluzZmiane);
  $('#karty').addEventListener('change', obsluzZmiane);
  $('#karty').addEventListener('focusin', zapamietajEdycje);
  $('#karty').addEventListener('focusout', () => { st.edycja = null; });
  $('#karty').addEventListener('keydown', strzalki);
  $('#b-pobierz').addEventListener('click', pobierz);
  $('#b-prompt').addEventListener('click', pokazPrompt);
  $('#b-wgraj').addEventListener('click', pokazWgraj);
  $('#b-odswiez').addEventListener('click', odswiez);
  $('#b-wyslij').addEventListener('click', wyslijNaStrone);
  $('#filtr').addEventListener('change', (e) => { st.filtr = e.target.value; zastosujFiltr(); });
  $('#b-linki').addEventListener('click', pokazLinkiSkanow);
  $('#b-scal').addEventListener('click', scalLata);
  // Enter = wycinek sprawdzony i następny (w polu wielowierszowym: Ctrl+Enter); jak w „genetyka weryfikacja”
  document.addEventListener('keydown', (e) => {
    if (e.key !== 'Enter' || e.shiftKey || e.altKey || e.metaKey || e.isComposing) return;
    if (!$('#okno').classList.contains('ukryta') || !st.wycinki.length) return;
    const pole = e.target instanceof Element ? e.target.closest('input, textarea, select, button, a, summary, [contenteditable]') : null;
    if (pole && !e.ctrlKey && ['TEXTAREA', 'BUTTON', 'A', 'SUMMARY'].includes(pole.tagName)) return;
    if (pole && !pole.closest('.karta')) return;
    const karta = (pole && pole.closest('.karta')) || biezacaKarta();
    if (!karta || !karta.ustawSprawdzony) return;
    e.preventDefault();
    if (pole) pole.blur();
    karta.ustawSprawdzony(true, true);
  });
  $('#b-zwin').addEventListener('click', () => {
    st.otwarte = !st.otwarte;
    try { localStorage.setItem('zwiniete', st.otwarte ? '0' : '1'); } catch (e) { /* bez znaczenia */ }
    $('#b-zwin').textContent = st.otwarte ? 'Zwiń wszystkie pola' : 'Rozwiń wszystkie pola';
    $$('.karta details').forEach((det) => { det.open = st.otwarte; });
  });
  $('#okno').addEventListener('click', (e) => { if (e.target.id === 'okno' && !$('#okno').dataset.trwale) zamknijOkno(); });
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && !$('#okno').dataset.trwale) zamknijOkno(); });
  window.addEventListener('beforeunload', (e) => { if (st.brudne) { zapiszTeraz(); e.preventDefault(); e.returnValue = ''; } });
  document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'hidden') zapiszTeraz(); });

  if (!WWW) {
    const ping = () => fetch('/api/ping').catch(() => ustawStanZapisu('Program zamknięty – uruchom go ponownie skrótem.', true));
    ping();
    setInterval(ping, 20000);
  }

  const naglowek = $('header');
  const ustawWysokosc = () => document.documentElement.style.setProperty('--wys-naglowka', naglowek.offsetHeight + 'px');
  new ResizeObserver(ustawWysokosc).observe(naglowek);
  ustawWysokosc();

  await odswiezDrzewo();
  let ostatni = new URLSearchParams(location.search).get('folder');
  try { ostatni = ostatni || localStorage.getItem('ostatni'); } catch (e) { /* bez znaczenia */ }
  if (ostatni) otworzFolder(ostatni).catch(() => {});
}

(WWW ? WWW.zaloguj() : Promise.resolve()).then(start).catch((e) => {
  document.body.prepend(el('div', { class: 'komunikat blad', text: 'Błąd uruchamiania: ' + e.message }));
});
