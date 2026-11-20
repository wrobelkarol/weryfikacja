// Konwersja odczytu z programu Weryfikacja (odczyt_zweryfikowany.json, pole polski{}) na wiersze arkusza Zgony
// w formacie Geneteki (kolumny A–S szablonu indeksy.xlsx). Ten sam moduł działa w przeglądarce i w Node (test).

export const KOLUMNY = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S'];
export const NAGLOWKI = {
  A: 'L.p.', B: 'rok', C: 'akt', D: 'sufix', E: '1 imię', F: 'imiona', G: 'nazwisko z księgi', H: 'nazwisko',
  I: 'vel', J: 'płeć', K: 'miejscowość', L: 'plik', M: 'uwagi', N: 'data', O: 'imię ojca', P: 'imię matki',
  Q: 'nazwisko matki', R: 'ID1', S: 'ID2',
};
export const LICZBOWE = new Set(['A', 'B', 'C', 'R']);          // w xlsx jako liczby (jak w plikach wysłanych do Geneteki)
export const WYMAGANE = ['B', 'C', 'E', 'H', 'N', 'R'];
export const ID_INDEKSUJACEGO = '110608';
// kolejność segmentów w kolumnie uwagi
export const SEGMENTY = ['dom', 'wiek', 'pogrzeb', 'przyczyna zgonu', 'małżonek', 'stan', 'stan męża', 'stan ojca',
  'stan matki', 'nr skanu'];

const OGOLNE_IMIONA = new Set(['chłopiec', 'dziewczynka', 'chłopcy', 'dziewczynki', 'dziecko', 'dzieci', 'bliźnięta',
  'bliźniak', 'bliźniaczka', 'puer', 'puella', 'pueri', 'puellae', 'infans', 'proles', 'nn', 'n.n.', 'nieznane', 'bez imienia']);
const DUZE = 'A-ZĄĆĘŁŃÓŚŹŻ';
const MIESIACE_LAC = ['jan', 'feb', 'mar', 'apr', 'mai', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
const MIESIACE_PL = ['sty', 'lut', 'mar', 'kwi', 'maj', 'cze', 'lip', 'sie', 'wrz', 'paź', 'lis', 'gru'];

// ------------------------------------------------------------------ znaki i liczby

const ULAMKI = { '½': 0.5, '¼': 0.25, '¾': 0.75, '⅓': 1 / 3, '⅔': 2 / 3, '⅙': 1 / 6, '⅚': 5 / 6, '⅛': 0.125 };

/** Usuwa znaki zabronione przez Genetekę: \ ( ) ' " * ; i wielokropek, poprawia „.,” i zbędne spacje. */
export function czysc(s) {
  return String(s ?? '')
    .replace(/\.\.\.|…/g, ' ')
    .replace(/[\\()'"*;„”“‚‘’`]/g, '')
    .replace(/\.,/g, ',')
    .replace(/\s+,/g, ',')
    .replace(/\s{2,}/g, ' ')
    .trim();
}

/** czysc + bez kropek i przecinków na końcu (nazwiska, imiona, numery). */
export function czyscKoniec(s) {
  return czysc(s).replace(/[\s.,]+$/, '').replace(/^[\s.,]+/, '');
}

export function pusty(s) {
  const t = String(s ?? '').trim();
  return !t || /^[—–\-?.]+$/.test(t);
}

export function liczba(v) {
  const m = String(v ?? '').match(/\d+/);
  return m ? parseInt(m[0], 10) : null;
}

/** Liczba dziesiętna po polsku: 1.5 -> „1,5”. */
export function dziesietna(x) {
  const t = (Math.round(x * 100) / 100).toString();
  return t.replace('.', ',');
}

/** Ułamki zwykłe w tekście (½, 1/2, 10/12, „1 1/2”) zamienia na liczby dziesiętne. */
export function ulamkiNaDziesietne(s) {
  let t = String(s ?? '');
  t = t.replace(/(\d+)?\s?([½¼¾⅓⅔⅙⅚⅛])/g, (_, c, u) => dziesietna((c ? +c : 0) + ULAMKI[u]));
  t = t.replace(/(?:(\d+)\s+)?(\d+)\/(\d+)(?!\d)/g, (cale, c, l, m) => (+m ? dziesietna((c ? +c : 0) + +l / +m) : cale));
  return t;
}

// ------------------------------------------------------------------ wiek

// jednostki wieku od największej; dni = ile dni ma jednostka (do sumowania „4 lata 3 miesiące”)
const JEDNOSTKI = [
  { re: /^(rok|roku|lat|lata|latek|r$|ann)/i, formy: ['rok', 'lata', 'lat', 'roku'], dni: 365.25 },
  { re: /^(miesiąc|miesiąca|miesiące|miesięcy|mies|m-c|mens|m$)/i, formy: ['miesiąc', 'miesiące', 'miesięcy', 'miesiąca'], dni: 30.4375 },
  { re: /^(tydzień|tygodnie|tygodni|tygodnia|tydz|tyg|hebd)/i, formy: ['tydzień', 'tygodnie', 'tygodni', 'tygodnia'], dni: 7 },
  { re: /^(dzień|dni|dnia|dzie|dier|dies|die|d$)/i, formy: ['dzień', 'dni', 'dni', 'dnia'], dni: 1 },
  { re: /^(godzina|godziny|godzin|godz|hor)/i, formy: ['godzina', 'godziny', 'godzin', 'godziny'], dni: 1 / 24 },
];
// liczba z ułamkiem i (opcjonalnie) jednostką: „4 9/12 roku”, „1/2 roku”, „40½ lat”, „1,5 roku”, „6 mens.”
const RE_WIEK = /(?:(?:(\d+)\s+)?(\d+)\s*\/\s*(\d+)|(\d+(?:\.\d+)?)?\s*([½¼¾⅓⅔⅙⅚⅛])|(\d+(?:\.\d+)?))\s*\.?\s*([a-ząćęłńóśźż-]+)?\.?/gi;

/** Polska odmiana: 1 rok, 2 lata, 5 lat, 1,5 roku. */
export function odmiana(n, formy) {
  if (!Number.isInteger(n)) return formy[3];
  if (n === 1) return formy[0];
  const j = n % 10, d = n % 100;
  if (j >= 2 && j <= 4 && !(d >= 12 && d <= 14)) return formy[1];
  return formy[2];
}

/** Zaokrąglenie do całości albo do ,5 (połówki w górę: 1,25 -> 1,5; 1,75 -> 2). */
export function zaokraglij(x) {
  return Math.floor(x * 2 + 0.5 + 1e-9) / 2;
}

/**
 * Wiek do uwag. Zwraca {tekst, ulamek, uwaga}: ulamek = w źródle była część ułamkowa (wtedy wiek jest zaokrąglany),
 * uwaga = powód, dla którego pole trzeba sprawdzić.
 */
export function wiek(zrodlo) {
  const surowy = String(zrodlo ?? '').trim();
  if (pusty(surowy)) return { tekst: '', ulamek: false };
  if (/\bn\.|\bur\.|urodz|natus|nata\b|\b1[5-9]\d\d\b/i.test(surowy)) {
    return { tekst: czysc(surowy), ulamek: false, uwaga: `zamiast wieku data urodzenia: „${surowy}”` };
  }
  const t = surowy.replace(/\(.*?\)/g, ' ').replace(/(\d),(\d)/g, '$1.$2').replace(/\s+/g, ' ').trim();   // „4 11/12 (5) lat” -> „4 11/12 lat”
  const nawias = /\(/.test(surowy);
  const czesci = [];
  let nieznana = false;
  for (const m of t.matchAll(RE_WIEK)) {
    if (!m[0].trim()) continue;
    let x;
    if (m[2] !== undefined) { if (!+m[3]) continue; x = (m[1] ? +m[1] : 0) + +m[2] / +m[3]; }
    else if (m[5]) x = (m[4] ? +m[4] : 0) + ULAMKI[m[5]];
    else x = +m[6];
    const j = m[7] ? JEDNOSTKI.find((jj) => jj.re.test(m[7])) : null;
    if (m[7] && !j) nieznana = true;
    czesci.push({ x, j });
  }
  if (!czesci.length || nieznana) {
    return { tekst: czysc(ulamkiNaDziesietne(surowy)), ulamek: false, uwaga: `wiek w nietypowym zapisie: „${surowy}”` };
  }
  // jedna jednostka: ta sama; kilka („4 lata 3 miesiące”): wszystko w największej z nich
  const jednostki = czesci.map((c) => c.j).filter(Boolean);
  const glowna = jednostki.length ? JEDNOSTKI.find((jj) => jednostki.includes(jj)) : null;
  const x = czesci.length === 1 ? czesci[0].x
    : czesci.reduce((s, c) => s + c.x * (c.j || glowna).dni, 0) / glowna.dni;
  const ulamek = !Number.isInteger(x);
  const y = zaokraglij(x);
  const tekst = glowna ? `${dziesietna(y)} ${odmiana(y, glowna.formy)}` : dziesietna(y);   // sama liczba zostaje liczbą
  const uwagi = [];
  if (Math.abs(y - x) > 1e-9) uwagi.push(`wiek zaokrąglony z „${surowy}”`);
  if (nawias) uwagi.push(`wiek z nawiasem w źródle: „${surowy}”`);
  return { tekst, ulamek, uwaga: uwagi.join('; ') || undefined };
}

// ------------------------------------------------------------------ daty

function miesiacZNazwy(s) {
  const t = String(s || '').toLowerCase().normalize('NFC');
  for (let i = 0; i < 12; i++) if (t.startsWith(MIESIACE_LAC[i]) || t.startsWith(MIESIACE_PL[i])) return i + 1;
  if (t.startsWith('iun')) return 6;
  if (t.startsWith('iul')) return 7;
  if (t.startsWith('ian')) return 1;
  return null;
}

/** „28” -> {d:28}; „28/2” -> {d:28, m:2}; „2.1.1894” -> {d:2, m:1, r:1894}; „2 Januarii 1897” -> {d:2, m:1, r:1897}. */
export function dzienMiesiacRok(v) {
  const t = String(v ?? '').trim();
  if (pusty(t)) return {};
  let m = t.match(/^(\d{1,2})\s*[./]?\s*(?:[./]\s*(\d{1,2})\s*\.?\s*)?(?:[./]?\s*(1[5-9]\d\d)\s*\.?)?\s*$/);
  if (m) return { d: +m[1], m: m[2] ? +m[2] : null, r: m[3] ? +m[3] : null };
  m = t.match(/^(\d{1,2})\s*\.?\s*([A-Za-ząćęłńóśźżĄĆĘŁŃÓŚŹŻ]+)\.?\s*(1[5-9]\d\d)?\s*\.?$/);
  if (m && miesiacZNazwy(m[2])) return { d: +m[1], m: miesiacZNazwy(m[2]), r: m[3] ? +m[3] : null };
  return { nieznane: t };
}

function dwie(n) {
  return String(n).padStart(2, '0');
}

/** Data zgonu (kolumna „data”) w formacie wymaganym przez Genetekę: RRRR.MM.DD. */
export function data(r, m, d) {
  return `${r}.${m ? dwie(m) : '00'}.${d ? dwie(d) : '00'}`;
}

/** Data w uwagach (pogrzeb): dd.mm.rrrr. */
export function dataUwag(r, m, d) {
  return `${d ? dwie(d) : '00'}.${m ? dwie(m) : '00'}.${r}`;
}

// ------------------------------------------------------------------ słownik imion łacińskich

/** Rozwija zapis ze słownika: „A(da)lbert(us)” -> Albert, Adalbert, Albertus, Adalbertus; „Cai(-j)etanus” -> Caietanus, Cajetanus. */
function rozwin(forma) {
  let wyniki = [forma];
  for (let i = 0; i < 6; i++) {
    const nowe = [];
    let zmiana = false;
    for (const f of wyniki) {
      const m = f.match(/\((-?)([^()]*)\)/);
      if (!m) { nowe.push(f); continue; }
      zmiana = true;
      const przed = f.slice(0, m.index), po = f.slice(m.index + m[0].length);
      if (m[1]) {                                       // (-j): litera przed nawiasem albo zamieniona
        nowe.push(przed + po, przed.slice(0, -1) + m[2] + po);
      } else if (/\s/.test(m[2]) || /^[A-ZĄĆĘŁŃÓŚŹŻ]/.test(m[2])) {
        nowe.push((przed + po).trim());                  // „(Joanne)”, „(także Gotfryd)” – dopisek, nie część imienia
      } else {
        nowe.push(przed + po, przed + m[2] + po);
      }
    }
    wyniki = nowe;
    if (!zmiana) break;
  }
  return wyniki;
}

function kluczImienia(s) {
  return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/[^a-zł]/g, '');
}

/**
 * Słownik z arkusza „Słownik imion łacińskich” szablonu: wiersze [mianownik, dopełniacz, biernik, skrót, polskie].
 * Zwraca {lacina: Map(klucz -> [polskie...]), polskie: Set(klucz)}.
 */
export function zbudujSlownik(wiersze) {
  const lacina = new Map();
  const polskie = new Set();
  for (const w of wiersze || []) {
    const pl = String(w[4] || '').split(/\s+-\s+|\s+czy\s+/)[0].replace(/\(.*?\)/g, '').replace(/\?/g, '');
    const formyPl = pl.split(/[,/]/).map((x) => x.trim()).filter((x) => /^[A-ZĄĆĘŁŃÓŚŹŻ][a-ząćęłńóśźż]+$/.test(x));
    if (!formyPl.length) continue;
    formyPl.forEach((x) => polskie.add(kluczImienia(x)));
    for (const kol of [w[0], w[1], w[2]]) {
      for (const czesc of String(kol || '').split(/,|\||\s+czy\s+/)) {
        for (const f of rozwin(czesc.trim())) {
          if (!f || /\s/.test(f.trim())) continue;
          const k = kluczImienia(f);
          if (k && !lacina.has(k)) lacina.set(k, formyPl);
        }
      }
    }
  }
  return { lacina, polskie };
}

/** Imię łacińskie -> {imie, uwaga}; imię polskie (albo nieznane) zostaje bez zmian. */
export function polskieImie(imie, slownik) {
  if (!slownik || !imie) return { imie };
  const k = kluczImienia(imie);
  if (slownik.polskie.has(k)) return { imie };
  const pl = slownik.lacina.get(k);
  if (!pl) return { imie };
  if (kluczImienia(pl[0]) === k) return { imie };
  const uwaga = pl.length > 1
    ? `imię łacińskie „${imie}” – ze słownika: ${pl.join(' / ')} (wybrano pierwsze)`
    : `imię łacińskie „${imie}” zamienione na „${pl[0]}”`;
  return { imie: pl[0], uwaga };
}

// ------------------------------------------------------------------ nazwiska i przyczyny

export function zenskie(n) {
  for (const [a, b] of [['dzki', 'dzka'], ['cki', 'cka'], ['ski', 'ska']]) if (n.endsWith(a)) return n.slice(0, -a.length) + b;
  return n;
}

export function meskie(n) {
  for (const [a, b] of [['dzka', 'dzki'], ['cka', 'cki'], ['ska', 'ski']]) if (n.endsWith(a)) return n.slice(0, -a.length) + b;
  return n;
}

const NAZWY_WLASNE = /^(Brighta|Bright|Basedowa|Addisona|Parkinsona|Potta|Heinego|Medina|Świętego|Wita|Antoniego)\b/;

/** „Starość Zapalenie” -> [„starość”, „zapalenie”]; „Choroba Brighta (zapalenie nerek)” -> [„choroba Brighta zapalenie nerek”]. */
export function przyczyny(tekst) {
  // ukośnik (dwie wersje od AI) -> spacja, jak w plikach wysłanych do Geneteki: „drgawki/atak konwulsji” -> „drgawki atak konwulsji”
  const t = String(tekst ?? '').replace(/\s*\/\s*/g, ' ').trim();
  if (pusty(t)) return [];
  const czesci = [];
  for (const kawalek of t.split(/\s*,\s*/)) {
    // kilka przyczyn wklejonych po spacji, każda z wielkiej litery (bez rozdzielania nazw własnych, np. Brighta)
    const slowa = kawalek.split(/\s+/);
    let biezaca = [], wNawiasie = false;
    for (const s of slowa) {
      const duza = !wNawiasie && new RegExp(`^[${DUZE}]`).test(s) && !NAZWY_WLASNE.test(s);
      if (duza && biezaca.length) { czesci.push(biezaca.join(' ')); biezaca = []; }
      biezaca.push(s);
      if (s.includes('[')) wNawiasie = true;
      if (s.includes(']')) wNawiasie = false;
    }
    if (biezaca.length) czesci.push(biezaca.join(' '));
  }
  return czesci
    .map((c) => czysc(ulamkiNaDziesietne(c)).replace(/[\s.,]+$/, ''))
    .filter(Boolean)
    .map((c) => (/^\[/.test(c) ? c : c.charAt(0).toLowerCase() + c.slice(1)));
}

// ------------------------------------------------------------------ link do Skanoteki i nr skanu

/** Link do jednostki w Skanotece z dowolnego linku do skanu: …?op=pg&id=5557&sy=5500&kt=3&plik=001.JPG -> …/id5557-sy5500-kt3 */
export function linkJednostki(url) {
  const u = String(url || '');
  const p = (n) => (u.match(new RegExp(`[?&]${n}=(\\d+)`)) || [])[1];
  const id = p('id'), sy = p('sy'), kt = p('kt');
  if (id && sy && kt) return `https://skanoteka.genealodzy.pl/id${id}-sy${sy}-kt${kt}`;
  const m = u.match(/id(\d+)-sy(\d+)-kt(\d+)/);
  return m ? `https://skanoteka.genealodzy.pl/id${m[1]}-sy${m[2]}-kt${m[3]}` : '';
}

export function linkRocznika(dane) {
  for (const l of dane?.linki_skanow || []) {
    const link = linkJednostki(l.url);
    if (link) return link;
  }
  return '';
}

/** Numer skanu z nazwy pliku zdjęcia („002.JPG” -> 2). */
export function nrZPliku(plik) {
  const m = String(plik || '').match(/(\d+)\.[a-z]+$/i);
  return m ? parseInt(m[1], 10) : null;
}

export function nrSkanu(wycinek, linki) {
  const n = liczba(wycinek);
  if (n === null) return null;
  for (const l of linki || []) {
    if (+l.od <= n && n <= +l.do) {
      const m = String(l.url || '').match(/plik=0*(\d+)/i);
      if (m) return parseInt(m[1], 10);
    }
  }
  return null;
}

// ------------------------------------------------------------------ jeden akt

function rozbijUkosnik(s) {
  return String(s || '').split('/').map((x) => x.trim()).filter(Boolean);
}

/**
 * Wiersz Geneteki dla jednego aktu.
 * @returns {{wiersz: Object<string,string>, flagi: Object<string,string[]>, podpowiedzi: string[]}}
 */
export function konwertujAkt(wycinek, kontekst) {
  const p = wycinek.polski || {};
  const { rok, slownik, linki, miejscowosc, link } = kontekst;
  const flagi = {};
  const flaga = (kol, tekst) => { (flagi[kol] ||= []).push(tekst); };
  const podpowiedzi = [];

  // data zgonu (miesiąc wpisany wyżej obowiązuje kolejne akty)
  const mies = liczba(p.miesiac);
  if (mies && mies >= 1 && mies <= 12) kontekst.miesiac = mies;
  const dz = dzienMiesiacRok(p.dzien_zgonu);
  const miesiacZgonu = dz.m || kontekst.miesiac || null;
  if (!dz.d || !miesiacZgonu) flaga('N', 'niepełna data zgonu');
  if (dz.nieznane) flaga('N', `dzień zgonu w nietypowym zapisie: „${dz.nieznane}”`);

  // pogrzebu nie ma w uwagach (Weryfikacja, od 13.11.2026) – datę pogrzebu sprawdzają reguły na polach łaciny

  // imiona
  let imiona = czyscKoniec(p.zm_imie).split(/\s+/).filter(Boolean);
  if (imiona.some((x) => x.includes('/'))) {
    flaga('E', `kilka wersji imienia: „${p.zm_imie}” (wybrano pierwszą)`);
    imiona = imiona.map((x) => rozbijUkosnik(x)[0]);
  }
  const nn = !imiona.length || OGOLNE_IMIONA.has(imiona[0].toLowerCase());
  let E = 'NN', F = '';
  if (!nn) {
    const przetl = imiona.map((x) => polskieImie(x, slownik));
    przetl.forEach((x, i) => x.uwaga && flaga(i ? 'F' : 'E', x.uwaga));
    E = przetl[0].imie;
    F = przetl.slice(1).map((x) => x.imie).join(' ');
  }

  // nazwisko (H) i vel (I)
  const zn = czyscKoniec(p.zm_nazwisko), zr = czyscKoniec(p.zm_rodowe), on = czyscKoniec(p.oj_nazwisko), mr = czyscKoniec(p.ma_rodowe);
  let H, I = '';
  const nazwiska = rozbijUkosnik(zn);
  if (nazwiska.length > 1) {
    H = nazwiska[nazwiska.length - 1];
    flaga('H', `kilka nazwisk: „${zn}” – wzięto ostatnie (po ostatnim mężu)`);
    flaga('I', 'kilka nazwisk – sprawdź vel (panieńskie i wcześniejsze po spacji)');
    const reszta = [zr, ...nazwiska.slice(0, -1)].filter((x) => x && x !== H);
    I = [...new Set(reszta)].join(' ');
  } else if (zn) {
    H = zn;
    if (zr && zr !== H) I = rozbijUkosnik(zr).join(' ');
  } else if (on) {
    H = p.plec === 'K' ? zenskie(on) : on;
  } else if (mr) {
    H = p.plec === 'K' ? zenskie(mr) : p.plec === 'M' ? meskie(mr) : mr;
    if (p.typ !== 'dziecko_nieslubne') flaga('H', 'nazwisko wzięte od matki (brak nazwiska zmarłego i ojca)');
  } else {
    H = '';
    flaga('H', 'brak nazwiska – pole wymagane');
  }
  if (/\//.test(zr)) flaga('I', `kilka nazwisk rodowych: „${zr}”`);
  if (/\[|\?/.test(H + I)) flaga(/\[|\?/.test(H) ? 'H' : 'I', 'niepewny odczyt nazwiska');

  // rodzice
  let O = czyscKoniec(p.oj_imie);
  if (O.includes('/')) { flaga('O', `kilka wersji imienia ojca: „${O}” (wybrano pierwszą)`); O = rozbijUkosnik(O)[0]; }
  const oj = polskieImie(O, slownik); if (oj.uwaga) flaga('O', oj.uwaga); O = oj.imie;
  let P = czyscKoniec(p.ma_imie);
  if (P.includes('/')) { flaga('P', `kilka wersji imienia matki: „${P}” (wybrano pierwszą)`); P = rozbijUkosnik(P)[0]; }
  const ma = polskieImie(P, slownik); if (ma.uwaga) flaga('P', ma.uwaga); P = ma.imie;
  const Q = mr;
  if (/\[|\?/.test(Q)) flaga('Q', 'niepewny odczyt nazwiska matki');

  // uwagi
  const seg = {};
  const dom = czyscKoniec(p.dom);
  if (!pusty(dom)) seg.dom = dom;
  const w = wiek(p.wiek);
  if (w.tekst) seg.wiek = w.tekst;
  if (w.uwaga) flaga('M', w.uwaga);
  const prz = przyczyny(p.przyczyna);
  if (nn && p.typ === 'martwe_urodzenie' && !prz.some((x) => /martwo urodz/.test(x))) prz.push('martwo urodzone');
  if (prz.length) seg['przyczyna zgonu'] = prz.join(', ');
  if (/\[łac|\[\?|\?\]/.test(p.przyczyna || '')) flaga('M', 'przyczyna zgonu niedotłumaczona albo niepewna');
  if (/\//.test(p.przyczyna || '')) flaga('M', `przyczyna z ukośnikiem: „${p.przyczyna}” (ukośnik zamieniony na spację) – sprawdź`);
  if (!pusty(p.mz_imie) || !pusty(p.mz_nazwisko)) {
    let mz = `${czyscKoniec(p.mz_imie)} ${czyscKoniec(p.mz_nazwisko)}`.trim();
    if (!pusty(p.mz_rodowe)) mz += ` z domu ${czyscKoniec(p.mz_rodowe)}`;
    seg['małżonek'] = mz;
  }
  if (p.plec === 'M' && !pusty(p.zm_stan)) seg.stan = czyscKoniec(p.zm_stan);
  if (p.plec === 'K' && !pusty(p.mz_stan)) seg['stan męża'] = czyscKoniec(p.mz_stan);
  if (!pusty(p.oj_stan)) seg['stan ojca'] = czyscKoniec(p.oj_stan);
  if (!pusty(p.ma_stan)) seg['stan matki'] = czyscKoniec(p.ma_stan);
  // nr skanu: z Tartaka (pocięte/skany.json – z którego zdjęcia jest wycinek), a bez niego z zakresów linków
  const sk = nrZPliku(kontekst.skany && kontekst.skany[wycinek.wycinek]) ?? nrSkanu(wycinek.wycinek, linki);
  if (sk !== null) seg['nr skanu'] = String(sk);
  const M = SEGMENTY.filter((k) => seg[k]).map((k) => `${k}: ${seg[k]}`).join(', ');

  if (!pusty(p.dodatkowe)) {
    flaga('M', 'w akcie są dodatkowe informacje – dopisz do uwag tylko to, co jest w oryginale');
    podpowiedzi.push(p.dodatkowe.trim());
  }
  const uw = String(wycinek.uwagi || '');
  if (/dwa akty|bliźni|gemell/i.test(uw) || /^(chłopcy|dziewczynki|bliźnięta|pueri|gemelli)$/i.test(imiona[0] || '')) {
    flaga('C', 'możliwe bliźnięta (dwa akty w jednym wycinku) – zduplikuj wiersz');
  }

  const nr = liczba(p.nr);
  if (nr === null) flaga('C', 'brak numeru aktu');

  const wiersz = {
    A: '', B: String(rok), C: nr === null ? '' : String(nr), D: '', E, F, G: '', H, I, J: '', K: miejscowosc || '',
    L: link || '', M, N: data(rok, miesiacZgonu, dz.d), O, P, Q, R: ID_INDEKSUJACEGO, S: '',
  };
  if (!wiersz.L) flaga('L', 'brak linku do jednostki w Skanotece – wpisz link rocznika');
  for (const k of KOLUMNY) {
    const zakazane = /[\\()'"*;]|\.\.\.|…|\.,/.test(wiersz[k]) || /[½¼¾]|\d\/\d/.test(wiersz[k]);
    if (zakazane) flaga(k, 'znaki niedozwolone w Genetece');
  }
  return { wiersz, flagi, podpowiedzi };
}

/**
 * Wszystkie akty rocznika. dane = odczyt_zweryfikowany.json.
 * opcje: {rok, slownik, miejscowosc, link} (link = link do jednostki w Skanotece; domyślnie z linki_skanow).
 */
export function konwertujRocznik(dane, opcje) {
  const kontekst = {
    rok: +opcje.rok,
    slownik: opcje.slownik,
    linki: dane.linki_skanow || [],
    skany: opcje.skany || null,
    miejscowosc: opcje.miejscowosc ?? 'Przytkowice',
    link: opcje.link ?? linkRocznika(dane),
    miesiac: null,
  };
  const akty = (dane.wycinki || []).filter((w) => (w.rodzaj || 'akt') === 'akt');
  akty.sort((a, b) => (liczba(a.wycinek) ?? 0) - (liczba(b.wycinek) ?? 0));
  return akty.map((w) => ({ wycinek: String(w.wycinek), ...konwertujAkt(w, kontekst) }));
}

/** Rozbija kolumnę uwagi na segmenty {klucz: wartość}. */
export function segmentyUwag(u) {
  const wynik = {};
  for (const czesc of String(u || '').split(/, (?=[a-ząćęłńóśźż ]+: )/)) {
    const i = czesc.indexOf(': ');
    if (i > 0) wynik[czesc.slice(0, i).trim()] = czesc.slice(i + 2).trim();
  }
  return wynik;
}

/** Ostateczne czyszczenie wartości przed zapisem do xlsx (znaki zabronione, ułamki). */
export function doZapisu(kol, v) {
  let t = String(v ?? '').trim();
  if (!t) return '';
  if (kol === 'L') return t;
  t = czysc(ulamkiNaDziesietne(t));
  return t;
}
