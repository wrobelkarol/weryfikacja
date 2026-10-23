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
const st = {
  sciezka: null, opis: null, wycinki: [], mapa: new Map(), flagi: new Map(), aiZmieniony: null,
  brudne: false, zapisTimer: null, flagiTimer: null, otwarte: true, filtr: 'wszystkie', drzewo: null, edycja: null,
  kolumna: 'lacina',     // kolumna, w której zostaje strzałka w dół/górę po przejściu przez pole wspólne
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

const IMIONA_MESKIE_NA_A = ['kuba', 'barnaba', 'bonawentura', 'kosma', 'jarema', 'zawisza', 'sasza', 'bonaventura'];
const IMIONA_ZENSKIE_BEZ_A = ['agnes', 'elisabeth', 'elizabeth', 'beatrix', 'gertrudis', 'hedvigis', 'ruth', 'rachel', 'dorothea'];

function plecImienia(imie) {
  const s = ((imie || '').trim().split(/\s+/)[0] || '').toLowerCase();
  if (s.length < 3 || /[\[\]]/.test(s)) return null;
  if (IMIONA_MESKIE_NA_A.includes(s)) return 'M';
  if (IMIONA_ZENSKIE_BEZ_A.includes(s)) return 'K';
  if (/^puer/.test(s)) return 'M';
  if (/^puell/.test(s)) return 'K';
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
  if (klucz === 'plec') {
    if (/^(m|masc|męż|mez|mezczyzna|mężczyzna)/.test(s)) return 'M';
    if (/^(k|f|fem|kob)/.test(s)) return 'K';
  }
  if (klucz === 'typ') {
    if (TYPY[v]) return v;
    const kod = Object.keys(TYPY).find((t) => TYPY[t] === s || t === s.replace(/\s+/g, '_'));
    if (kod) return kod;
  }
  if (klucz === 'miesiac') {           // "Januarius.", "styczeń", "1" -> "1"
    const n = miesiacZ(v);
    if (n) return String(n);
  }
  return v;
}

const NAZWISKA = ['zm_nazwisko', 'zm_rodowe', 'oj_nazwisko', 'ma_rodowe', 'mz_nazwisko', 'mz_rodowe'];

const NAZWY_MIESIECY = ['styczeń', 'luty', 'marzec', 'kwiecień', 'maj', 'czerwiec', 'lipiec', 'sierpień', 'wrzesień',
  'październik', 'listopad', 'grudzień'];

/* Miesiąc dla aktów, przy których nie jest napisany: z daty "28/2" albo z najbliższego wcześniejszego
   aktu (w księdze nazwa miesiąca stoi tylko nad pierwszym aktem miesiąca). Taki miesiąc jest
   oznaczony jako automatyczny - zmienia się sam, gdy poprawisz miesiąc wcześniej. */
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
  };
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
      opis: st.opis, ai_zmieniony: st.aiZmieniony,
      wycinki: st.wycinki.map((w) => ({
        wycinek: w.wycinek, plik: w.plik, rodzaj: w.rodzaj, lacina: w.lacina, polski: w.polski,
        niepewne: w.niepewne, uwagi: w.uwagi, sprawdzony: w.sprawdzony, miesiac_auto: w.miesiacAuto,
        edytowany: w.edytowane.size > 0, edytowane_pola: [...w.edytowane], zmieniono: w.zmieniono,
      })),
    },
  };
}

/* Wpis z Twoimi poprawkami (z) + aktualny odczyt AI (a), pole po polu: pola, które sam wpisałeś
   albo zmieniłeś, zostają; pozostałe zawsze biorą wartość z odczytu AI (np. gdy AI skończyło
   czytać już po tym, jak coś wpisałeś). */
function polacz(z, a, plik, wycinek) {
  if (!z) {
    const w = normalizuj(a, plik, wycinek);
    if (a) zapamietajAi(w, w);
    return { w, zmiany: 0, zdjete: false };
  }
  const w = normalizuj(z, plik, wycinek);
  if (!a) return { w, zmiany: 0, zdjete: false };
  const ai = normalizuj(a, plik, wycinek);
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
  return n;
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
    if (mz && m && mz !== m && !w.miesiacAuto) dodaj('miesiac', `data zgonu zapisana jako ${L.dzien_zgonu}, a miesiąc ${m}`);
    const [dp, mp] = dzienMiesiac(L.dzien_pogrzebu);
    if (L.dzien_zgonu && dz === null) dodaj('dzien_zgonu', 'dzień zgonu to nie liczba');
    if (L.dzien_pogrzebu && dp === null) dodaj('dzien_pogrzebu', 'dzień pogrzebu to nie liczba');
    const mies = mz || miesiac;
    let zgon = null;
    if (dz !== null && rok && mies) {
      zgon = data(rok, mies, dz);
      if (!zgon) dodaj('dzien_zgonu', 'niemożliwa data zgonu');
    }
    if (zgon && dp !== null) {
      let pm, pr;
      if (mp) { pm = mp; pr = rok + (mp < mies ? 1 : 0); }
      else if (dp >= dz) { pm = mies; pr = rok; }
      else { pm = (mies % 12) + 1; pr = rok + (mies === 12 ? 1 : 0); }
      const pogrzeb = data(pr, pm, dp);
      if (!pogrzeb) dodaj('dzien_pogrzebu', 'niemożliwa data pogrzebu');
      else {
        const dni = Math.round((pogrzeb - zgon) / 864e5);
        if (dni > 5) dodaj('dzien_pogrzebu', `pogrzeb ${dni} dni po zgonie`);
        else if (dni === 0) dodaj('dzien_pogrzebu', 'pogrzeb tego samego dnia co zgon');
      }
    }
    if (zgon && poprzData && zgon < poprzData) dodaj('dzien_zgonu', 'data zgonu wcześniejsza niż w poprzednim akcie');
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
      if (L[k] && P[k] && L[k] !== P[k]) dodaj(k, 'nazwisko po polsku różni się od łacińskiego (nazwisk się nie tłumaczy)');
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
  st.sciezka = d.sciezka;
  st.opis = d.opis;
  st.aiZmieniony = d.ai_zmieniony;
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
  pokazKomunikaty(kom);

  $('#naglowek-folder').textContent = nazwa;
  $('#naglowek-folder').title = d.sciezka;
  $('#pusto').classList.add('ukryta');
  $('#narzedzia').classList.toggle('ukryta', !st.wycinki.length);
  $('#b-pobierz').disabled = !st.wycinki.length;
  uzupelnijMiesiace();
  obliczFlagi();
  rysujKarty();
  $$('.wezel').forEach((w) => w.classList.toggle('aktywny', w.dataset.sciezka === d.opis.rocznik));
  window.scrollTo(0, 0);
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
  // ten sam przycisk u góry i na dole karty (żeby po sprawdzeniu pól nie przewijać do góry)
  const sprawdzone = [0, 1].map(() => el('button', { class: 'b-sprawdzone' }));
  const pokazSprawdzone = () => sprawdzone.forEach((b) => {
    b.classList.toggle('wcisniety', w.sprawdzony);
    b.textContent = w.sprawdzony ? '✓ Sprawdzone' : 'Oznacz jako sprawdzone';
  });
  pokazSprawdzone();
  sprawdzone.forEach((b, i) => b.addEventListener('click', () => {
    w.sprawdzony = !w.sprawdzony;
    w.zmieniono = Date.now();
    pokazSprawdzone();
    s.classList.toggle('gotowa', w.sprawdzony);
    // sprawdzony wycinek zwija pola; cofnięcie oznaczenia znów je pokazuje
    szczegoly.open = w.sprawdzony ? false : st.otwarte;
    zaznaczFlagiKarty(s, w);
    odswiezLiczniki();
    zaplanujZapis();
    zastosujFiltr(true);
    if (w.sprawdzony && i === 1) {
      // z dolnego przycisku: po zwinięciu przejdź do następnego wycinka (inaczej strona "skacze")
      let nast = s.nextElementSibling;
      while (nast && nast.classList.contains('ukryta')) nast = nast.nextElementSibling;
      (nast || s).scrollIntoView({ block: 'start' });
    }
  }));
  const img = el('img', { loading: 'lazy', alt: `wycinek ${w.wycinek}` });
  if (WWW) WWW.obraz(img, st.sciezka, w.plik);       // zaszyfrowany na GitHubie - odszyfrowuje www.js
  else img.src = `/api/obraz?sciezka=${enc(st.sciezka)}&plik=${enc(w.plik)}&v=${w.zmieniony}`;
  const obraz = el('div', { class: 'obraz', title: 'Kliknij, żeby powiększyć' }, img);
  img.addEventListener('click', () => obraz.classList.toggle('duzy'));
  const szczegoly = el('details', { open: st.otwarte && !w.sprawdzony },
    el('summary', { text: 'Wszystkie pola (możesz poprawić każde)' }),
    el('div', { class: 'cialo' }));
  szczegoly.addEventListener('toggle', () => { if (szczegoly.open) $$('textarea', szczegoly).forEach(dopasujWysokosc); });
  s.append(
    el('h2', {},
      el('span', { text: `Wycinek ${w.wycinek}` }),
      el('span', { class: 'opis', text: opisKarty(w) }),
      el('span', { class: 'znaczek ukryta' }),
      el('span', { class: 'znaczek ok poprawki ukryta', title: 'Pola poprawione względem odczytu AI – przy każdym z nich jest ✓' }),
      el('span', { class: 'prawo' }, rodzaj, sprawdzone[0])),
    obraz, szczegoly,
    el('div', { class: 'stopka' }, sprawdzone[1]));
  zaznaczFlagiKarty(s, w);
  return s;
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

function zaznaczFlagiKarty(karta, w) {
  const f = st.flagi.get(w.wycinek) || {};
  const ile = Object.keys(f).length;
  karta.classList.toggle('ma-flagi', ile > 0 && !w.sprawdzony);
  const znaczek = $('.znaczek', karta);
  if (znaczek) {
    znaczek.classList.toggle('ukryta', ile === 0);
    znaczek.textContent = `do sprawdzenia: ${ile}`;
    znaczek.title = Object.values(f).flat().filter((v, i, a) => a.indexOf(v) === i).join('\n');
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
    obliczFlagi();
    for (const k of $$('.karta')) zaznaczFlagiKarty(k, st.mapa.get(k.dataset.wycinek));
    odswiezLiczniki();
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
      RODZAJE_POL[t.dataset.pole] !== 'tekst') {
    st.edycja = null;
    return;
  }
  const w = st.mapa.get(karta.dataset.wycinek);
  const k = t.dataset.pole;
  st.edycja = {
    el: t,
    stare: { lacina: t.dataset.jezyk === 'lacina' ? w.lacina[k] : null, polski: w.polski[k] },
    tekst: { lacina: w.lacina.tekst, polski: w.polski.tekst },
  };
}

function zamienSlowo(tekstZ, stare, nowe) {
  if (!tekstZ || !stare || !stare.trim()) return tekstZ;
  const wzor = stare.trim().replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return tekstZ.replace(new RegExp(`(?<![\\p{L}\\p{N}])${wzor}(?![\\p{L}\\p{N}])`, 'gu'), () => nowe.trim());
}

function poprawTekstAktu(w, karta, edycja, nowe, jezyki) {
  for (const j of jezyki) {
    const stare = edycja.stare[j];
    if (!stare || stare === nowe) continue;
    const poprawiony = zamienSlowo(edycja.tekst[j], stare, nowe);
    if (poprawiony === edycja.tekst[j] || poprawiony === w[j].tekst) continue;
    w[j].tekst = poprawiony;
    w.edytowane.add(`${j}.tekst`);
    const pole = karta.querySelector(`[data-pole="tekst"][data-jezyk="${j}"]`);
    if (pole) {
      pole.value = poprawiony;
      dopasujWysokosc(pole);
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
    if (st.edycja && st.edycja.el === t) poprawTekstAktu(w, karta, st.edycja, t.value, jezyki);
  }
  if (klucz === 'miesiac') w.miesiacAuto = false;
  if (t.tagName === 'TEXTAREA') dopasujWysokosc(t);
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
    ustawStanWDrzewie(d.pobrano ? 'gotowe_zmiany' : 'w_trakcie');
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
  const zFlagami = st.wycinki.filter((w) => st.flagi.has(w.wycinek) && !w.sprawdzony).length;
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
    if (st.filtr === 'flagi') widoczna = st.flagi.has(w.wycinek) || w.niepewne.length > 0;
    if (st.filtr === 'poprawione') widoczna = liczPoprawione(w) > 0;
    // karta, którą właśnie edytujesz, nie znika spod kursora
    if (tylkoKlasy && !widoczna && k.contains(document.activeElement)) continue;
    k.classList.toggle('ukryta', !widoczna);
  }
}

// ---------------------------------------------------------------- pobieranie xlsx

async function pobierz() {
  if (!st.sciezka) return;
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
  } catch (e) {
    toast('Nie udało się zrobić pliku xlsx: ' + e.message);
  } finally {
    przycisk.disabled = false;
  }
}

// ---------------------------------------------------------------- okna: prompt, wgraj

function okno(tytul, cialo, przyciski) {
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
  POLA = d.pola;
  TYPY = d.typy;
  RODZAJE = d.rodzaje;
  for (const [k, , r] of POLA) RODZAJE_POL[k] = r;
  try { st.otwarte = localStorage.getItem('zwiniete') !== '1'; } catch (e) { /* bez znaczenia */ }
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
  $('#b-zwin').addEventListener('click', () => {
    st.otwarte = !st.otwarte;
    try { localStorage.setItem('zwiniete', st.otwarte ? '0' : '1'); } catch (e) { /* bez znaczenia */ }
    $('#b-zwin').textContent = st.otwarte ? 'Zwiń wszystkie pola' : 'Rozwiń wszystkie pola';
    $$('.karta details').forEach((det) => { det.open = st.otwarte; });
  });
  $('#okno').addEventListener('click', (e) => { if (e.target.id === 'okno') zamknijOkno(); });
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape') zamknijOkno(); });
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
