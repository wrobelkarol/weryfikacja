// Weryfikacja Geneteka – sprawdzanie rocznika przerobionego z Weryfikacji na format Geneteki.
import * as K from './konwersja.js';
import { zbudujXlsx, czytajZgony, czytajSlownikImion } from './xlsx.js';
import { DyskSerwer, DyskFolder, DyskGitHub, folderDostepny, pobierzPlik } from './dysk.js';

const JSZip = window.JSZip;
const PLIK_POPRAWEK = 'genetyka_poprawki.json';
const PLIK_ZWERYFIKOWANY = 'odczyt_zweryfikowany.json';
const FOLDER_WYCINKOW = 'pocięte';
const XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
// początek nazwy pliku dla Geneteki: parafia (miejscowość i wezwanie) + rodzaj rejestru
const NAZWY_PARAFII = { Przytkowice: 'Przytkowice-Z-przenajswietszej-trojcy' };

// pola z Weryfikacji (jak w weryfikacja.pyw, lista POLA)
const POLA = [
  ['nr', 'Numer aktu'], ['miesiac', 'Miesiąc zgonu'], ['dzien_zgonu', 'Dzień zgonu'], ['dzien_pogrzebu', 'Dzień pogrzebu'],
  ['miesiac_pogrzebu', 'Miesiąc pogrzebu'], ['dom', 'Numer domu'], ['katolik', 'Katolik'], ['plec', 'Płeć'], ['wiek', 'Wiek'],
  ['przyczyna', 'Przyczyna zgonu'], ['typ', 'Rodzaj aktu'], ['zm_imie', 'Zmarły – imię'], ['zm_nazwisko', 'Zmarły – nazwisko'],
  ['zm_rodowe', 'Zmarły – nazwisko rodowe'], ['zm_stan', 'Zmarły – stan/zawód'], ['oj_imie', 'Ojciec – imię'],
  ['oj_nazwisko', 'Ojciec – nazwisko'], ['oj_stan', 'Ojciec – stan/zawód'], ['ma_imie', 'Matka – imię'],
  ['ma_rodowe', 'Matka – nazwisko rodowe'], ['ma_stan', 'Matka – stan/zawód'], ['mz_imie', 'Małżonek – imię'],
  ['mz_nazwisko', 'Małżonek – nazwisko'], ['mz_rodowe', 'Małżonek – nazwisko rodowe'], ['mz_stan', 'Małżonek – stan/zawód'],
  ['sakramenty', 'Zaopatrzony sakramentami'], ['czynnosc', 'Ksiądz pochował / zapisał'], ['ksiadz', 'Ksiądz'],
];
const DLUGIE_POLA = [['dodatkowe', 'Inne informacje z aktu'], ['adnotacje', 'Adnotacje / dopiski'], ['tekst', 'Tekst aktu (cały)']];
const TYPY = { zona: 'żona', wdowa: 'wdowa', maz: 'mąż', wdowiec: 'wdowiec', dziecko: 'dziecko', dziecko_nieslubne: 'dziecko nieślubne',
  martwe_urodzenie: 'martwe urodzenie', inny: 'inny' };
// układ pól Geneteki na stronie (szerokie = cały wiersz)
const SZEROKIE = new Set(['M']);
const SREDNIE = new Set(['L', 'H', 'I']);

const st = {
  dysk: null, roczniki: [], r: null, dane: null, wyniki: [], popr: null, i: 0,
  szablon: null, slownik: null, brudne: false, zapisTimer: null, zapisywanie: null,
};

const $ = (s) => document.querySelector(s);
function el(tag, atr = {}, ...dzieci) {
  const e = document.createElement(tag);
  for (const [k, v] of Object.entries(atr)) {
    if (v === undefined || v === null || v === false) continue;
    if (k === 'text') e.textContent = v;
    else if (k === 'class') e.className = v;
    else if (k.startsWith('on')) e.addEventListener(k.slice(2), v);
    else if (k in e && typeof v !== 'string') e[k] = v;
    else e.setAttribute(k, v === true ? '' : v);
  }
  for (const d of dzieci.flat()) if (d !== null && d !== undefined && d !== false) e.append(d instanceof Node ? d : String(d));
  return e;
}
function teraz() {
  const d = new Date(), z = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${z(d.getMonth() + 1)}-${z(d.getDate())} ${z(d.getHours())}:${z(d.getMinutes())}:${z(d.getSeconds())}`;
}
function pamiec(klucz, wartosc) {
  try {
    if (wartosc === undefined) return localStorage.getItem('genetyka:' + klucz);
    localStorage.setItem('genetyka:' + klucz, wartosc);
  } catch { /* przeglądarka bez pamięci – nic się nie dzieje */ }
  return null;
}

// ------------------------------------------------------------------ komunikaty, okno, toast

function toast(tekst, ms = 2600) {
  const t = $('#toast');
  t.textContent = tekst;
  t.classList.remove('ukryta');
  clearTimeout(toast.timer);
  toast.timer = setTimeout(() => t.classList.add('ukryta'), ms);
}
function komunikat(tresc, rodzaj = '', przyciski = []) {
  const k = el('div', { class: `komunikat ${rodzaj}` }, el('div', {}, tresc), przyciski.length ? el('div', { class: 'przyciski' }, przyciski) : null);
  $('#komunikaty').append(k);
  return k;
}
function okno(tytul, cialo, przyciski) {
  return new Promise((ok) => {
    $('#okno-tytul').textContent = tytul;
    $('#okno-cialo').replaceChildren(...[cialo].flat());
    const zamknij = (w) => { $('#okno').classList.add('ukryta'); document.removeEventListener('keydown', esc, true); ok(w); };
    const esc = (e) => { if (e.key === 'Escape') { e.stopPropagation(); zamknij(null); } };
    document.addEventListener('keydown', esc, true);
    $('#okno-przyciski').replaceChildren(...przyciski.map(([tekst, wartosc, glowny]) =>
      el('button', { text: tekst, class: glowny ? 'glowny' : '', onclick: () => zamknij(wartosc) })));
    $('#okno').classList.remove('ukryta');
    $('#okno-przyciski button.glowny')?.focus();
  });
}
function ustawStanZapisu(tekst, blad = false) {
  const s = $('#stan-zapisu');
  s.textContent = tekst;
  s.classList.toggle('blad', blad);
}

// ------------------------------------------------------------------ start: dostęp do dysku

async function start() {
  try {
    st.szablon = await (await fetch('szablon/indeksy.xlsx')).arrayBuffer();
    st.slownik = K.zbudujSlownik(await czytajSlownikImion(JSZip, st.szablon));
  } catch (e) {
    komunikat(`Nie udało się wczytać szablonu Geneteki (szablon/indeksy.xlsx): ${e.message}`, 'blad');
    return;
  }
  // strona w internecie (geneteka/ na stronie Weryfikacji): dane z GitHuba, logowanie jak na stronie Weryfikacji
  if (window.WWW) {
    st.dysk = await DyskGitHub.wykryj();
    $('#b-weryfikacja')?.classList.remove('ukryta');
    return wczytajRoczniki();
  }
  st.dysk = await DyskSerwer.wykryj();
  if (st.dysk) {
    setInterval(() => st.dysk.ping().catch(() => ustawStanZapisu('Program zamknięty – uruchom go ponownie skrótem z pulpitu.', true)), 20000);
    return wczytajRoczniki();
  }
  const zap = await DyskFolder.zapamietany();
  if (zap?.uprawnienie === 'granted') { st.dysk = zap; return wczytajRoczniki(); }
  pokazPolaczenie(zap);
}

function pokazPolaczenie(zapamietany) {
  const p = $('#polaczenie');
  if (!folderDostepny()) {
    p.replaceChildren(
      el('p', {}, el('b', { text: 'Ta przeglądarka nie daje stronie dostępu do folderu na dysku.' })),
      el('p', { text: 'Najprościej: otwórz program skrótem „Weryfikacja Geneteka” z pulpitu (działa bez żadnych ustawień).' }),
      el('p', {}, 'W Brave można też włączyć dostęp: wpisz w pasku adresu ', el('code', { text: 'brave://flags/#file-system-access-api' }),
        ', ustaw ', el('b', { text: 'Enabled' }), ' i uruchom Brave ponownie.'));
    return;
  }
  const przyciski = [];
  if (zapamietany) {
    przyciski.push(el('button', { class: 'glowny', text: `Połącz z folderem „${zapamietany.nazwa}”`, onclick: async () => {
      if (await zapamietany.polacz()) { st.dysk = zapamietany; wczytajRoczniki(); }
    } }));
  }
  przyciski.push(el('button', { class: zapamietany ? '' : 'glowny', text: 'Wybierz folder „SKANY MAM DOŚĆ”…', onclick: wybierzFolder }));
  p.replaceChildren(el('p', { text: 'Wskaż folder „SKANY MAM DOŚĆ” (raz – przeglądarka go zapamięta). Pliki zostają na Twoim dysku.' }),
    el('div', { class: 'boczne-przyciski' }, przyciski));
}

async function wybierzFolder() {
  try {
    await zapiszTeraz();
    st.dysk = await DyskFolder.wybierz();
    st.r = null;
    await wczytajRoczniki();
  } catch (e) {
    if (e.name !== 'AbortError') toast(e.message);
  }
}

// ------------------------------------------------------------------ roczniki

function nfc(s) { return String(s).normalize('NFC'); }

async function wczytajRoczniki() {
  $('#polaczenie').replaceChildren(el('p', { text: 'Szukam roczników…' }));
  const roczniki = [];
  const parafie = (await st.dysk.lista(''))?.foldery || [];
  for (const parafia of parafie) {
    const ksiegi = ((await st.dysk.lista(parafia))?.foldery || []).filter((k) => /^zgony$/i.test(nfc(k)));
    for (const ksiega of ksiegi) {
      const lata = ((await st.dysk.lista(`${parafia}/${ksiega}`))?.foldery || []).filter((r) => /^\d{4}$/.test(r));
      await Promise.all(lata.map(async (rok) => {
        const sciezka = `${parafia}/${ksiega}/${rok}`;
        const zaw = await st.dysk.lista(sciezka);
        const pociete = zaw?.foldery.find((f) => nfc(f) === FOLDER_WYCINKOW);
        if (!pociete) return;
        let dane = null;
        try { dane = JSON.parse(((await st.dysk.czytaj(`${sciezka}/${pociete}/${PLIK_ZWERYFIKOWANY}`)) || 'null').replace(/^﻿/, '')); } catch { dane = null; }
        if (!dane?.pobrano) return;                  // tylko roczniki, dla których w Weryfikacji pobrano już xlsx
        if (dane.xlsx_geneteki) return;              // ... ale nie te, które Weryfikacja zapisała od razu w formacie Geneteki
        roczniki.push({ parafia, ksiega, rok: +rok, sciezka, pociete, pliki: zaw.pliki, dane,
          akty: (dane.wycinki || []).filter((w) => (w.rodzaj || 'akt') === 'akt').length });
      }));
    }
  }
  roczniki.sort((a, b) => a.parafia.localeCompare(b.parafia) || a.rok - b.rok);
  st.roczniki = roczniki;
  const sel = $('#rocznik');
  sel.replaceChildren(...roczniki.map((r, i) => el('option', { value: String(i), text: `${r.parafia} · ${r.ksiega} · ${r.rok}` })));
  sel.disabled = !roczniki.length;
  $('#b-scal').disabled = !roczniki.length;
  $('#b-odswiez').disabled = false;
  $('#b-folder').classList.toggle('ukryta', st.dysk.rodzaj !== 'folder');
  if (!roczniki.length) {
    $('#widok').classList.add('ukryta');
    $('#start').classList.remove('ukryta');
    const wewnatrz = /^\d{4}$|^zgony$|^urodzenia$|^małżeństwa$|^pocięte$/i.test(nfc(st.dysk.nazwa)) || parafie.some((p) => /^\d{4}$|^pocięte$/.test(nfc(p)));
    $('#polaczenie').replaceChildren(
      el('p', {}, el('b', { text: `W folderze „${st.dysk.nazwa}” nie ma roczników zgonów pobranych z Weryfikacji.` })),
      wewnatrz || st.dysk.rodzaj === 'folder'
        ? el('p', { text: 'Wskaż folder „SKANY MAM DOŚĆ” – ten, w którym jest folder „Przytkowice” (nie rocznik ani folder „Zgony”).' }) : null,
      st.dysk.rodzaj === 'folder' ? el('button', { class: 'glowny', text: 'Wybierz inny folder…', onclick: wybierzFolder }) : null);
    return;
  }
  $('#polaczenie').replaceChildren();
  const ostatni = pamiec('rocznik');
  const i = Math.max(0, roczniki.findIndex((r) => `${r.sciezka}` === ostatni));
  sel.value = String(i);
  await otworzRocznik(roczniki[i]);
}

async function otworzRocznik(r) {
  await zapiszTeraz();
  st.r = r;
  st.dane = r.dane;
  pamiec('rocznik', r.sciezka);
  ustawStanZapisu('');
  $('#komunikaty').replaceChildren();
  let popr = null;
  try { popr = JSON.parse((await st.dysk.czytaj(`${r.sciezka}/${PLIK_POPRAWEK}`)) || 'null'); } catch (e) {
    komunikat(`Plik ${PLIK_POPRAWEK} jest uszkodzony (${e.message}) – zaczynam od nowa, stary plik zostanie nadpisany przy pierwszej zmianie.`, 'blad');
  }
  st.popr = popr && typeof popr === 'object' ? popr : {};
  st.popr.program ||= 'genetyka-weryfikacja';
  st.popr.rok = r.rok;
  st.popr.akty ||= {};
  // z którego zdjęcia jest każdy wycinek (Tartak) – nr skanu w uwagach, gdy rocznik nie ma zakresów linków
  st.skany = null;
  try { st.skany = JSON.parse((await st.dysk.czytaj(`${r.sciezka}/${r.pociete}/skany.json`)) || 'null')?.wycinki || null; } catch { st.skany = null; }
  przelicz();
  if (r.dane.zmiany_po_pobraniu) {
    komunikat('W Weryfikacji są zmiany zrobione po ostatnim pobraniu xlsx – pola poniżej już je uwzględniają.', 'uwaga');
  }
  const link = $('#link-rocznika');
  link.disabled = false;
  link.value = biezacyLink();
  odswiezPtaszekLinku();
  $('#b-xlsx').disabled = false;
  $('#start').classList.add('ukryta');
  $('#widok').classList.remove('ukryta');
  const zapamietany = st.wyniki.findIndex((w) => w.wycinek === String(pamiec(`ostatni:${r.sciezka}`) ?? ''));
  pokazAkt(zapamietany >= 0 ? zapamietany : 0, false);
}

function biezacyLink() {
  return st.popr.link ?? K.linkRocznika(st.dane);
}

function przelicz() {
  st.wyniki = K.konwertujRocznik(st.dane, { rok: st.r.rok, slownik: st.slownik, miejscowosc: st.r.parafia, link: biezacyLink(),
    skany: st.skany });
}

// ------------------------------------------------------------------ wiersze aktu (automatyczne + poprawki)

function aktPopr(wycinek, utworz = false) {
  let a = st.popr.akty[wycinek];
  if (!a && utworz) a = st.popr.akty[wycinek] = { wiersze: [{ poprawki: {} }] };
  if (a && !a.wiersze?.length) a.wiersze = [{ poprawki: {} }];
  return a;
}

/** Wiersze aktu: [{wartosci, auto, poprawki, flagi(kolumna -> [powody]), zmiany(kolumna -> poprawka ze zmienionym źródłem)}] */
function wierszeAktu(i) {
  const wynik = st.wyniki[i];
  const a = aktPopr(wynik.wycinek);
  const lista = a?.wiersze?.length ? a.wiersze : [{ poprawki: {} }];
  return lista.map((w, nr) => {
    const poprawki = w.poprawki || {};
    const wartosci = {}, flagi = {}, zmiany = {};
    for (const k of K.KOLUMNY) {
      const p = poprawki[k];
      wartosci[k] = p ? p.wartosc : wynik.wiersz[k];
      const powody = [];
      if (!p) powody.push(...(wynik.flagi[k] || []));
      if (p && p.auto !== wynik.wiersz[k]) {
        zmiany[k] = p;
        powody.push(`w Weryfikacji zmieniło się: było „${p.auto || '(puste)'}”, teraz „${wynik.wiersz[k] || '(puste)'}” – zostaje Twoja poprawka`);
      }
      if (k !== 'A' && K.WYMAGANE.includes(k) && !String(wartosci[k] || '').trim()) {
        if (!powody.some((x) => /wymagane/.test(x))) powody.push('pole wymagane przez Genetekę jest puste');
      }
      if (k !== 'L' && p && /[\\()'"*;]|\.\.\.|…|\.,|[½¼¾]/.test(p.wartosc)) powody.push('znaki niedozwolone w Genetece (zostaną usunięte przy zapisie)');
      if (powody.length) flagi[k] = powody;
    }
    if (nr > 0) wartosci.A = '';
    return { nr, wartosci, auto: wynik.wiersz, poprawki, flagi, zmiany };
  });
}

function stanAktu(i) {
  const wynik = st.wyniki[i];
  const a = aktPopr(wynik.wycinek);
  const wiersze = wierszeAktu(i);
  return {
    sprawdzony: !!a?.sprawdzony,
    zmieniony: wiersze.some((w) => Object.keys(w.poprawki).length) || wiersze.length > 1,
    flagi: wiersze.some((w) => Object.keys(w.flagi).length),
    zmianyZrodla: wiersze.some((w) => Object.keys(w.zmiany).length),
  };
}

// ------------------------------------------------------------------ lista aktów

function rysujListe() {
  const ul = $('#lista-aktow');
  ul.replaceChildren(...st.wyniki.map((w, i) => {
    const s = stanAktu(i);
    const ikona = s.zmianyZrodla ? el('span', { class: 'ikona stan-flagi', text: '●', title: 'w Weryfikacji zmieniono dane po Twojej poprawce' })
      : s.sprawdzony ? el('span', { class: 'ikona stan-sprawdzony', text: '✓', title: 'sprawdzony' })
      : s.flagi ? el('span', { class: 'ikona stan-flagi', text: '●', title: 'do sprawdzenia' })
        : el('span', { class: 'ikona stan-nowy', text: '○', title: 'jeszcze nie sprawdzony' });
    const nazwisko = `${w.wiersz.E} ${w.wiersz.H}`.trim();
    return el('li', { class: i === st.i ? 'aktywny' : '', onclick: () => pokazAkt(i), title: nazwisko },
      ikona, el('span', { class: 'nazwa' }, `Akt ${w.wiersz.C || '?'}`, ' ', el('span', { class: 'kto', text: nazwisko })),
      s.zmieniony ? el('span', { class: 'zmieniony', text: '✎', title: 'zmieniony ręcznie' }) : el('span', { class: 'zmieniony' }));
  }));
  const sprawdzone = st.wyniki.filter((_, i) => stanAktu(i).sprawdzony).length;
  $('#licznik').replaceChildren('sprawdzone ', el('b', { text: String(sprawdzone) }), ` / ${st.wyniki.length}`);
  $('#pasek').style.width = st.wyniki.length ? `${(100 * sprawdzone) / st.wyniki.length}%` : '0';
  ul.querySelector('li.aktywny')?.scrollIntoView({ block: 'nearest' });
}

// ------------------------------------------------------------------ jeden akt

async function pokazAkt(i, przewin = true) {
  if (!st.wyniki.length) return;
  st.i = Math.max(0, Math.min(i, st.wyniki.length - 1));
  const wynik = st.wyniki[st.i];
  const wycinek = st.dane.wycinki.find((w) => String(w.wycinek) === wynik.wycinek);
  pamiec(`ostatni:${st.r.sciezka}`, wynik.wycinek);
  $('#tytul-aktu').textContent = `${wynik.wycinek} z ${st.wyniki.length}`;
  $('#opis-aktu').replaceChildren(`akt nr ${wynik.wiersz.C || '?'} · ${st.r.rok}`,
    ...(stanAktu(st.i).sprawdzony ? [el('span', { class: 'znaczek ok', text: 'sprawdzony' })] : []));
  const img = $('#obraz-img');
  img.src = await st.dysk.adresObrazu(`${st.r.sciezka}/${st.r.pociete}/${wycinek.plik || `${wynik.wycinek}.jpg`}`);
  rysujWeryfikacje(wycinek);
  rysujGeneteke();
  rysujListe();
  odswiezPrzyciski();
  // po przejściu do aktu pokazujemy całą część „Format Geneteki” (data, imię ojca i matki, nazwisko matki…),
  // a nie górę strony – wycinek i pola z Weryfikacji są kawałek wyżej
  if (przewin) $('#sekcja-geneteki').scrollIntoView({ block: 'start' });
}

function rysujWeryfikacje(wycinek) {
  const p = wycinek.polski || {};
  // miesiąc wpisany wyżej obowiązuje dla kolejnych aktów – pokazujemy go kursywą
  let miesiac = null;
  for (const w of st.dane.wycinki) {
    if ((w.rodzaj || 'akt') !== 'akt') continue;
    if (w.polski?.miesiac) miesiac = w.polski.miesiac;
    if (w === wycinek) break;
  }
  const pola = POLA.map(([k, etykieta]) => {
    let v = p[k] ?? '', auto = false;
    if (k === 'miesiac' && !v && miesiac) { v = miesiac; auto = true; }
    if (k === 'typ') v = TYPY[v] || v;
    if (String(v).trim() === '') return null;
    return el('div', { class: 'pole' }, el('span', { class: 'etykieta', text: etykieta }),
      el('span', { class: `wartosc${auto ? ' auto' : ''}`, text: v, title: auto ? 'miesiąc z poprzedniego aktu' : null }));
  }).filter(Boolean);
  $('#pola-weryfikacji').replaceChildren(...pola);
  const dlugie = DLUGIE_POLA.map(([k, etykieta]) => (String(p[k] || '').trim()
    ? el('div', { class: `pole${k === 'dodatkowe' ? ' wazne' : ''}` }, el('span', { class: 'etykieta', text: etykieta }), el('span', { class: 'wartosc', text: p[k] }))
    : null)).filter(Boolean);
  if (String(wycinek.uwagi || '').trim()) {
    dlugie.unshift(el('div', { class: 'pole wazne' }, el('span', { class: 'etykieta', text: 'Uwagi do wycinka (z Weryfikacji)' }),
      el('span', { class: 'wartosc', text: wycinek.uwagi })));
  }
  $('#dlugie-weryfikacji').replaceChildren(...dlugie);
}

function rysujGeneteke() {
  const wynik = st.wyniki[st.i];
  const wiersze = wierszeAktu(st.i);
  const kontener = $('#wiersze-geneteki');
  kontener.replaceChildren(...wiersze.map((w) => rysujWiersz(wynik, w, wiersze.length)));
}

function rysujWiersz(wynik, w, ile) {
  const blok = el('div', { class: 'wiersz-geneteki', 'data-wiersz': String(w.nr) });
  if (ile > 1) {
    blok.append(el('h3', {}, `Wiersz ${w.nr + 1}`, w.nr > 0
      ? el('button', { text: 'Usuń ten wiersz', onclick: () => usunWiersz(w.nr) }) : el('span', { class: 'opis', text: '(główny)' })));
  }
  const siatka = el('div', { class: 'siatka-geneteki' });
  for (const k of K.KOLUMNY) {
    const etykieta = el('span', { class: 'etykieta' }, el('span', { class: 'litera', text: k }), ' ',
      K.WYMAGANE.includes(k) ? el('b', { text: K.NAGLOWKI[k] }) : K.NAGLOWKI[k]);
    let pole;
    if (k === 'A') {
      pole = el('input', { value: w.nr > 0 ? '(kolejny)' : 'automatycznie', readOnly: true, class: 'tylko-odczyt', tabIndex: -1,
        title: 'L.p. jest numerowane automatycznie przy zapisie xlsx' });
    } else if (SZEROKIE.has(k)) {
      pole = el('textarea', { rows: 2, spellcheck: false });
      pole.value = w.wartosci[k] ?? '';
    } else {
      pole = el('input', { value: w.wartosci[k] ?? '', spellcheck: false });
    }
    pole.dataset.kol = k;
    pole.setAttribute('aria-label', `${k} ${K.NAGLOWKI[k]}`);
    const ptaszek = el('span', { class: 'ptaszek', text: '✓', role: 'button', tabIndex: -1 });
    ptaszek.addEventListener('click', () => przywroc(w.nr, k));
    const opak = el('div', { class: `pole${SZEROKIE.has(k) ? ' szerokie' : SREDNIE.has(k) ? ' sredni' : ''}`, 'data-kol': k },
      etykieta, el('span', { class: 'z-ptaszkiem' }, pole, ptaszek));
    if (k !== 'A') {
      pole.addEventListener('input', () => zmienPole(w.nr, k, pole.value));
      if (pole.tagName === 'TEXTAREA') {
        const dopasuj = () => { pole.style.height = 'auto'; pole.style.height = `${pole.scrollHeight + 2}px`; };
        pole.addEventListener('input', dopasuj);
        requestAnimationFrame(dopasuj);
      }
    }
    siatka.append(opak);
  }
  blok.append(siatka);
  blok.append(el('ul', { class: 'powody ukryta' }));
  if (w.nr === 0 && wynik.podpowiedzi.length) {
    blok.append(el('div', { class: 'dopiski' }, wynik.podpowiedzi.map((t) => el('div', { class: 'dopisek' },
      el('b', { text: 'Inne informacje z aktu:' }), el('span', { text: t }),
      el('button', { text: 'Dopisz do uwag', title: 'Dopisuje tekst na końcu kolumny „uwagi” – potem zostaw tylko to, co jest w oryginale',
        onclick: () => dopiszDoUwag(w.nr, t) })))));
  }
  odswiezStanWiersza(blok, w);
  return blok;
}

/** Kolory, ptaszki i lista powodów „do sprawdzenia” bez przerysowywania pól (fokus zostaje w polu). */
function odswiezStanWiersza(blok, w) {
  for (const k of K.KOLUMNY) {
    const opak = blok.querySelector(`.pole[data-kol="${k}"]`);
    if (!opak) continue;
    const pole = opak.querySelector('input, textarea');
    const flagi = w.flagi[k];
    pole.classList.toggle('do-sprawdzenia', !!flagi);
    opak.classList.toggle('ma-flage', !!flagi);
    pole.title = flagi ? flagi.join('\n') : '';
    const p = w.poprawki[k];
    const ptaszek = opak.querySelector('.ptaszek');
    ptaszek.classList.toggle('widoczny', !!p);
    ptaszek.title = p ? `Zmieniono ręcznie. Automatycznie: „${w.auto[k] || '(puste)'}”. Kliknij, żeby przywrócić.` : '';
  }
  const ul = blok.querySelector('.powody');
  const pozycje = [];
  for (const k of K.KOLUMNY) {
    for (const powod of w.flagi[k] || []) {
      const li = el('li', {}, el('b', { text: `${k} (${K.NAGLOWKI[k]}): ` }), powod);
      if (w.zmiany[k] && powod.startsWith('w Weryfikacji')) {
        li.append(el('button', { text: 'Zostaw moją', onclick: () => potwierdzPoprawke(w.nr, k) }),
          el('button', { text: 'Weź nową', onclick: () => przywroc(w.nr, k) }));
      }
      pozycje.push(li);
    }
  }
  ul.replaceChildren(...pozycje);
  ul.classList.toggle('ukryta', !pozycje.length);
}

function odswiezBiezacy() {
  const wiersze = wierszeAktu(st.i);
  for (const w of wiersze) {
    const blok = document.querySelector(`.wiersz-geneteki[data-wiersz="${w.nr}"]`);
    if (blok) odswiezStanWiersza(blok, w);
  }
  rysujListe();
}

// ------------------------------------------------------------------ edycja

function zmienPole(nrWiersza, k, wartosc) {
  const wynik = st.wyniki[st.i];
  const a = aktPopr(wynik.wycinek, true);
  const w = a.wiersze[nrWiersza] ||= { poprawki: {} };
  w.poprawki ||= {};
  if (wartosc === wynik.wiersz[k]) delete w.poprawki[k];
  else w.poprawki[k] = { wartosc, auto: wynik.wiersz[k] };
  a.zmieniono = teraz();
  zaplanujZapis();
  odswiezBiezacy();
}

function przywroc(nrWiersza, k) {
  const wynik = st.wyniki[st.i];
  const a = aktPopr(wynik.wycinek, true);
  const w = a.wiersze[nrWiersza];
  if (!w?.poprawki?.[k]) return;
  delete w.poprawki[k];
  a.zmieniono = teraz();
  zaplanujZapis();
  rysujGeneteke();
  rysujListe();
}

function potwierdzPoprawke(nrWiersza, k) {
  const wynik = st.wyniki[st.i];
  const w = aktPopr(wynik.wycinek, true).wiersze[nrWiersza];
  if (w?.poprawki?.[k]) w.poprawki[k].auto = wynik.wiersz[k];
  zaplanujZapis();
  odswiezBiezacy();
}

function dopiszDoUwag(nrWiersza, tekst) {
  const pole = document.querySelector(`.wiersz-geneteki[data-wiersz="${nrWiersza}"] textarea[data-kol="M"]`);
  const dopisek = K.czysc(K.ulamkiNaDziesietne(tekst)).replace(/\.\s*$/, '');
  pole.value = pole.value ? `${pole.value.replace(/[\s,]+$/, '')}, ${dopisek}` : dopisek;
  pole.dispatchEvent(new Event('input'));
  pole.focus();
}

function duplikujWiersz() {
  const wynik = st.wyniki[st.i];
  const a = aktPopr(wynik.wycinek, true);
  // drugi wiersz zaczyna jako kopia pierwszego (z jego poprawkami)
  a.wiersze.push({ poprawki: JSON.parse(JSON.stringify(a.wiersze[0].poprawki || {})) });
  a.zmieniono = teraz();
  zaplanujZapis();
  rysujGeneteke();
  rysujListe();
  toast('Dodano drugi wiersz z tym samym numerem aktu – popraw imię i uwagi.');
  document.querySelector(`.wiersz-geneteki[data-wiersz="${a.wiersze.length - 1}"]`)?.scrollIntoView({ behavior: 'smooth', block: 'center' });
}

async function usunWiersz(nr) {
  const tak = await okno('Usunąć wiersz?', el('p', { text: `Wiersz ${nr + 1} tego aktu zostanie usunięty.` }),
    [['Anuluj', false], ['Usuń', true, true]]);
  if (!tak) return;
  const a = aktPopr(st.wyniki[st.i].wycinek, true);
  a.wiersze.splice(nr, 1);
  a.zmieniono = teraz();
  zaplanujZapis();
  rysujGeneteke();
  rysujListe();
}

function oznaczSprawdzony(wartosc) {
  const a = aktPopr(st.wyniki[st.i].wycinek, true);
  if (wartosc) a.sprawdzony = teraz(); else delete a.sprawdzony;
  zaplanujZapis();
}

function odswiezPrzyciski() {
  $('#b-poprzedni').disabled = st.i <= 0;
  $('#b-nastepny').disabled = st.i >= st.wyniki.length - 1;
  const spr = stanAktu(st.i).sprawdzony;
  const b = $('#b-sprawdzone');
  b.classList.toggle('wcisniety', spr);
  b.replaceChildren(spr ? 'Sprawdzony ✓ (kliknij, żeby cofnąć)' : 'Sprawdzone ✓', el('kbd', { text: 'Enter' }));
}

function sprawdzoneIDalej() {
  oznaczSprawdzony(true);
  if (st.i < st.wyniki.length - 1) pokazAkt(st.i + 1);
  else {
    pokazAkt(st.i, false);
    const brak = st.wyniki.filter((_, i) => !stanAktu(i).sprawdzony).length;
    toast(brak ? `To ostatni akt. Niesprawdzonych zostało: ${brak}.` : 'Wszystkie akty sprawdzone – możesz pobrać xlsx.');
  }
}

// ------------------------------------------------------------------ zapis poprawek

function zaplanujZapis(pokaz = true) {
  st.brudne = true;
  if (pokaz) ustawStanZapisu('Zapisywanie…');
  clearTimeout(st.zapisTimer);
  st.zapisTimer = setTimeout(zapiszTeraz, 700);
}

async function zapiszTeraz() {
  clearTimeout(st.zapisTimer);
  if (st.zapisywanie) await st.zapisywanie;
  if (!st.brudne || !st.r) return;
  st.brudne = false;
  st.popr.zapisano = teraz();
  const tresc = JSON.stringify(st.popr, null, 1);
  const sciezka = `${st.r.sciezka}/${PLIK_POPRAWEK}`;
  st.zapisywanie = (async () => {
    try {
      await st.dysk.zapisz(sciezka, new TextEncoder().encode(tresc));
      ustawStanZapisu(`Zapisano ${st.popr.zapisano.slice(11, 16)}`);
    } catch (e) {
      st.brudne = true;
      ustawStanZapisu(`Nie zapisano poprawek: ${e.message}`, true);
    }
  })();
  await st.zapisywanie;
  st.zapisywanie = null;
}

// ------------------------------------------------------------------ xlsx rocznika

function nazwaPliku(r, lata) {
  const poczatek = NAZWY_PARAFII[r.parafia] || `${r.parafia}-Z`;
  return `${poczatek}-${lata}.xlsx`;
}

function wszystkieWiersze() {
  const wiersze = [];
  st.wyniki.forEach((_, i) => wierszeAktu(i).forEach((w) => wiersze.push(w.wartosci)));
  return wiersze;
}

async function pobierzXlsx() {
  await zapiszTeraz();
  const niesprawdzone = st.wyniki.map((w, i) => (stanAktu(i).sprawdzony ? null : w.wiersz.C || w.wycinek)).filter(Boolean);
  const puste = [];
  st.wyniki.forEach((_, i) => wierszeAktu(i).forEach((w) => {
    const brak = K.WYMAGANE.filter((k) => !String(w.wartosci[k] || '').trim());
    if (brak.length) puste.push(`akt ${w.wartosci.C || '?'}: ${brak.map((k) => K.NAGLOWKI[k]).join(', ')}`);
  }));
  const ostrzezenia = [];
  if (niesprawdzone.length) ostrzezenia.push(`Niesprawdzone akty (${niesprawdzone.length}): ${niesprawdzone.slice(0, 30).join(', ')}${niesprawdzone.length > 30 ? '…' : ''}`);
  if (puste.length) ostrzezenia.push(`Puste pola wymagane: ${puste.slice(0, 10).join('; ')}${puste.length > 10 ? '…' : ''}`);
  if (!biezacyLink()) ostrzezenia.push('Brak linku do jednostki w Skanotece (kolumna „plik”) – wpisz go po lewej.');
  const nazwa = nazwaPliku(st.r, st.r.rok);
  if (ostrzezenia.length) {
    const dalej = await okno('Zapisać xlsx mimo uwag?', [el('ul', {}, ostrzezenia.map((o) => el('li', { text: o })))],
      [['Wróć', false], ['Zapisz mimo to', true, true]]);
    if (!dalej) return;
  }
  try {
    const plik = await zbudujXlsx(JSZip, st.szablon, wszystkieWiersze());
    await zapiszPlikXlsx(`${st.r.sciezka}/${nazwa}`, nazwa, plik);
    st.popr.xlsx_zapisano = teraz();
    st.brudne = true;
    await zapiszTeraz();
    st.r.pliki = [...new Set([...(st.r.pliki || []), nazwa])];
  } catch (e) {
    komunikat(`Nie udało się zapisać xlsx: ${e.message}`, 'blad');
  }
}

async function zapiszPlikXlsx(sciezka, nazwa, plik) {
  try {
    const wynik = await st.dysk.zapisz(sciezka, plik);
    toast(wynik?.komunikat || `Zapisano ${nazwa} w folderze ${sciezka.split('/').slice(0, -1).join('/')}`, 4500);
  } catch (e) {
    pobierzPlik(nazwa, plik, XLSX);
    komunikat(`Nie udało się zapisać w folderze (${e.message}) – plik ${nazwa} został pobrany do folderu pobranych plików.`, 'uwaga');
  }
}

// ------------------------------------------------------------------ scalanie lat

async function scalLata() {
  await zapiszTeraz();
  const r0 = st.r || st.roczniki[0];
  const roczniki = st.roczniki.filter((r) => r.parafia === r0.parafia && r.ksiega === r0.ksiega);
  const pozycje = [];
  for (const r of roczniki) {
    const nazwa = nazwaPliku(r, r.rok);
    const lista = await st.dysk.lista(r.sciezka);
    r.pliki = lista?.pliki || r.pliki;
    const jest = (r.pliki || []).some((p) => nfc(p) === nfc(nazwa));
    let popr = null;
    try { popr = JSON.parse((await st.dysk.czytaj(`${r.sciezka}/${PLIK_POPRAWEK}`)) || 'null'); } catch { popr = null; }
    const sprawdzone = Object.values(popr?.akty || {}).filter((a) => a.sprawdzony).length;
    const info = [];
    if (!jest) info.push('brak xlsx Geneteki');
    else {
      info.push(`sprawdzone ${sprawdzone}/${r.akty}`);
      if (popr?.xlsx_zapisano && popr.zapisano && popr.zapisano > popr.xlsx_zapisano && popr.akty
          && Object.values(popr.akty).some((a) => a.zmieniono && a.zmieniono > popr.xlsx_zapisano)) info.push('zmiany po zapisie xlsx – zapisz ponownie');
    }
    pozycje.push({ r, nazwa, jest, pelny: sprawdzone >= r.akty, info: info.join(' · ') });
  }
  const dostepne = pozycje.filter((p) => p.jest);
  if (dostepne.length < 2) {
    await okno('Scal lata', el('p', { text: `Do scalenia potrzebne są co najmniej dwa roczniki z zapisanym xlsx Geneteki (przycisk „Pobierz xlsx genetyka”). Teraz: ${dostepne.length}.` }),
      [['OK', null, true]]);
    return;
  }
  const lista = el('ul', { class: 'lata' }, pozycje.map((p) => el('li', {},
    el('input', { type: 'checkbox', id: `rok-${p.r.rok}`, checked: p.jest, disabled: !p.jest, 'data-rok': String(p.r.rok) }),
    el('label', { for: `rok-${p.r.rok}`, text: String(p.r.rok) }),
    el('span', { class: `info${!p.jest || !p.pelny || /zmiany/.test(p.info) ? ' uwaga' : ''}`, text: p.info }))));
  const wybor = await okno('Scal lata – jeden plik dla Geneteki', [
    el('p', { text: 'Zaznacz roczniki. W pliku wynikowym będą jeden pod drugim, bez powtarzania nagłówków i bez pustych wierszy, z ciągłą L.p.' }),
    lista], [['Anuluj', null], ['Dalej', 'dalej', true]]);
  if (!wybor) return;
  const lata = [...lista.querySelectorAll('input:checked')].map((c) => +c.dataset.rok).sort((a, b) => a - b);
  if (lata.length < 2) { toast('Zaznacz co najmniej dwa roczniki.'); return; }
  const wybrane = pozycje.filter((p) => lata.includes(p.r.rok));
  const ostrzezenia = [];
  const dziury = [];
  for (let y = lata[0]; y <= lata[lata.length - 1]; y++) if (!lata.includes(y)) dziury.push(y);
  if (dziury.length) ostrzezenia.push(`W zakresie ${lata[0]}–${lata[lata.length - 1]} brakuje lat: ${dziury.join(', ')} (Geneteka publikuje tylko pełne roczniki – nazwa pliku i tak poda cały zakres).`);
  const niepelne = wybrane.filter((p) => !p.pelny || /zmiany/.test(p.info)).map((p) => `${p.r.rok} (${p.info})`);
  if (niepelne.length) ostrzezenia.push(`Nie w pełni sprawdzone albo zmienione po zapisie xlsx: ${niepelne.join(', ')}.`);
  const nazwa = nazwaPliku(r0, `${lata[0]}-${lata[lata.length - 1]}`);
  const potw = await okno('Scal lata', [
    el('p', {}, 'Powstanie plik ', el('b', { text: nazwa }), ` w folderze ${r0.parafia}/${r0.ksiega}.`),
    ostrzezenia.length ? el('ul', {}, ostrzezenia.map((o) => el('li', { text: o }))) : null,
  ], [['Anuluj', false], [ostrzezenia.length ? 'Scal mimo to' : 'Scal', true, true]]);
  if (!potw) return;
  try {
    const wszystkie = [];
    for (const p of wybrane) {
      const bufor = await st.dysk.czytaj(`${p.r.sciezka}/${p.nazwa}`, 'binarny');
      if (!bufor) throw new Error(`nie ma pliku ${p.nazwa}`);
      const wiersze = await czytajZgony(JSZip, bufor);
      wszystkie.push(...wiersze);
    }
    const plik = await zbudujXlsx(JSZip, st.szablon, wszystkie);
    await zapiszPlikXlsx(`${r0.parafia}/${r0.ksiega}/${nazwa}`, nazwa, plik);
  } catch (e) {
    komunikat(`Nie udało się scalić lat: ${e.message}`, 'blad');
  }
}

// ------------------------------------------------------------------ link rocznika

function odswiezPtaszekLinku() {
  const p = $('#ptaszek-linku');
  const reczny = st.popr?.link !== undefined;
  p.classList.toggle('widoczny', reczny);
  p.title = reczny ? `Wpisany ręcznie. Z linków w Weryfikacji: „${K.linkRocznika(st.dane) || '(brak)'}”. Kliknij, żeby przywrócić.` : '';
  $('#link-rocznika').classList.toggle('do-sprawdzenia', !biezacyLink());
  $('#link-rocznika').title = biezacyLink() ? '' : 'Brak linku – wpisz link do jednostki w Skanotece (np. …/id5557-sy5500-kt8)';
}

function zmienLink(v) {
  const auto = K.linkRocznika(st.dane);
  const t = v.trim();
  if (t === auto) delete st.popr.link; else st.popr.link = t;
  przelicz();
  odswiezPtaszekLinku();
  zaplanujZapis();
  rysujGeneteke();
  rysujListe();
}

// ------------------------------------------------------------------ zdarzenia

function wPolu(e) {
  return e.target instanceof Element ? e.target.closest('input, textarea, select, [contenteditable]') : null;
}

function podlacz() {
  $('#rocznik').addEventListener('change', (e) => otworzRocznik(st.roczniki[+e.target.value]));
  $('#b-xlsx').addEventListener('click', pobierzXlsx);
  $('#b-scal').addEventListener('click', scalLata);
  $('#b-folder').addEventListener('click', wybierzFolder);
  $('#b-odswiez').addEventListener('click', async () => { await zapiszTeraz(); await wczytajRoczniki(); toast('Wczytano ponownie.'); });
  $('#b-poprzedni').addEventListener('click', () => pokazAkt(st.i - 1));
  $('#b-nastepny').addEventListener('click', () => pokazAkt(st.i + 1));
  $('#b-sprawdzone').addEventListener('click', () => {
    if (stanAktu(st.i).sprawdzony) { oznaczSprawdzony(false); pokazAkt(st.i, false); } else sprawdzoneIDalej();
  });
  $('#b-duplikuj').addEventListener('click', duplikujWiersz);
  $('#obraz').addEventListener('click', () => $('#obraz').classList.toggle('duzy'));
  const link = $('#link-rocznika');
  link.addEventListener('change', () => zmienLink(link.value));
  $('#ptaszek-linku').addEventListener('click', () => { link.value = K.linkRocznika(st.dane); zmienLink(link.value); });
  document.addEventListener('keydown', (e) => {
    if (!st.r || !$('#okno').classList.contains('ukryta')) return;
    const pole = wPolu(e);
    if (e.key === 'Enter' && !e.shiftKey && !e.altKey && (!pole || pole.tagName === 'INPUT' || e.ctrlKey)) {
      if (pole?.id === 'link-rocznika') return;
      if (pole?.tagName === 'TEXTAREA' && !e.ctrlKey) return;
      e.preventDefault();
      if (pole) pole.blur();
      sprawdzoneIDalej();
    } else if ((e.altKey && e.key === 'ArrowRight') || (!pole && (e.key === 'ArrowRight' || e.key === 'PageDown'))) {
      e.preventDefault(); pokazAkt(st.i + 1);
    } else if ((e.altKey && e.key === 'ArrowLeft') || (!pole && (e.key === 'ArrowLeft' || e.key === 'PageUp'))) {
      e.preventDefault(); pokazAkt(st.i - 1);
    }
  });
  window.addEventListener('beforeunload', (e) => { if (st.brudne) { zapiszTeraz(); e.preventDefault(); e.returnValue = ''; } });
  document.addEventListener('visibilitychange', () => { if (document.hidden) zapiszTeraz(); });
  const naglowek = document.querySelector('header');
  const ustawWys = () => document.documentElement.style.setProperty('--wys-naglowka', `${naglowek.offsetHeight}px`);
  new ResizeObserver(ustawWys).observe(naglowek);
}

podlacz();
start().catch((e) => komunikat(`Błąd: ${e.message}`, 'blad'));
