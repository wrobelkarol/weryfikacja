'use strict';
/* Strona www (GitHub Pages): logowanie, odszyfrowanie danych i zapis poprawek na GitHubie.
   Udaje serwer weryfikacja.pyw (te same adresy /api/...), więc app.js działa tak samo jak na komputerze.
   Z tego samego pliku korzysta Weryfikacja Geneteka na stronie (folder geneteka/, window.WWW.dane).
   Dane w repo są zaszyfrowane (AES-GCM) kluczem z loginu i hasła - hasła nie ma w kodzie strony.
   Poprawki idą do gałęzi "poprawki", jeden plik na rocznik i rodzaj: d/w<id> (odczyt_zweryfikowany - Weryfikacja),
   d/g<id> (genetyka_poprawki - Weryfikacja Geneteka), d/x<id> (ostatni xlsx Geneteki), skąd weryfikacja.pyw
   przenosi je do folderów roczników. Do zapisu potrzebny jest token GitHuba (zaszyfrowany w d/1). */

window.WWW = (() => {
  const REPO = 'wrobelkarol/weryfikacja';
  const GALAZ_POPRAWEK = 'poprawki';
  const SOL = '4a867d1d82a22daebae50b4d6a48f7d6';       // ta sama w weryfikacja.pyw
  const ITERACJE = 250000;
  const CO_ILE_WYSYLAC = 20000;                        // ms między zapisami na GitHubie (każdy zapis to commit)
  const LS_KLUCZ = 'weryfikacja.klucz', LS_ZALEGLE = 'weryfikacja.zalegle';
  // token wklejony na stronie (przycisk „Wklej token”) – zaszyfrowany kluczem z loginu, tylko w tej przeglądarce
  const LS_TOKEN = 'weryfikacja.token';
  const GH = `https://api.github.com/repos/${REPO}/contents/`;
  // strona Weryfikacji Geneteka leży w podfolderze – dane (d/...) są piętro wyżej
  const BAZA = window.WWW_BAZA || '';

  let klucz = null, token = null, spis = null, spisPobrany = 0;
  const sha = new Map();          // plik poprawek -> sha na GitHubie (bez niego GitHub nie pozwoli nadpisać)
  const zalegle = new Map();      // plik poprawek -> { plik, dane, rodzaj }: zapisane w przeglądarce, jeszcze nie na GitHubie
  const adresyObrazow = [];
  let wysylanie = false, ostatniaWysylka = 0, timerWysylki = null;

  // ---------------------------------------------------------------- szyfrowanie

  const hex = (s) => new Uint8Array(s.match(/../g).map((b) => parseInt(b, 16)));

  function doB64(bufor) {
    const b = new Uint8Array(bufor);
    let s = '';
    for (let i = 0; i < b.length; i += 0x8000) s += String.fromCharCode.apply(null, b.subarray(i, i + 0x8000));
    return btoa(s);
  }

  function zB64(s) {
    const t = atob(s.replace(/\s/g, ''));
    const b = new Uint8Array(t.length);
    for (let i = 0; i < t.length; i++) b[i] = t.charCodeAt(i);
    return b.buffer;
  }

  async function wyprowadzKlucz(login, haslo) {
    const baza = await crypto.subtle.importKey('raw', new TextEncoder().encode(`${login}\n${haslo}`), 'PBKDF2', false, ['deriveKey']);
    return crypto.subtle.deriveKey({ name: 'PBKDF2', salt: hex(SOL), iterations: ITERACJE, hash: 'SHA-256' },
      baza, { name: 'AES-GCM', length: 256 }, true, ['encrypt', 'decrypt']);
  }

  function odszyfruj(bufor) {
    return crypto.subtle.decrypt({ name: 'AES-GCM', iv: bufor.slice(0, 12) }, klucz, bufor.slice(12));
  }

  async function zaszyfruj(dane) {
    const iv = crypto.getRandomValues(new Uint8Array(12));
    const szyfr = new Uint8Array(await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, klucz, dane));
    const wynik = new Uint8Array(12 + szyfr.length);
    wynik.set(iv);
    wynik.set(szyfr, 12);
    return wynik.buffer;
  }

  const strumien = (dane, przez) => new Response(new Blob([dane]).stream().pipeThrough(przez));

  async function rozpakuj(bufor) {        // zaszyfrowany gzip z JSON -> obiekt
    return JSON.parse(await strumien(await odszyfruj(bufor), new DecompressionStream('gzip')).text());
  }

  async function spakuj(obiekt) {
    return zaszyfruj(await strumien(JSON.stringify(obiekt), new CompressionStream('gzip')).arrayBuffer());
  }

  // ---------------------------------------------------------------- pobieranie z GitHuba

  const naglowki = (accept = 'application/vnd.github.raw+json') => ({ Authorization: `Bearer ${token}`, Accept: accept });

  /* Spis, odczyty i poprawki idą przez API GitHuba (od razu aktualne), wycinki z Pages (szybciej;
     świeżo wysłanych może tam jeszcze nie być - wtedy też przez API). undefined = API niedostępne. */
  async function zApi(sciezka, ref = 'main') {
    if (!token) return undefined;
    const r = await fetch(`${GH}${sciezka}?ref=${ref}`, { cache: 'no-store', headers: naglowki() });
    if (r.status === 404) return null;
    if (r.status === 401) tokenOdrzucony();
    return r.ok ? r.arrayBuffer() : undefined;
  }

  /* Token wygasł albo jest zły: strona dalej czyta dane (bez tokenu, z Pages i raw.githubusercontent),
     poprawki czekają w przeglądarce, a przycisk prosi o nowy token. */
  let tokenZly = false;
  function tokenOdrzucony() {
    if (!token) return;
    token = null;
    tokenZly = true;
    sha.clear();
    pokazPrzyciskTokenu();
    stan('Token GitHuba wygasł – kliknij „Wklej token” i wklej nowy. Do tego czasu poprawki zostają w tej przeglądarce.', true);
  }

  async function zPages(sciezka, swiezy) {
    const r = await fetch(BAZA + sciezka, { cache: swiezy ? 'no-cache' : 'default' });
    if (r.status === 404) return null;
    if (!r.ok) throw new Error(`błąd ${r.status} przy pobieraniu danych`);
    return r.arrayBuffer();
  }

  async function plik(nazwa, swiezy) {
    if (swiezy) {
      const d = await zApi('d/' + nazwa);
      if (d !== undefined) return d;
    }
    const d = await zPages('d/' + nazwa, swiezy);
    return d === null && !swiezy ? ((await zApi('d/' + nazwa)) || null) : d;
  }

  async function json(nazwa) {
    const d = nazwa && await plik(nazwa, true);
    return d ? rozpakuj(d) : null;
  }

  async function poprawki(nazwa) {       // plik poprawek z gałęzi "poprawki"; zapamiętuje jego sha
    if (!nazwa) return null;
    if (!token) {
      const r = await fetch(`https://raw.githubusercontent.com/${REPO}/${GALAZ_POPRAWEK}/d/${nazwa}`, { cache: 'no-cache' });
      return r.ok ? rozpakuj(await r.arrayBuffer()) : null;
    }
    const r = await fetch(`${GH}d/${nazwa}?ref=${GALAZ_POPRAWEK}`, { cache: 'no-store', headers: naglowki('application/vnd.github+json') });
    if (r.status === 404) {
      sha.set(nazwa, null);
      return null;
    }
    if (r.status === 401) {
      tokenOdrzucony();
      return poprawki(nazwa);           // jeszcze raz, bez tokenu
    }
    if (!r.ok) throw new Error(`GitHub: błąd ${r.status}`);
    const d = await r.json();
    sha.set(nazwa, d.sha);
    return rozpakuj(d.content ? zB64(d.content) : await zApi('d/' + nazwa, GALAZ_POPRAWEK));
  }

  /* Plik poprawek: wersja z GitHuba połączona z tym, co czeka w tej przeglądarce na wysłanie. */
  async function aktualne(nazwa, rodzaj) {
    const czekajace = zalegle.get(nazwa);
    return SCAL[rodzaj](await poprawki(nazwa), czekajace && czekajace.dane);
  }

  /* odczyt_zweryfikowany z kilku miejsc: z każdego wycinka zostaje wersja poprawiona najpóźniej;
     linki do skanów - z zapisu, w którym zmieniono je najpóźniej */
  function scal(...zapisy) {
    const wycinki = new Map();
    let wynik = null, linki = null;
    for (const z of zapisy) {
      if (!z) continue;
      // dane rocznika z pierwszego zapisu (z komputera); z kolejnych tylko nowsze „pobrano” (xlsx zapisany na stronie)
      if (!wynik) wynik = { ...z };
      else {
        if (String(z.pobrano || '') > String(wynik.pobrano || '')) {
          wynik.pobrano = z.pobrano;
          wynik.zmiany_po_pobraniu = z.zmiany_po_pobraniu;
        }
        if (String(z.xlsx_geneteki || '') > String(wynik.xlsx_geneteki || '')) wynik.xlsx_geneteki = z.xlsx_geneteki;
      }
      if (!linki || (+z.linki_zmieniono || 0) >= (+linki.linki_zmieniono || 0)) linki = z;
      for (const w of z.wycinki || []) {
        const k = String(w.wycinek), byl = wycinki.get(k);
        if (!byl || (+w.zmieniono || 0) >= (+byl.zmieniono || 0)) wycinki.set(k, w);
      }
    }
    if (wynik) {
      wynik.wycinki = [...wycinki.values()];
      for (const k of ['linki_skanow', 'link_jednostki', 'linki_zmieniono']) {
        if (linki[k] === undefined) delete wynik[k]; else wynik[k] = linki[k];
      }
    }
    return wynik;
  }

  /* genetyka_poprawki z kilku miejsc: z każdego aktu wersja zmieniona najpóźniej, reszta z najnowszego zapisu */
  function scalGeneteki(...zapisy) {
    const czas = (a) => [a && a.zmieniono, a && a.sprawdzony].filter((t) => typeof t === 'string').sort().pop() || '';
    let wynik = null;
    const akty = {};
    for (const z of zapisy) {
      if (!z) continue;
      if (!wynik || String(z.zapisano || '') >= String(wynik.zapisano || '')) wynik = { ...z };
      for (const [k, a] of Object.entries(z.akty || {})) if (!akty[k] || czas(a) >= czas(akty[k])) akty[k] = a;
    }
    if (wynik) wynik.akty = akty;
    return wynik;
  }

  /* xlsx Geneteki: zostaje ostatnio zapisany */
  function nowszy(...zapisy) {
    return zapisy.filter(Boolean).sort((a, b) => (+a.zapisano || 0) - (+b.zapisano || 0)).pop() || null;
  }

  const SCAL = { w: scal, g: scalGeneteki, x: nowszy };

  // ---------------------------------------------------------------- adresy /api/... jak w weryfikacja.pyw

  const odpowiedz = (d) => new Response(JSON.stringify(d), { headers: { 'Content-Type': 'application/json' } });

  function teraz() {
    const d = new Date(), p = (n) => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
  }

  async function odswiezSpis() {
    if (Date.now() - spisPobrany > 5000) {
      spis = (await json('0')) || spis;
      spisPobrany = Date.now();
    }
    return spis;
  }

  async function get(url) {
    const [adres, zapytanie] = url.split('?');
    const q = new URLSearchParams(zapytanie);
    if (adres === '/api/start') return { korzen: '', pola: spis.pola, pola_urodzenia: spis.pola_urodzenia, typy: spis.typy, rodzaje: spis.rodzaje };
    if (adres === '/api/drzewo') return (await odswiezSpis()).drzewo;
    if (adres === '/api/folder') return folder(q.get('sciezka'));
    if (adres === '/api/geneteka_lata') return lataGeneteki(q.get('sciezka'));
    throw new Error('tego nie ma na stronie www');
  }

  async function post(url, cialo) {
    if (url === '/api/zapisz') return zapisz(cialo);
    if (url === '/api/xlsx_geneteki') return xlsxGeneteki(cialo);
    if (url === '/api/xlsx_scalone') {
      pobierzPlik(cialo.nazwa, zB64(cialo.xlsx));
      return odpowiedz({ folder: 'pobrane pliki', komunikat: `Pobrano ${cialo.nazwa} (folder „Pobrane” w przeglądarce).` });
    }
    throw new Error('tego nie ma na stronie www');
  }

  /** Plik binarny (np. /api/plik_xlsx?sciezka=…) – odpowiednik fetch na komputerze. */
  async function binarny(url) {
    const [adres, zapytanie] = url.split('?');
    if (adres !== '/api/plik_xlsx') throw new Error('tego nie ma na stronie www');
    const sciezka = new URLSearchParams(zapytanie).get('sciezka');
    const i = sciezka.lastIndexOf('/');
    const dane = await xlsxRocznika(sciezka.slice(0, i), sciezka.slice(i + 1));
    if (!dane) throw new Error('nie ma pliku ' + sciezka.slice(i + 1));
    return dane;
  }

  async function folder(rocznik) {
    const r = spis.roczniki[rocznik];
    if (!r) throw new Error('tego rocznika nie ma na stronie');
    const [ai, lokalny, www, skany] = await Promise.all([json(r.ai), json(r.zw), aktualne(r.www, 'w'), json(r.skany)]);
    adresyObrazow.splice(0).forEach((u) => URL.revokeObjectURL(u));
    return {
      sciezka: rocznik, opis: r.opis,
      wycinki: r.wycinki.map(({ plik: p, wycinek, zmieniony }) => ({ plik: p, wycinek, zmieniony })),
      ai, ai_blad: null, ai_zmieniony: r.ai_zmieniony,
      zweryfikowany: scal(lokalny, www),
      skany: skany && skany.wycinki,
      geneteka_poprawki: !!r.gp,
    };
  }

  /** odczyt_zweryfikowany rocznika: z komputera + poprawki ze strony + czekające w przeglądarce. */
  async function zweryfikowany(rel) {
    const r = spis.roczniki[rel];
    const [lokalny, www] = await Promise.all([json(r.zw), aktualne(r.www, 'w')]);
    return scal(lokalny, www);
  }

  /** Ostatni xlsx Geneteki rocznika (zapisany na stronie albo wysłany z komputera) o tej nazwie. */
  async function xlsxRocznika(rel, nazwa) {
    const r = spis.roczniki[rel];
    if (!r) return null;
    const x = await aktualne(r.xl, 'x');
    const zKomputera = r.xlsx && (!nazwa || r.xlsx.nazwa === nazwa) ? r.xlsx : null;
    if (x && (!nazwa || x.nazwa === nazwa) && (!zKomputera || +x.zapisano >= +zKomputera.zapisano)) return zB64(x.xlsx);
    return zKomputera ? odszyfruj(await plik(zKomputera.id, true)) : null;
  }

  /** genetyka_poprawki rocznika: z komputera + ze strony + czekające w przeglądarce. */
  async function genetykaRocznika(rel) {
    const r = spis.roczniki[rel];
    return scalGeneteki(await json(r.gp), await aktualne(r.gen, 'g'));
  }

  async function nazwaXlsxRocznika(rel) {
    const r = spis.roczniki[rel];
    const x = await aktualne(r.xl, 'x');
    return (x && x.nazwa) || (r.xlsx && r.xlsx.nazwa) || null;
  }

  async function lataGeneteki(rel) {
    const r0 = spis.roczniki[rel];
    const { parafia, ksiega } = r0.opis;
    const lata = [];
    for (const [k, r] of Object.entries(spis.roczniki)) {
      if (r.opis.parafia !== parafia || r.opis.ksiega !== ksiega || !/^\d{4}$/.test(String(r.opis.rok))) continue;
      const d = (await zweryfikowany(k)) || {};
      const akty = (d.wycinki || []).filter((w) => (w.rodzaj || 'akt') === 'akt');
      const nazwa = await nazwaXlsxRocznika(k);
      lata.push({ rok: +r.opis.rok, sciezka: k, akty: akty.length, sprawdzone: akty.filter((w) => w.sprawdzony).length,
        pliki: nazwa ? [nazwa] : [], pobrano: d.pobrano || null, zmiany_po_pobraniu: !!d.zmiany_po_pobraniu });
    }
    lata.sort((a, b) => a.rok - b.rok);
    return { folder: `${parafia}/${ksiega}`, parafia, ksiega, lata };
  }

  // ---------------------------------------------------------------- zapis poprawek

  function stan(t, blad = false) {
    const e = document.getElementById('stan-zapisu');
    if (!e) return;
    e.textContent = t;
    e.classList.toggle('blad', blad);
  }

  function zapamietajZalegle() {
    try {
      if (zalegle.size) localStorage.setItem(LS_ZALEGLE, JSON.stringify([...zalegle]));
      else localStorage.removeItem(LS_ZALEGLE);
    } catch (e) { /* brak miejsca - poprawki zostają w tej karcie do wysłania */ }
  }

  function wczytajZalegle() {
    try {
      // starsze zapisy: klucz = rocznik, bez rodzaju (zawsze odczyt_zweryfikowany)
      for (const [, v] of JSON.parse(localStorage.getItem(LS_ZALEGLE) || '[]')) zalegle.set(v.plik, { rodzaj: 'w', ...v });
    } catch (e) { /* bez znaczenia */ }
  }

  /** Zapis pliku poprawek: od razu w przeglądarce, na GitHuba w kolejce (co CO_ILE_WYSYLAC). */
  function kolejkuj(nazwa, dane, rodzaj) {
    const byl = zalegle.get(nazwa);
    zalegle.set(nazwa, { plik: nazwa, rodzaj, dane: byl && rodzaj !== 'x' ? SCAL[rodzaj](byl.dane, dane) : dane });
    zapamietajZalegle();
    const zapisano = teraz();
    if (!token) {
      return { zapisano, ostrzezenie: true,
        komunikat: 'Zapisano tylko w tej przeglądarce – strona nie ma tokenu (kliknij „Wklej token”), poprawki czekają na wysłanie.' };
    }
    zaplanujWysylke(document.visibilityState === 'hidden');
    return { zapisano, komunikat: `Zapisano ${zapisano.slice(11)}, wysyłanie…` };
  }

  function zapisz(cialo) {
    const r = spis.roczniki[cialo.sciezka];
    if (!r) throw new Error('tego rocznika nie ma na stronie');
    const d = cialo.dane || {};
    return odpowiedz({ pobrano: d.pobrano || null, zmiany_po_pobraniu: d.zmiany_po_pobraniu, ...kolejkuj(r.www, d, 'w') });
  }

  /* „Pobierz xlsx genetyka” na stronie: plik pobiera przeglądarka, a na GitHubie zostaje jego kopia
     (weryfikacja.pyw wstawi ją do folderu rocznika). */
  function xlsxGeneteki(cialo) {
    const r = spis.roczniki[cialo.sciezka];
    if (!r) throw new Error('tego rocznika nie ma na stronie');
    const t = teraz();
    Object.assign(cialo.dane, { pobrano: t, xlsx_geneteki: t, zmiany_po_pobraniu: false });
    kolejkuj(r.www, cialo.dane, 'w');
    const wynik = kolejkuj(r.xl, { nazwa: cialo.nazwa, zapisano: Date.now(), zrodlo: 'weryfikacja', xlsx: cialo.xlsx }, 'x');
    pobierzPlik(cialo.nazwa, zB64(cialo.xlsx));
    return odpowiedz({ folder: 'pobrane pliki', zapisano: t,
      komunikat: `Pobrano ${cialo.nazwa}.` + (token ? ' Kopia trafi do folderu rocznika, gdy uruchomisz Weryfikację na komputerze.' : ` ${wynik.komunikat}`) });
  }

  function zaplanujWysylke(odRazu = false) {
    clearTimeout(timerWysylki);
    const za = odRazu ? 0 : Math.max(0, ostatniaWysylka + CO_ILE_WYSYLAC - Date.now());
    timerWysylki = setTimeout(() => wyslij(document.visibilityState === 'hidden'), za);
  }

  async function wyslij(keepalive = false) {
    if (wysylanie || !zalegle.size || !token) return;
    wysylanie = true;
    try {
      for (const [nazwa, z] of [...zalegle]) {
        await wyslijPlik(z, keepalive);
        if (zalegle.get(nazwa) === z) zalegle.delete(nazwa);   // w trakcie wysyłania mogły dojść nowe poprawki
      }
      zapamietajZalegle();
      if (!zalegle.size) stan('Zapisano i wysłano ' + teraz().slice(11));
    } catch (e) {
      stan(`Nie udało się wysłać poprawek (${e.message}) – są zapisane w tej przeglądarce, spróbuję ponownie.`, true);
    } finally {
      ostatniaWysylka = Date.now();
      wysylanie = false;
      if (zalegle.size) zaplanujWysylke();
    }
  }

  async function wyslijPlik(z, keepalive) {
    let dane = z.dane;
    for (let proba = 0; proba < 3; proba++) {
      if (!sha.has(z.plik) || proba) {
        // ktoś (np. druga karta albo drugi komputer) zapisał w międzyczasie - jego poprawki zostają
        dane = SCAL[z.rodzaj || 'w'](await poprawki(z.plik), dane);
      }
      const tresc = doB64(await spakuj(dane));
      const s = sha.get(z.plik);
      const cialo = JSON.stringify({ message: 'Poprawki', content: tresc, branch: GALAZ_POPRAWEK, ...(s ? { sha: s } : {}) });
      const r = await fetch(GH + 'd/' + z.plik, {
        method: 'PUT', body: cialo, keepalive: keepalive && cialo.length < 60000,
        headers: { ...naglowki('application/vnd.github+json'), 'Content-Type': 'application/json' },
      });
      if (r.ok) {
        sha.set(z.plik, (await r.json()).content.sha);
        return;
      }
      if (r.status === 401) {
        tokenOdrzucony();
        throw new Error('token GitHuba wygasł – kliknij „Wklej token”');
      }
      if (r.status === 403) throw new Error('GitHub nie pozwolił zapisać – token nie ma uprawnienia „Contents: Read and write”');
      if (r.status !== 409 && r.status !== 422) throw new Error(`GitHub: błąd ${r.status}`);
    }
    throw new Error('konflikt zapisu na GitHubie');
  }

  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'hidden' && zalegle.size) setTimeout(() => wyslij(true), 0);
  });
  window.addEventListener('beforeunload', (e) => {
    if (zalegle.size && token) {
      wyslij(true);
      e.preventDefault();
      e.returnValue = '';
    }
  });

  /** Plik do pobrania przez przeglądarkę (xlsx). */
  function pobierzPlik(nazwa, dane) {
    const url = URL.createObjectURL(new Blob([dane], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    const a = Object.assign(document.createElement('a'), { href: url, download: nazwa });
    document.body.append(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 20000);
  }

  // ---------------------------------------------------------------- wycinki (obrazy)

  const obserwator = new IntersectionObserver((wpisy) => {
    for (const w of wpisy) {
      if (!w.isIntersecting) continue;
      obserwator.unobserve(w.target);
      zaladujObraz(w.target);
    }
  }, { rootMargin: '1500px 0px' });

  function obraz(img, rocznik, nazwaPliku) {
    img.dataset.rocznik = rocznik;
    img.dataset.plik = nazwaPliku;
    obserwator.observe(img);
  }

  /** Adres (blob:) odszyfrowanego wycinka. */
  async function adresWycinka(rocznik, nazwaPliku) {
    const r = spis.roczniki[rocznik];
    const w = r && r.wycinki.find((x) => x.plik === nazwaPliku);
    if (!w) return '';
    const dane = await plik(w.id, false);
    if (!dane) throw new Error('jeszcze się wysyła, odśwież stronę za minutę');
    const typ = { png: 'image/png', gif: 'image/gif', webp: 'image/webp' }[w.plik.split('.').pop().toLowerCase()] || 'image/jpeg';
    const url = URL.createObjectURL(new Blob([await odszyfruj(dane)], { type: typ }));
    adresyObrazow.push(url);
    return url;
  }

  async function zaladujObraz(img) {
    try {
      const url = await adresWycinka(img.dataset.rocznik, img.dataset.plik);
      if (url) img.src = url;
    } catch (e) {
      img.alt = 'Nie udało się wczytać wycinka: ' + e.message;
    }
  }

  // ---------------------------------------------------------------- logowanie

  async function wczytajStart() {
    const t = await zPages('d/1', true);
    const zeStrony = t ? new TextDecoder().decode(await odszyfruj(t)) : null;     // zły login albo hasło -> OperationError
    token = (await tokenZPrzegladarki()) || zeStrony;           // wklejony tutaj ma pierwszeństwo
    spis = await json('0');
    if (!spis) throw new Error('Na stronie nie ma jeszcze danych.');
    spisPobrany = Date.now();
  }

  function formularz(blad) {
    return new Promise((gotowe) => {
      const okno = document.createElement('div');
      okno.className = 'logowanie';
      okno.innerHTML = `<form>
        <h1>${document.title || 'Weryfikacja odczytów'}</h1>
        <label>Login<input name="login" autocomplete="username" required></label>
        <label>Hasło<input name="haslo" type="password" autocomplete="current-password" required></label>
        <p class="blad"></p>
        <button class="glowny" type="submit">Zaloguj</button>
      </form>`;
      document.body.append(okno);
      const f = okno.querySelector('form'), przycisk = f.querySelector('button'), komunikat = f.querySelector('.blad');
      komunikat.textContent = blad || '';
      f.elements.login.focus();
      f.addEventListener('submit', async (e) => {
        e.preventDefault();
        przycisk.disabled = true;
        komunikat.textContent = 'Sprawdzam…';
        try {
          klucz = await wyprowadzKlucz(f.elements.login.value.trim(), f.elements.haslo.value);
          await wczytajStart();
          try { localStorage.setItem(LS_KLUCZ, doB64(await crypto.subtle.exportKey('raw', klucz))); } catch (e2) { /* zaloguje się następnym razem */ }
          okno.remove();
          gotowe();
        } catch (err) {
          klucz = null;
          komunikat.textContent = err.name === 'OperationError' ? 'Zły login albo hasło.' : err.message;
          przycisk.disabled = false;
        }
      });
    });
  }

  // ---------------------------------------------------------------- token wklejony na stronie

  async function tokenZPrzegladarki() {
    try {
      const z = localStorage.getItem(LS_TOKEN);
      return z ? new TextDecoder().decode(await odszyfruj(zB64(z))) : null;
    } catch (e) {
      return null;           // inny login albo uszkodzony wpis
    }
  }

  /** Czy GitHub przyjmuje token do zapisu poprawek (gałąź „poprawki”). */
  async function sprawdzToken(t) {
    const r = await fetch(`${GH}.gitkeep?ref=${GALAZ_POPRAWEK}`, { cache: 'no-store',
      headers: { Authorization: `Bearer ${t}`, Accept: 'application/vnd.github+json' } });
    if (r.status === 401 || r.status === 403) throw new Error('GitHub nie przyjął tego tokenu – sprawdź, czy skopiowałeś go w całości i czy ma dostęp do repozytorium „weryfikacja”.');
    if (!r.ok && r.status !== 404) throw new Error(`GitHub: błąd ${r.status}`);
  }

  function pokazPrzyciskTokenu() {
    const b = document.getElementById('b-token');
    if (!b) return;
    b.classList.remove('ukryta');
    b.textContent = token ? 'Token ✓' : tokenZly ? 'Wklej token (stary wygasł)' : 'Wklej token';
    b.title = token ? 'Strona zapisuje poprawki na GitHubie. Kliknij, żeby wkleić inny token.'
      : 'Strona nie ma tokenu – poprawki zostają w tej przeglądarce. Wklej token GitHuba, żeby szły na GitHuba.';
    b.classList.toggle('glowny', !token);
  }

  function oknoTokenu() {
    const okno = document.createElement('div');
    okno.className = 'logowanie';
    okno.innerHTML = `<form>
      <h1>Token GitHuba</h1>
      <p style="margin:0;font-size:13px">Token (zaczyna się od github_pat_) pozwala stronie zapisywać poprawki na GitHubie.
        Zostanie zapisany tylko w tej przeglądarce, zaszyfrowany.</p>
      <label>Token<input name="token" type="password" autocomplete="off" required></label>
      <p class="blad"></p>
      <div style="display:flex;gap:8px;justify-content:flex-end">
        <button type="button" data-anuluj>Anuluj</button>
        <button class="glowny" type="submit">Zapisz</button>
      </div>
    </form>`;
    document.body.append(okno);
    const f = okno.querySelector('form'), komunikat = f.querySelector('.blad');
    f.elements.token.focus();
    okno.querySelector('[data-anuluj]').addEventListener('click', () => okno.remove());
    f.addEventListener('submit', async (e) => {
      e.preventDefault();
      const t = f.elements.token.value.trim();
      komunikat.textContent = 'Sprawdzam…';
      try {
        await sprawdzToken(t);
        localStorage.setItem(LS_TOKEN, doB64(await zaszyfruj(new TextEncoder().encode(t))));
        token = t;
        sha.clear();
        okno.remove();
        pokazPrzyciskTokenu();
        stan(zalegle.size ? 'Token zapisany – wysyłam poprawki czekające w przeglądarce…' : 'Token zapisany – poprawki będą szły na GitHuba.');
        if (zalegle.size) zaplanujWysylke(true);
      } catch (err) {
        komunikat.textContent = err.message;
      }
    });
  }

  function wyloguj() {
    if (zalegle.size && !confirm('Część poprawek nie jest jeszcze wysłana. Wylogować mimo to?')) return;
    try { localStorage.removeItem(LS_KLUCZ); } catch (e) { /* bez znaczenia */ }
    location.reload();
  }

  async function zaloguj() {
    const b = document.getElementById('b-wyloguj');
    if (b) b.addEventListener('click', wyloguj);
    wczytajZalegle();
    let zapamietany = null, blad = '';
    try { zapamietany = localStorage.getItem(LS_KLUCZ); } catch (e) { /* bez znaczenia */ }
    if (zapamietany) {
      try {
        klucz = await crypto.subtle.importKey('raw', zB64(zapamietany), 'AES-GCM', true, ['encrypt', 'decrypt']);
        await wczytajStart();
      } catch (e) {
        klucz = null;
        if (e.name !== 'OperationError') blad = e.message;
      }
    }
    if (!klucz) await formularz(blad);
    const bt = document.getElementById('b-token');
    if (bt) bt.addEventListener('click', oknoTokenu);
    pokazPrzyciskTokenu();
    if (zalegle.size) zaplanujWysylke(true);
  }

  // dla strony Weryfikacji Geneteka (geneteka/dysk.js: DyskGitHub)
  const dane = {
    spis: () => spis, odswiezSpis, json, aktualne, zweryfikowany, genetykaRocznika, xlsxRocznika, nazwaXlsxRocznika, kolejkuj,
    adresWycinka, pobierzPlik, zB64, doB64, maToken: () => !!token, teraz,
  };

  return { zaloguj, get, post, binarny, obraz, wyloguj, dane };
})();
