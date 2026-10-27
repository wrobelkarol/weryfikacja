'use strict';
/* Strona www dla pomocnika (GitHub Pages): logowanie, odszyfrowanie danych i zapis poprawek na GitHubie.
   Udaje serwer weryfikacja.pyw (te same adresy /api/...), więc app.js działa tak samo jak na komputerze.
   Dane w repo są zaszyfrowane (AES-GCM) kluczem z loginu i hasła - hasła nie ma w kodzie strony.
   Poprawki idą do gałęzi "poprawki" (d/w<id>, jeden plik na rocznik), skąd weryfikacja.pyw wczytuje je
   do odczyt_zweryfikowany.json. Do zapisu potrzebny jest token GitHuba (zaszyfrowany w d/1). */

window.WWW = (() => {
  const REPO = 'wrobelkarol/weryfikacja';
  const GALAZ_POPRAWEK = 'poprawki';
  const SOL = '4a867d1d82a22daebae50b4d6a48f7d6';       // ta sama w weryfikacja.pyw
  const ITERACJE = 250000;
  const CO_ILE_WYSYLAC = 20000;                        // ms między zapisami na GitHubie (każdy zapis to commit)
  const LS_KLUCZ = 'weryfikacja.klucz', LS_ZALEGLE = 'weryfikacja.zalegle';
  const GH = `https://api.github.com/repos/${REPO}/contents/`;

  let klucz = null, token = null, spis = null, spisPobrany = 0;
  const sha = new Map();          // plik poprawek -> sha na GitHubie (bez niego GitHub nie pozwoli nadpisać)
  const zalegle = new Map();      // rocznik -> { plik, dane }: zapisane w przeglądarce, jeszcze nie na GitHubie
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
    return r.ok ? r.arrayBuffer() : undefined;
  }

  async function zPages(sciezka, swiezy) {
    const r = await fetch(sciezka, { cache: swiezy ? 'no-cache' : 'default' });
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
    if (!token) {
      const r = await fetch(`https://raw.githubusercontent.com/${REPO}/${GALAZ_POPRAWEK}/d/${nazwa}`, { cache: 'no-cache' });
      return r.ok ? rozpakuj(await r.arrayBuffer()) : null;
    }
    const r = await fetch(`${GH}d/${nazwa}?ref=${GALAZ_POPRAWEK}`, { cache: 'no-store', headers: naglowki('application/vnd.github+json') });
    if (r.status === 404) {
      sha.set(nazwa, null);
      return null;
    }
    if (!r.ok) throw new Error(`GitHub: błąd ${r.status}`);
    const d = await r.json();
    sha.set(nazwa, d.sha);
    return rozpakuj(d.content ? zB64(d.content) : await zApi('d/' + nazwa, GALAZ_POPRAWEK));
  }

  /* odczyt_zweryfikowany z kilku miejsc: z każdego wycinka zostaje wersja poprawiona najpóźniej */
  function scal(...zapisy) {
    const wycinki = new Map();
    let wynik = null;
    for (const z of zapisy) {
      if (!z) continue;
      wynik = { ...z };
      for (const w of z.wycinki || []) {
        const k = String(w.wycinek), byl = wycinki.get(k);
        if (!byl || (+w.zmieniono || 0) >= (+byl.zmieniono || 0)) wycinki.set(k, w);
      }
    }
    if (wynik) wynik.wycinki = [...wycinki.values()];
    return wynik;
  }

  // ---------------------------------------------------------------- adresy /api/... jak w weryfikacja.pyw

  const odpowiedz = (d) => new Response(JSON.stringify(d), { headers: { 'Content-Type': 'application/json' } });

  function teraz() {
    const d = new Date(), p = (n) => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
  }

  async function get(url) {
    const [adres, zapytanie] = url.split('?');
    if (adres === '/api/start') return { korzen: '', pola: spis.pola, pola_urodzenia: spis.pola_urodzenia, typy: spis.typy, rodzaje: spis.rodzaje };
    if (adres === '/api/drzewo') {
      if (Date.now() - spisPobrany > 5000) {
        spis = (await json('0')) || spis;
        spisPobrany = Date.now();
      }
      return spis.drzewo;
    }
    if (adres === '/api/folder') return folder(new URLSearchParams(zapytanie).get('sciezka'));
    throw new Error('tego nie ma na stronie www');
  }

  async function post(url, cialo) {
    if (url === '/api/zapisz') return zapisz(cialo);
    throw new Error('tego nie ma na stronie www');
  }

  async function folder(rocznik) {
    const r = spis.roczniki[rocznik];
    if (!r) throw new Error('tego rocznika nie ma na stronie');
    const [ai, lokalny, www] = await Promise.all([json(r.ai), json(r.zw), poprawki(r.www)]);
    const czekajace = zalegle.get(rocznik);
    adresyObrazow.splice(0).forEach((u) => URL.revokeObjectURL(u));
    return {
      sciezka: rocznik, opis: r.opis,
      wycinki: r.wycinki.map(({ plik: p, wycinek, zmieniony }) => ({ plik: p, wycinek, zmieniony })),
      ai, ai_blad: null, ai_zmieniony: r.ai_zmieniony,
      zweryfikowany: scal(lokalny, www, czekajace && czekajace.dane),
    };
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
      for (const [k, v] of JSON.parse(localStorage.getItem(LS_ZALEGLE) || '[]')) zalegle.set(k, v);
    } catch (e) { /* bez znaczenia */ }
  }

  function zapisz(cialo) {
    const r = spis.roczniki[cialo.sciezka];
    if (!r) throw new Error('tego rocznika nie ma na stronie');
    zalegle.set(cialo.sciezka, { plik: r.www, dane: cialo.dane });
    zapamietajZalegle();
    const zapisano = teraz();
    if (!token) {
      return odpowiedz({ zapisano, pobrano: null, ostrzezenie: true,
        komunikat: 'Zapisano tylko w tej przeglądarce – strona nie ma tokenu, poprawki nie idą na GitHuba.' });
    }
    zaplanujWysylke(document.visibilityState === 'hidden');
    return odpowiedz({ zapisano, pobrano: null, komunikat: `Zapisano ${zapisano.slice(11)}, wysyłanie…` });
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
      for (const [rocznik, z] of [...zalegle]) {
        await wyslijPlik(z, keepalive);
        if (zalegle.get(rocznik) === z) zalegle.delete(rocznik);   // w trakcie wysyłania mogły dojść nowe poprawki
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
        // ktoś (np. druga karta) zapisał w międzyczasie - jego poprawki zostają, z każdego wycinka nowsza wersja
        dane = scal(await poprawki(z.plik), dane);
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
      if (r.status === 401 || r.status === 403) throw new Error('GitHub nie przyjął tokenu – mógł wygasnąć');
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

  async function zaladujObraz(img) {
    const r = spis.roczniki[img.dataset.rocznik];
    const w = r && r.wycinki.find((x) => x.plik === img.dataset.plik);
    if (!w) return;
    try {
      const dane = await plik(w.id, false);
      if (!dane) throw new Error('jeszcze się wysyła, odśwież stronę za minutę');
      const typ = { png: 'image/png', gif: 'image/gif', webp: 'image/webp' }[w.plik.split('.').pop().toLowerCase()] || 'image/jpeg';
      const url = URL.createObjectURL(new Blob([await odszyfruj(dane)], { type: typ }));
      adresyObrazow.push(url);
      img.src = url;
    } catch (e) {
      img.alt = 'Nie udało się wczytać wycinka: ' + e.message;
    }
  }

  // ---------------------------------------------------------------- logowanie

  async function wczytajStart() {
    const t = await zPages('d/1', true);
    token = t ? new TextDecoder().decode(await odszyfruj(t)) : null;     // zły login albo hasło -> OperationError
    spis = await json('0');
    if (!spis) throw new Error('Na stronie nie ma jeszcze danych.');
    spisPobrany = Date.now();
  }

  function formularz(blad) {
    return new Promise((gotowe) => {
      const okno = document.createElement('div');
      okno.className = 'logowanie';
      okno.innerHTML = `<form>
        <h1>Weryfikacja odczytów</h1>
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
    if (zalegle.size) zaplanujWysylke(true);
  }

  return { zaloguj, get, post, obraz };
})();
