// Dostęp do folderu „SKANY MAM DOŚĆ” na dysku użytkownika. Dane nigdy nie wychodzą z komputera.
//  - serwer: strona otwarta skrótem z pulpitu (127.0.0.1) – pliki czyta i zapisuje lokalny program genetyka.pyw;
//  - folder: File System Access API (np. strona z GitHub Pages) – folder wybierany raz, uchwyt pamiętany w IndexedDB.
// Ścieżki są względne wobec „SKANY MAM DOŚĆ”, segmenty oddzielone „/”.

const PODPIS = 'genetyka-weryfikacja';

function zakoduj(sciezka) {
  return encodeURIComponent(sciezka);
}

export class DyskSerwer {
  constructor() { this.rodzaj = 'serwer'; this.nazwa = 'SKANY MAM DOŚĆ'; }

  static async wykryj() {
    if (!/^(127\.0\.0\.1|localhost)$/.test(location.hostname)) return null;
    try {
      const r = await fetch('/api/ping', { cache: 'no-store' });
      const d = await r.json();
      if (d.podpis !== PODPIS) return null;
      const dysk = new DyskSerwer();
      dysk.korzen = d.korzen;
      dysk.nazwa = (d.korzen || '').split(/[\\/]/).filter(Boolean).pop() || dysk.nazwa;
      return dysk;
    } catch {
      return null;
    }
  }

  async lista(sciezka) {
    const r = await fetch(`/api/lista?sciezka=${zakoduj(sciezka)}`, { cache: 'no-store' });
    if (r.status === 404) return null;
    if (!r.ok) throw new Error((await r.json()).blad || `błąd ${r.status}`);
    return r.json();
  }

  async czytaj(sciezka, typ = 'tekst') {
    // brak=204: brakujący plik to nie błąd (np. genetyka_poprawki.json przy pierwszym otwarciu rocznika)
    const r = await fetch(`/api/plik?brak=204&sciezka=${zakoduj(sciezka)}`, { cache: 'no-store' });
    if (r.status === 404 || r.status === 204) return null;
    if (!r.ok) throw new Error((await r.json()).blad || `błąd ${r.status}`);
    return typ === 'binarny' ? r.arrayBuffer() : r.text();
  }

  async zapisz(sciezka, dane) {
    const r = await fetch(`/api/plik?sciezka=${zakoduj(sciezka)}`, {
      method: 'POST', headers: { 'X-Genetyka': '1', 'Content-Type': 'application/octet-stream' }, body: dane,
      keepalive: dane.byteLength < 60000,            // małe zapisy (poprawki) dochodzą też przy zamykaniu karty
    });
    if (!r.ok) throw new Error((await r.json().catch(() => ({}))).blad || `błąd zapisu ${r.status}`);
    return r.json();
  }

  async adresObrazu(sciezka) {
    return `/api/plik?sciezka=${zakoduj(sciezka)}`;
  }

  async ping() {
    await fetch('/api/ping', { cache: 'no-store' });
  }
}

// ------------------------------------------------------------------ File System Access API

const BAZA = 'genetyka-weryfikacja';

function bazaIdb() {
  return new Promise((ok, zle) => {
    const z = indexedDB.open(BAZA, 1);
    z.onupgradeneeded = () => z.result.createObjectStore('uchwyty');
    z.onsuccess = () => ok(z.result);
    z.onerror = () => zle(z.error);
  });
}

async function idb(tryb, fn) {
  const db = await bazaIdb();
  return new Promise((ok, zle) => {
    const tx = db.transaction('uchwyty', tryb);
    const wynik = fn(tx.objectStore('uchwyty'));
    tx.oncomplete = () => ok(wynik.result);
    tx.onerror = () => zle(tx.error);
  });
}

export function folderDostepny() {
  return typeof window.showDirectoryPicker === 'function';
}

export class DyskFolder {
  constructor(uchwyt) { this.rodzaj = 'folder'; this.uchwyt = uchwyt; this.nazwa = uchwyt.name; this.obrazy = new Map(); }

  /** Zapamiętany folder (null, gdy brak). uprawnienie: 'granted' albo 'prompt' (wtedy trzeba kliknąć „Połącz”). */
  static async zapamietany() {
    if (!folderDostepny()) return null;
    try {
      const uchwyt = await idb('readonly', (s) => s.get('korzen'));
      if (!uchwyt) return null;
      const dysk = new DyskFolder(uchwyt);
      dysk.uprawnienie = await uchwyt.queryPermission({ mode: 'readwrite' });
      return dysk;
    } catch {
      return null;
    }
  }

  static async wybierz() {
    const uchwyt = await window.showDirectoryPicker({ id: 'skany', mode: 'readwrite' });
    await idb('readwrite', (s) => s.put(uchwyt, 'korzen'));
    const dysk = new DyskFolder(uchwyt);
    dysk.uprawnienie = 'granted';
    return dysk;
  }

  async polacz() {
    this.uprawnienie = await this.uchwyt.requestPermission({ mode: 'readwrite' });
    return this.uprawnienie === 'granted';
  }

  async folder(segmenty, utworz = false) {
    let h = this.uchwyt;
    for (const seg of segmenty) {
      try {
        h = await h.getDirectoryHandle(seg, { create: utworz });
      } catch (e) {
        // nazwa z innym zapisem Unicode (NFC/NFD) – szukamy po normalizacji
        let znaleziony = null;
        for await (const [n, uch] of h.entries()) {
          if (uch.kind === 'directory' && n.normalize('NFC') === seg.normalize('NFC')) { znaleziony = uch; break; }
        }
        if (!znaleziony) {
          if (e.name === 'NotFoundError' || e.name === 'TypeError') return null;
          throw e;
        }
        h = znaleziony;
      }
    }
    return h;
  }

  async plik(sciezka) {
    const seg = sciezka.split('/').filter(Boolean);
    const folder = await this.folder(seg.slice(0, -1));
    if (!folder) return null;
    const nazwa = seg[seg.length - 1];
    try {
      return await (await folder.getFileHandle(nazwa)).getFile();
    } catch {
      for await (const [n, uch] of folder.entries()) {
        if (uch.kind === 'file' && n.normalize('NFC') === nazwa.normalize('NFC')) return uch.getFile();
      }
      return null;
    }
  }

  async lista(sciezka) {
    const folder = await this.folder(sciezka.split('/').filter(Boolean));
    if (!folder) return null;
    const foldery = [], pliki = [];
    for await (const [n, uch] of folder.entries()) {
      if (n.startsWith('.')) continue;
      (uch.kind === 'directory' ? foldery : pliki).push(n);
    }
    return { foldery: foldery.sort(), pliki: pliki.sort() };
  }

  async czytaj(sciezka, typ = 'tekst') {
    const f = await this.plik(sciezka);
    if (!f) return null;
    return typ === 'binarny' ? f.arrayBuffer() : f.text();
  }

  async zapisz(sciezka, dane) {
    const seg = sciezka.split('/').filter(Boolean);
    const folder = await this.folder(seg.slice(0, -1));
    if (!folder) throw new Error('Nie ma folderu: ' + seg.slice(0, -1).join('/'));
    const uch = await folder.getFileHandle(seg[seg.length - 1], { create: true });
    const zapis = await uch.createWritable();
    await zapis.write(dane);
    await zapis.close();
    return { ok: true };
  }

  async adresObrazu(sciezka) {
    if (this.obrazy.has(sciezka)) return this.obrazy.get(sciezka);
    const f = await this.plik(sciezka);
    if (!f) return '';
    const url = URL.createObjectURL(f);
    if (this.obrazy.size > 40) {               // nie trzymamy w pamięci wszystkich obrazów
      const [k, v] = this.obrazy.entries().next().value;
      URL.revokeObjectURL(v);
      this.obrazy.delete(k);
    }
    this.obrazy.set(sciezka, url);
    return url;
  }

  async ping() {}
}

// ------------------------------------------------------------------ strona www (GitHub)

/* Strona w internecie (folder geneteka/ na stronie Weryfikacji): dane z repozytorium GitHub – zaszyfrowane, te same
   co na stronie Weryfikacji (../www.js, window.WWW). Rocznik = „Parafia/Księga/Rok”, jak na dysku.
   Poprawki (genetyka_poprawki.json) i xlsx idą do gałęzi „poprawki”; weryfikacja.pyw przenosi je do folderu rocznika. */
export class DyskGitHub {
  constructor(www) { this.rodzaj = 'github'; this.nazwa = 'strona'; this.d = www.dane; }

  static async wykryj() {
    if (!window.WWW) return null;
    await window.WWW.zaloguj();
    return new DyskGitHub(window.WWW);
  }

  /** „Przytkowice/Zgony/1890/pocięte/plik” -> [rocznik, reszta po roczniku] */
  podziel(sciezka) {
    const seg = sciezka.split('/').filter(Boolean);
    return [seg.slice(0, 3).join('/'), seg.slice(3)];
  }

  async lista(sciezka) {
    const spis = await this.d.odswiezSpis();
    const seg = sciezka.split('/').filter(Boolean);
    const roczniki = Object.keys(spis.roczniki).map((r) => r.split('/'));
    if (seg.length < 3) {
      const foldery = [...new Set(roczniki.filter((r) => r.length >= 3 && seg.every((s, i) => r[i] === s)).map((r) => r[seg.length]))];
      return foldery.length || !seg.length ? { foldery: foldery.sort(), pliki: [] } : null;
    }
    const rel = seg.slice(0, 3).join('/');
    if (!spis.roczniki[rel]) return null;
    if (seg.length > 3) return { foldery: [], pliki: [] };
    const xlsx = await this.d.nazwaXlsxRocznika(rel);
    return { foldery: ['pocięte'], pliki: xlsx ? [xlsx] : [] };
  }

  async czytaj(sciezka, typ = 'tekst') {
    const [rel, reszta] = this.podziel(sciezka);
    if (!this.d.spis().roczniki[rel]) return null;
    const nazwa = reszta[reszta.length - 1] || '';
    let obiekt;
    if (nazwa === 'odczyt_zweryfikowany.json') obiekt = await this.d.zweryfikowany(rel);
    else if (nazwa === 'genetyka_poprawki.json') obiekt = await this.d.genetykaRocznika(rel);
    else if (nazwa === 'skany.json') obiekt = await this.d.json(this.d.spis().roczniki[rel].skany);
    else if (/\.xlsx$/i.test(nazwa)) {
      const bufor = await this.d.xlsxRocznika(rel, nazwa);
      return bufor && (typ === 'binarny' ? bufor : new TextDecoder().decode(bufor));
    } else return null;
    return obiekt ? JSON.stringify(obiekt) : null;
  }

  async zapisz(sciezka, dane) {
    const seg = sciezka.split('/').filter(Boolean);
    const nazwa = seg[seg.length - 1];
    const r = this.d.spis().roczniki[seg.slice(0, 3).join('/')];
    if (nazwa === 'genetyka_poprawki.json' && r) {
      const wynik = this.d.kolejkuj(r.gen, JSON.parse(new TextDecoder().decode(dane)), 'g');
      if (wynik.ostrzezenie) throw new Error(wynik.komunikat);
      return { ok: true };
    }
    if (/\.xlsx$/i.test(nazwa)) {
      // plik pobiera przeglądarka; xlsx rocznika zostaje też na GitHubie (trafi do folderu rocznika na komputerze)
      this.d.pobierzPlik(nazwa, dane);
      if (r && seg.length === 4) {
        this.d.kolejkuj(r.xl, { nazwa, zapisano: Date.now(), zrodlo: 'geneteka', xlsx: this.d.doB64(dane) }, 'x');
        return { ok: true, komunikat: `Pobrano ${nazwa}.` + (this.d.maToken()
          ? ' Kopia trafi do folderu rocznika, gdy uruchomisz Weryfikację na komputerze.'
          : ' Strona nie ma tokenu – plik jest tylko w pobranych.') };
      }
      return { ok: true, komunikat: `Pobrano ${nazwa} (folder „Pobrane” w przeglądarce).` };
    }
    throw new Error('Na stronie nie da się zapisać pliku ' + nazwa);
  }

  async adresObrazu(sciezka) {
    const [rel, reszta] = this.podziel(sciezka);
    return this.d.adresWycinka(rel, reszta[reszta.length - 1]);
  }

  async ping() {}
}

/** Zapis przez pobranie pliku (gdy nie ma ani serwera, ani dostępu do folderu). */
export function pobierzPlik(nazwa, dane, typ = 'application/octet-stream') {
  const url = URL.createObjectURL(new Blob([dane], { type: typ }));
  const a = Object.assign(document.createElement('a'), { href: url, download: nazwa });
  document.body.append(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}
