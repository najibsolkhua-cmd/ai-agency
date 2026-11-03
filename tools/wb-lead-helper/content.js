// WB Lead Helper — content script for www.wildberries.ru.
// Works only on the page the user has open; it never navigates by itself.
// WB renders its pages with React and class names change often, so everything
// here is found by visible text ("ИНН:", "товаров продано", ...), not by CSS classes.

(() => {
  'use strict';

  // ---------- parsing (pure functions, also used by tests) ----------

  const PERSONAL_MAIL = /@(yandex|ya|mail|bk|list|inbox|internet|gmail|rambler|icloud|me|outlook|hotmail)\./i;

  function parseLegal(text) {
    const t = text.replace(/ /g, ' ');
    const lines = t.split('\n').map((s) => s.trim()).filter(Boolean);
    const get = (re) => (t.match(re) || [])[1] || '';
    const email = get(/([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})/);
    const ogrnip = get(/ОГРНИП:\s*(\d{15})/);
    const ogrn = ogrnip || get(/ОГРН:\s*(\d{13})/);
    const inn = get(/ИНН:\s*(\d{10,12})/);
    const kpp = get(/КПП:\s*(\d{9})/);
    const address = (lines.find((l) => /^\d{6},/.test(l)) || '').trim();
    const nameLines = [];
    for (const l of lines) {
      if (/^(ОГРН|ИНН|КПП|Номер регистрации|Электронная почта|Адрес)/i.test(l) || /^\d{6},/.test(l) || l.includes('@')) break;
      nameLines.push(l);
    }
    const legalName = nameLines.join(' ').replace(/\s+/g, ' ').trim();
    const isIP = /^(ИП\b|Индивидуальный предприниматель)/i.test(legalName) || Boolean(ogrnip) || inn.length === 12;
    let greeting = '';
    if (isIP) {
      const fio = legalName.replace(/^(ИП|Индивидуальный предприниматель)\s*/i, '').split(/\s+/);
      if (fio.length >= 3) greeting = `${fio[1]} ${fio[2]}`;
      else if (fio.length === 2) greeting = fio[1];
    }
    return {
      legalName,
      type: isIP ? 'ИП' : 'Юрлицо',
      greeting,
      inn,
      ogrn,
      kpp,
      address,
      email,
      emailType: email ? (PERSONAL_MAIL.test(email) ? 'личная' : 'корпоративная') : '',
    };
  }

  function parseStats(text) {
    const t = text.replace(/ /g, ' ');
    const num = (s) => (s || '').replace(/\s+/g, '');
    return {
      sold: num((t.match(/([\d ]{1,15})\s*товаров продано/i) || [])[1]),
      buyout: ((t.match(/(\d{1,3})\s*%\s*заказов выкуплено/i) || [])[1] || ''),
      onWb: ((t.match(/(\d+\s*(?:год|года|лет)(?:\s*\d+\s*мес)?|\d+\s*мес)\s*на Wildberries/i) || [])[1] || '').replace(/\s+/g, ' '),
      rating: ((t.match(/(\d[.,]\d)\s*·\s*[\d ]+\s*оцен/i) || [])[1] || ''),
      ratingsCount: num((t.match(/\d[.,]\d\s*·\s*([\d ]+)\s*оцен/i) || [])[1]),
    };
  }

  function brandFromCardText(text) {
    const line = text.split('\n').map((s) => s.trim()).find((l) => l.includes(' / '));
    if (!line) return { brand: '', title: '' };
    const i = line.indexOf(' / ');
    return { brand: line.slice(0, i).trim(), title: line.slice(i + 3).trim() };
  }

  if (typeof window === 'undefined' || typeof document === 'undefined') {
    module.exports = { parseLegal, parseStats, brandFromCardText };
    return;
  }

  // Do not run twice when the popup injects the script into an open tab.
  if (window.__wblhLoaded) return;
  window.__wblhLoaded = true;

  // ---------- helpers ----------

  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const safe = (s) => (s || 'seller').replace(/[\\/:*?"<>|\s]+/g, '_').slice(0, 60);
  const send = (msg) => chrome.runtime.sendMessage(msg);

  function toast(text, ms = 3500) {
    let el = document.getElementById('wblh-toast');
    if (!el) {
      el = document.createElement('div');
      el.id = 'wblh-toast';
      document.documentElement.appendChild(el);
    }
    el.textContent = text;
    el.style.display = 'block';
    clearTimeout(el._t);
    el._t = setTimeout(() => (el.style.display = 'none'), ms);
  }

  function isSellerPage() {
    return /\/seller\//.test(location.pathname) || /каталог продавца/i.test(document.title);
  }
  function isSearchPage() {
    return /search\.aspx/.test(location.pathname) || new URLSearchParams(location.search).has('search');
  }

  // Smallest visible element whose text contains the legal info block.
  function findLegalBlock() {
    let best = null;
    for (const el of document.querySelectorAll('div, section, ul, dl, aside')) {
      if (el.closest('#wblh-panel')) continue;
      const t = el.innerText || '';
      if (t.length > 1500 || !/ИНН:/.test(t) || !/ОГРН/.test(t)) continue;
      if (!el.offsetParent && getComputedStyle(el).position !== 'fixed') continue;
      if (!best || t.length < best.innerText.length) best = el;
    }
    return best;
  }

  async function waitFor(fn, timeout = 2500, step = 150) {
    const end = Date.now() + timeout;
    while (Date.now() < end) {
      const v = fn();
      if (v) return v;
      await sleep(step);
    }
    return null;
  }

  // Try to open the "i" tooltip next to the seller name: hover, then click
  // small icon-like elements right after the <h1>.
  async function tryOpenInfo() {
    const h1 = document.querySelector('h1');
    if (!h1) return null;
    const scope = h1.parentElement || document.body;
    const candidates = [...scope.querySelectorAll('button, [role="button"], span, i, svg')].filter((el) => {
      if (el.closest('#wblh-panel') || h1.contains(el)) return false;
      const r = el.getBoundingClientRect();
      return r.width > 0 && r.width <= 40 && r.height <= 40;
    });
    for (const el of candidates.slice(0, 6)) {
      for (const type of ['pointerover', 'mouseover', 'mouseenter', 'pointerenter']) {
        el.dispatchEvent(new MouseEvent(type, { bubbles: true }));
      }
      let found = await waitFor(findLegalBlock, 600);
      if (found) return found;
      el.dispatchEvent(new MouseEvent('click', { bubbles: true }));
      found = await waitFor(findLegalBlock, 800);
      if (found) return found;
    }
    return null;
  }

  function productImages(limit = 3) {
    const urls = [];
    for (const img of document.querySelectorAll('img')) {
      const src = img.currentSrc || img.src || '';
      if (/wbbasket\.ru\/.*\/images\//.test(src) && !urls.includes(src)) urls.push(src);
      if (urls.length >= limit) break;
    }
    return urls;
  }

  async function loadData() {
    const d = await chrome.storage.local.get(['rows', 'targets']);
    return { rows: d.rows || {}, targets: d.targets || {} };
  }

  // ---------- actions ----------

  let pendingSave = null;

  async function saveSeller(withShot) {
    if (!isSellerPage()) {
      toast('Откройте страницу продавца (название продавца в карточке товара → «Все товары продавца»).');
      return;
    }
    let block = findLegalBlock() || (await tryOpenInfo());
    if (!block) {
      pendingSave = { withShot };
      toast('Наведите или нажмите на значок «i» рядом с названием продавца — данные сохранятся автоматически.', 6000);
      return;
    }
    pendingSave = null;
    await finishSave(block, withShot);
  }

  async function finishSave(block, withShot) {
    const legal = parseLegal(block.innerText);
    const stats = parseStats(document.body.innerText);
    const h1 = document.querySelector('h1');
    const sellerName = (h1 && h1.innerText.trim()) || document.title.split(' - ')[0].trim();
    const imgs = productImages();
    const key = legal.inn || location.pathname;
    const base = `wb-leads/${safe(sellerName)}_${legal.inn || 'noinn'}`;

    let photoFile = '';
    if (imgs[0]) {
      const ext = (imgs[0].match(/\.(webp|jpe?g|png)(\?|$)/i) || [, 'webp'])[1];
      const r = await send({ type: 'download', url: imgs[0], filename: `${base}_товар.${ext}` });
      if (r && r.ok) photoFile = `${base}_товар.${ext}`;
    }

    let shotFile = '';
    if (withShot) {
      const panel = document.getElementById('wblh-panel');
      if (panel) panel.style.visibility = 'hidden';
      await sleep(120);
      const cap = await send({ type: 'capture' });
      if (panel) panel.style.visibility = '';
      if (cap && cap.ok) {
        const r = await send({ type: 'download', url: cap.dataUrl, filename: `${base}_скрин.png` });
        if (r && r.ok) shotFile = `${base}_скрин.png`;
      }
    }

    const { rows } = await loadData();
    const prev = rows[key] || {};
    rows[key] = {
      ...prev,
      savedAt: new Date().toISOString().slice(0, 16).replace('T', ' '),
      sellerName,
      ...legal,
      ...stats,
      sellerUrl: location.href.split('?')[0],
      img1: imgs[0] || '',
      img2: imgs[1] || '',
      img3: imgs[2] || '',
      photoFile: photoFile || prev.photoFile || '',
      shotFile: shotFile || prev.shotFile || '',
      status: prev.status || 'Новый',
      comment: prev.comment || '',
    };
    await chrome.storage.local.set({ rows });
    const n = Object.keys(rows).length;
    const what = [legal.email ? `почта ${legal.email}` : 'почты нет', withShot ? (shotFile ? 'скрин ✓' : 'скрин не удался') : ''].filter(Boolean).join(', ');
    toast(`${prev.savedAt ? 'Обновлено' : 'Сохранено'}: ${sellerName} — ${what}. Всего: ${n}`);
    renderPanel();
  }

  async function collectFromSearch() {
    const links = [...document.querySelectorAll('a[href*="/catalog/"][href*="detail.aspx"]')];
    const { targets } = await loadData();
    let added = 0;
    for (const a of links) {
      const card = a.closest('article') || a.closest('[class*="card"]') || a.parentElement;
      if (!card) continue;
      const { brand, title } = brandFromCardText(card.innerText || '');
      if (!brand || targets[brand.toLowerCase()]) continue;
      const img = card.querySelector('img');
      targets[brand.toLowerCase()] = {
        brand,
        title,
        url: a.href.split('?')[0],
        img: img ? img.currentSrc || img.src : '',
        query: new URLSearchParams(location.search).get('search') || '',
        foundAt: new Date().toISOString().slice(0, 10),
      };
      added++;
    }
    await chrome.storage.local.set({ targets });
    toast(`Новых брендов: ${added}. В списке для обхода: ${Object.keys(targets).length}. Прокрутите ниже и нажмите ещё раз, чтобы добавить следующие.`, 5000);
    renderPanel();
  }

  // Finish a pending save when the user opens the "i" tooltip by hand.
  new MutationObserver(() => {
    if (!pendingSave) return;
    const block = findLegalBlock();
    if (block) {
      const { withShot } = pendingSave;
      pendingSave = null;
      setTimeout(() => finishSave(block, withShot), 250);
    }
  }).observe(document.documentElement, { childList: true, subtree: true });

  // ---------- panel ----------

  async function renderPanel() {
    let panel = document.getElementById('wblh-panel');
    if (!panel) {
      panel = document.createElement('div');
      panel.id = 'wblh-panel';
      (document.body || document.documentElement).appendChild(panel);
    }
    const { rows, targets } = await loadData();
    const btns = [];
    if (isSellerPage()) btns.push('<button data-a="save">Сохранить продавца</button>');
    if (isSearchPage()) btns.push('<button data-a="collect">Собрать бренды с поиска</button>');
    btns.push('<button data-a="table" class="wblh-secondary">Таблица и выгрузка xlsx</button>');
    panel.innerHTML = `
      <span class="wblh-min" data-a="min">–</span>
      <b>WB Lead Helper</b>
      <div class="wblh-count">Продавцов: ${Object.keys(rows).length} · Брендов к обходу: ${Object.keys(targets).length}</div>
      ${btns.join('')}
      <div class="wblh-hint">Со скриншотом: значок расширения → «Сохранить продавца + скриншот» или Alt+Shift+S</div>`;
  }

  document.addEventListener('click', (e) => {
    const a = e.target.closest && e.target.closest('#wblh-panel [data-a]');
    if (!a) return;
    e.preventDefault();
    const act = a.dataset.a;
    if (act === 'save') saveSeller(false);
    if (act === 'collect') collectFromSearch();
    if (act === 'table') send({ type: 'openTable' });
    if (act === 'min') document.getElementById('wblh-panel').classList.toggle('wblh-collapsed');
  }, true);

  chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
    if (msg.type === 'ping') {
      sendResponse({ ok: true, page: isSellerPage() ? 'seller' : isSearchPage() ? 'search' : 'other' });
      return false;
    }
    if (msg.type === 'save') saveSeller(Boolean(msg.withShot));
    if (msg.type === 'collect') collectFromSearch();
    sendResponse({ ok: true });
    return false;
  });

  // WB is a single-page app: re-render the panel when the URL changes,
  // and put it back if the site's own scripts remove it.
  let lastUrl = '';
  setInterval(() => {
    if (location.href !== lastUrl) {
      lastUrl = location.href;
      pendingSave = null;
      renderPanel();
    } else if (!document.getElementById('wblh-panel')) {
      renderPanel();
    }
  }, 1000);
})();
