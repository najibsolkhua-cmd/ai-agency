// Toolbar popup. Opening it grants activeTab, so screenshots work from here.

const $ = (id) => document.getElementById(id);
const WB = /^https:\/\/([a-z0-9-]+\.)?wildberries\.ru\//i;

function status(text, isError = false) {
  $('status').textContent = text;
  $('status').className = isError ? 'err' : '';
}

(async () => {
  const d = await chrome.storage.local.get(['rows', 'targets']);
  $('count').textContent = `Продавцов: ${Object.keys(d.rows || {}).length} · Брендов к обходу: ${Object.keys(d.targets || {}).length}`;

  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab || !WB.test(tab.url || '')) {
    status('Откройте сайт wildberries.ru в этой вкладке.', true);
    return;
  }
  const r = await chrome.runtime.sendMessage({ type: 'ensure', tabId: tab.id });
  if (!r || !r.ok) {
    status('Не удалось подключиться к странице. Обновите её (F5) и откройте это окно снова. ' + ((r && r.error) || ''), true);
    return;
  }
  let info = null;
  try { info = await chrome.tabs.sendMessage(tab.id, { type: 'ping' }); } catch (_) { /* ignore */ }
  if (!info || !info.ok) {
    status('Страница не отвечает. Обновите её (F5) и откройте это окно снова.', true);
    return;
  }
  if (info.page === 'seller') {
    status('Страница продавца — можно сохранять.');
    $('save').disabled = false;
  } else if (info.page === 'search') {
    status('Страница поиска — можно собрать бренды.');
    $('collect').disabled = false;
  } else {
    status('Откройте поиск или страницу продавца (в карточке товара: название продавца → «Все товары продавца»).');
  }

  $('save').addEventListener('click', async () => {
    await chrome.tabs.sendMessage(tab.id, { type: 'save', withShot: true });
    window.close();
  });
  $('collect').addEventListener('click', async () => {
    await chrome.tabs.sendMessage(tab.id, { type: 'collect' });
    window.close();
  });
})();

$('table').addEventListener('click', () => {
  chrome.tabs.create({ url: chrome.runtime.getURL('table.html') });
  window.close();
});
