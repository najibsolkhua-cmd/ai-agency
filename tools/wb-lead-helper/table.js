// Table page: preview collected data and export it to xlsx.

const COLS = [
  ['Статус', 'status'],
  ['Магазин на WB', 'sellerName'],
  ['Юрлицо / ИП', 'legalName'],
  ['Тип', 'type'],
  ['Обращение', 'greeting'],
  ['Email', 'email'],
  ['Тип почты', 'emailType'],
  ['ИНН', 'inn'],
  ['ОГРН / ОГРНИП', 'ogrn'],
  ['КПП', 'kpp'],
  ['Адрес', 'address'],
  ['Продано товаров', 'sold'],
  ['% выкупа', 'buyout'],
  ['На WB', 'onWb'],
  ['Рейтинг', 'rating'],
  ['Оценок', 'ratingsCount'],
  ['Страница продавца', 'sellerUrl'],
  ['Фото товара 1', 'img1'],
  ['Фото товара 2', 'img2'],
  ['Фото товара 3', 'img3'],
  ['Файл фото', 'photoFile'],
  ['Файл скрина', 'shotFile'],
  ['Сохранено', 'savedAt'],
  ['Дата письма', 'sentAt'],
  ['Комментарий', 'comment'],
];

const TARGET_COLS = [
  ['Бренд', 'brand'],
  ['Пример товара', 'title'],
  ['Ссылка на товар', 'url'],
  ['Фото', 'img'],
  ['Поисковый запрос', 'query'],
  ['Найден', 'foundAt'],
];

const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

async function load() {
  const d = await chrome.storage.local.get(['rows', 'targets']);
  return { rows: Object.values(d.rows || {}), targets: Object.values(d.targets || {}) };
}

function htmlTable(items, cols, imgKeys) {
  if (!items.length) return '<p class="muted">Пока пусто.</p>';
  const head = cols.map(([h]) => `<th>${esc(h)}</th>`).join('');
  const body = items.map((r) => '<tr>' + cols.map(([, k]) => {
    const v = r[k] || '';
    if (imgKeys.includes(k) && v) return `<td><img src="${esc(v)}" alt=""></td>`;
    if (/^https?:/.test(v)) return `<td><a href="${esc(v)}" target="_blank">открыть</a></td>`;
    return `<td>${esc(v)}</td>`;
  }).join('') + '</tr>').join('');
  return `<table><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table>`;
}

async function render() {
  const { rows, targets } = await load();
  const withMail = rows.filter((r) => r.email).length;
  document.getElementById('summary').textContent =
    `Продавцов: ${rows.length} (с почтой: ${withMail}) · Брендов к обходу: ${targets.length}`;
  document.getElementById('rows').innerHTML = htmlTable(rows, COLS, ['img1']);
  document.getElementById('targets').innerHTML = htmlTable(targets, TARGET_COLS, ['img']);
}

function sheet(items, cols) {
  const ws = XLSX.utils.aoa_to_sheet([cols.map(([h]) => h), ...items.map((r) => cols.map(([, k]) => r[k] || ''))]);
  ws['!cols'] = cols.map(([h]) => ({ wch: Math.max(12, Math.min(40, h.length + 6)) }));
  return ws;
}

const LETTER = [
  ['Тема', 'Фото на модели для {Магазин на WB}'],
  ['Текст',
    'Здравствуйте{, Обращение}!\n\n' +
    'Увидел ваш магазин {Магазин на WB} на Wildberries. Я делаю фото товаров на ИИ-моделях для маркетплейсов — пара примеров «до/после» по ссылке: [ссылка на работы].\n\n' +
    'Готов бесплатно сделать одно фото на модели по вашему товару, чтобы вы оценили качество. Полный сет из 10 фото на артикул — 3 990 ₽, срок 24 часа, главное фото на белом фоне в подарок. Под ваш бренд разрабатываю отдельную модель, она будет только у вас.\n\n' +
    'Если интересно — пришлите фото товара в ответ на это письмо или в Telegram: [ник].\n\n' +
    'С уважением,\n[Имя]\n[Telegram] · [сайт]\n\n' +
    'Если неинтересно — ответьте «нет», больше не напишу.'],
  ['Правила', 'Каждое письмо — лично: обращение, название магазина, их товар. 20–30 писем в день, одно напоминание через 3 дня, не писать тем, кто ответил «нет».'],
];

document.getElementById('export').addEventListener('click', async () => {
  const { rows, targets } = await load();
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, sheet(rows, COLS), 'Продавцы');
  XLSX.utils.book_append_sheet(wb, sheet(targets, TARGET_COLS), 'Бренды для обхода');
  const ls = XLSX.utils.aoa_to_sheet(LETTER);
  ls['!cols'] = [{ wch: 12 }, { wch: 110 }];
  XLSX.utils.book_append_sheet(wb, ls, 'Шаблон письма');
  XLSX.writeFile(wb, `wb-leads-${new Date().toISOString().slice(0, 10)}.xlsx`);
});

document.getElementById('refresh').addEventListener('click', render);
document.getElementById('clearTargets').addEventListener('click', async () => {
  if (confirm('Очистить список брендов для обхода?')) { await chrome.storage.local.set({ targets: {} }); render(); }
});
document.getElementById('clearAll').addEventListener('click', async () => {
  if (confirm('Удалить всех сохранённых продавцов и бренды? Сначала скачайте xlsx.')) { await chrome.storage.local.clear(); render(); }
});
chrome.storage.onChanged.addListener(render);
render();
