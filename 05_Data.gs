/**
 * СЛУЖЕБНЫЙ ЛИСТ ДАННЫХ + ЛОГ
 * ==========================
 * _Данные — лог метрик по дням. Источник истины для рендера и алертов.
 * Запись уникальна по ключу (accountId | sku | date).
 *
 * Записи ЧАСТИЧНЫЕ: один запуск пишет обложку в день «сегодня», а метрики —
 * в день «вчера». mergeUpsertRecords_ сливает их по ключу, обновляя только
 * переданные поля, поэтому обложка и метрики одного дня не затирают друг друга.
 */

// Колонки листа _Данные: [заголовок, ключ записи]
const DATA_COLUMNS = [
  ['Дата',               'date'],
  ['ID аккаунта',        'accountId'],
  ['Кабинет',            'accountName'],
  ['Артикул',            'sku'],
  ['Ozon SKU',           'numericSku'],
  ['Название',           'name'],
  ['URL обложки',        'coverUrl'],
  ['Хэш обложки',        'coverHash'],
  ['CTR',                'ctr'],
  ['Показы',             'views'],
  ['В корзину (шт)',     'toCart'],
  ['Конв. в корзину %',  'cartConvPct'],
  ['Заказы',             'orders'],
  ['CR корзина→заказ %', 'cartToOrderPct'],
  ['Выкупы',             'buyouts'],
  ['% выкупа',           'buyoutPct'],
  ['Выручка',            'revenue'],
  // диагностические сырые показатели аналитики
  ['Переходы pdp (hits)','pdpViews'],
  ['Показы поиск (сесс)','sessSearch'],
  ['Переходы pdp (сесс)','sessPdp'],
  // рекламные метрики Performance API (в конце — безопасная миграция)
  ['Рекл. CTR %',        'adCtr']
];
const DATA_KEYS = DATA_COLUMNS.map(c => c[1]);

function getDataSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET.DATA);
  if (!sh) {
    sh = ss.insertSheet(SHEET.DATA);
  }
  // Всегда приводим заголовок к актуальному набору колонок (на случай миграции)
  sh.getRange(1, 1, 1, DATA_COLUMNS.length).setValues([DATA_COLUMNS.map(c => c[0])]);
  sh.setFrozenRows(1);
  return sh;
}

function rowToRecord_(row) {
  const rec = {};
  DATA_KEYS.forEach((k, i) => { rec[k] = row[i]; });
  // Sheets превращает 'yyyy-MM-dd' в Date при записи — нормализуем обратно,
  // иначе ключ дня и разбор даты ломаются (и upsert плодит дубли).
  if (rec.date instanceof Date) {
    const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
    rec.date = Utilities.formatDate(rec.date, tz, 'yyyy-MM-dd');
  } else {
    rec.date = String(rec.date);
  }
  return rec;
}

function recordKey_(rec) {
  return `${rec.accountId}|${rec.sku}|${rec.date}`;
}

/**
 * Сливает набор ЧАСТИЧНЫХ записей в лист _Данные.
 * Для существующей строки (по accountId|sku|date) обновляются ТОЛЬКО переданные
 * в записи поля; остальные сохраняются. Новые ключи добавляются.
 */
function mergeUpsertRecords_(records) {
  const sh = getDataSheet_();
  const ncol = DATA_COLUMNS.length;
  const lastRow = sh.getLastRow();

  const existing = lastRow > 1 ? sh.getRange(2, 1, lastRow - 1, ncol).getValues() : [];

  // индекс по нормализованному ключу
  const keyToIndex = {};
  existing.forEach((row, i) => { keyToIndex[recordKey_(rowToRecord_(row))] = i; });

  for (const rec of records) {
    const key = recordKey_(rec);
    let row;
    if (key in keyToIndex) {
      row = existing[keyToIndex[key]];
    } else {
      row = new Array(ncol).fill('');
      existing.push(row);
      keyToIndex[key] = existing.length - 1;
    }
    // обновляем только те поля, что реально присутствуют в записи
    DATA_KEYS.forEach((k, i) => {
      if (Object.prototype.hasOwnProperty.call(rec, k)) row[i] = rec[k];
    });
  }

  if (existing.length) {
    sh.getRange(2, 1, existing.length, ncol).setValues(existing);
  }
}

/** Все записи листа _Данные как массив объектов (дата нормализована). */
function readAllData_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET.DATA);
  if (!sh || sh.getLastRow() < 2) return [];
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, DATA_COLUMNS.length).getValues();
  return rows.map(rowToRecord_);
}


// ====== ЛОГ (буферизованный) ======
const LOG_BUFFER_ = [];

function log_(level, func, message) {
  console.log(level, func, message);
  LOG_BUFFER_.push([new Date(), level, func, message]);
}

function flushLog_() {
  if (LOG_BUFFER_.length === 0) return;
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET.LOG);
    if (!sh) return;
    const startRow = sh.getLastRow() + 1;
    const n = LOG_BUFFER_.length;
    sh.getRange(startRow, 1, n, 4).setValues(LOG_BUFFER_);
    // показываем дату И время в колонке «Время»
    sh.getRange(startRow, 1, n, 1).setNumberFormat('yyyy-MM-dd HH:mm:ss');
    LOG_BUFFER_.length = 0;
    if (sh.getLastRow() > 5000) sh.deleteRows(2, 1000);
  } catch (e) {
    console.log('flushLog_ error:', e.message);
  }
}
