// ══════════════════════════════════════════════════════════════════
//  САЙДБАР С КАЛЕНДАРЁМ
// ══════════════════════════════════════════════════════════════════

/** Открывает боковую панель с выбором периода */
function showCalendarSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('CalendarSidebar')
    .setTitle('📅 Период калькулятора')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/** Вызывается из сайдбара при загрузке: возвращает список блоков и текущую дату */
function getCalendarData() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();

  if (!isCabinetSheet_(sh.getName())) {
    return { error: 'Перейдите на лист кабинета (ИП Иванова, ООО...)' };
  }

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const colsToRead = Math.max(CALC_COL, 3);
  const allData = lastRow > 0
    ? sh.getRange(1, 1, lastRow, colsToRead).getValues()
    : [];

  const skus = [];
  for (let b = 0; ; b++) {
    const topIdx = HEADER_ROWS + b * BLOCK_H;
    if (topIdx >= allData.length) break;
    const sku      = String(allData[topIdx][0] || '').trim();
    const name     = String(allData[topIdx][2] || '').trim();
    const existing = String(allData[topIdx][CALC_COL - 1] || '').trim();
    if (!sku) break;
    skus.push({ sku, name: name.slice(0, 45), existing });
  }

  const nDays    = lastCol >= FIRST_DATA_COL ? lastCol - FIRST_DATA_COL + 1 : 0;
  const dateKeys = nDays > 0 ? loadDateKeys_(sh, nDays).filter(Boolean) : [];

  const activeRow      = sh.getActiveCell().getRow();
  const activeBlockIdx = Math.max(0,
    Math.min(
      Math.floor((activeRow - HEADER_ROWS - 1) / BLOCK_H),
      skus.length - 1
    )
  );

  return {
    sheetName:      sh.getName(),
    skus,
    dateKeys,
    activeBlockIdx,
    today: Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd')
  };
}

/**
 * Вызывается из сайдбара: пишет диапазон в ячейки D и запускает калькулятор.
 * opts: { sheetName, selectedSkus: ['sku1',...] | 'all', from: 'yyyy-MM-dd', to: 'yyyy-MM-dd' }
 */
function applyCalcPeriod(opts) {
  const ss      = SpreadsheetApp.getActive();
  const sh      = ss.getSheetByName(opts.sheetName);
  if (!sh) return { error: 'Лист не найден' };

  const accounts = getAccounts_();
  const account  = accounts.find(a => a.name === opts.sheetName);
  if (!account) return { error: 'Кабинет не найден' };

  const products = getActiveProducts_().filter(p => p.accountId === account.id);
  const lastRow  = sh.getLastRow();
  const rangeStr = formatCalcRangeStr_(opts.from, opts.to);
  const skuSet   = opts.selectedSkus === 'all' ? null : new Set(opts.selectedSkus);

  for (let b = 0; b < products.length; b++) {
    const sku    = products[b].sku;
    if (skuSet && !skuSet.has(sku)) continue;
    const topRow = HEADER_ROWS + b * BLOCK_H + 1;
    if (topRow > lastRow) break;
    sh.getRange(topRow, CALC_COL).setValue(rangeStr);
  }

  runCalculatorForSheet_(sh, products);
  runPrevPeriodForSheet_(sh, products);
  return { ok: true, range: rangeStr };
}

/** Форматирует диапазон из 'yyyy-MM-dd' в '01.06–15.06' */
function formatCalcRangeStr_(from, to) {
  function fmt(dk) {
    const [, m, d] = dk.split('-');
    return `${d}.${m}`;
  }
  return from === to ? fmt(from) : `${fmt(from)}–${fmt(to)}`;
}


// ══════════════════════════════════════════════════════════════════
//  КАЛЬКУЛЯТОР ПЕРИОДА
//  Среднее для % метрик, сумма для абсолютных. Дни (F+) не трогает.
// ══════════════════════════════════════════════════════════════════

// Метрики, для которых считается сумма (остальные — среднее)
const SUM_LABELS = ['Показы', 'Заказы', 'Выкупы'];

/** Пункт меню: пересчитать калькулятор + прошлый период для всех кабинетных листов */
function runCalculatorMenu() {
  const ss       = SpreadsheetApp.getActive();
  const accounts = getAccounts_();
  const products = getActiveProducts_();
  let total = 0;

  for (const account of accounts) {
    const sh = ss.getSheetByName(account.name);
    if (!sh) continue;
    const prods = products.filter(p => p.accountId === account.id);
    total += runCalculatorForSheet_(sh, prods);
    runPrevPeriodForSheet_(sh, prods);
  }

  ss.toast(`Обновлено блоков: ${total}`, '🧮 Калькулятор', 4);
}

/**
 * Пересчитывает колонку D (CALC_COL) для одного листа кабинета.
 * Возвращает число блоков, для которых был задан и применён диапазон.
 */
function runCalculatorForSheet_(sh, prods) {
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < HEADER_ROWS + BLOCK_H || lastCol < FIRST_DATA_COL) return 0;

  const nDays = lastCol - FIRST_DATA_COL + 1;
  if (nDays <= 0) return 0;

  const dateKeys = loadDateKeys_(sh, nDays);   // ['yyyy-MM-dd', ...] из метаданных
  const allVals  = sh.getRange(1, 1, lastRow, lastCol).getValues();

  const writes = [];   // [{row: 1-based, value}]
  let updated  = 0;

  for (let b = 0; b < prods.length; b++) {
    const topIdx = HEADER_ROWS + b * BLOCK_H;  // 0-based индекс строки «Обложка»
    if (topIdx >= allVals.length) break;

    // Диапазон дат из ячейки D (CALC_COL) строки «Обложка»
    const rangeRaw = String(allVals[topIdx][CALC_COL - 1] || '').trim();
    if (!rangeRaw) continue;

    const range = parseDateRange_(rangeRaw);
    if (!range) continue;

    // Индексы дней, попадающих в диапазон
    const dayIdxs = [];
    for (let i = 0; i < nDays; i++) {
      const dk = dateKeys[i];
      if (dk && dk >= range.from && dk <= range.to) dayIdxs.push(i);
    }
    if (!dayIdxs.length) continue;

    // Считаем значение для каждой строки метрики
    for (let li = 0; li < BLOCK_H; li++) {
      if (li === IMAGE_ROW_OFFSET) continue;         // строка «Обложка» — там диапазон

      const rowIdx = topIdx + li;
      const label  = String(allVals[rowIdx][FIRST_DATA_COL - 2] || ''); // col F = index 5

      if (MANUAL_LABELS.indexOf(label) >= 0) continue;  // ручные строки — не трогаем
      if (!(label in METRIC_ROW_FIELD)) continue;

      const vals = dayIdxs
        .map(i => allVals[rowIdx][(FIRST_DATA_COL - 1) + i])
        .filter(v => v !== '' && v !== null && v !== undefined)
        .map(Number)
        .filter(n => !isNaN(n));

      if (!vals.length) {
        writes.push({ row: rowIdx + 1, value: '' });
        continue;
      }

      const sum    = vals.reduce((a, x) => a + x, 0);
      const result = SUM_LABELS.indexOf(label) >= 0 ? sum : sum / vals.length;
      writes.push({ row: rowIdx + 1, value: Math.round(result * 100) / 100 });
    }

    updated++;
  }

  // Пишем результаты в колонку D (CALC_COL)
  for (const w of writes) {
    sh.getRange(w.row, CALC_COL).setValue(w.value);
  }

  return updated;
}

/**
 * Читает сохранённые диапазоны из строк «Обложка» колонки CALC_COL.
 * Вызывается ДО sh.clear() в renderCabinet_, чтобы не потерять введённые даты.
 * Поддерживает миграцию: если CALC_COL пуст, пробует CALC_COL-1 (старая позиция).
 * Возвращает {sku → строка диапазона}.
 */
function readCalcRanges_(sh) {
  const map     = {};
  const lastRow = sh.getLastRow();
  if (lastRow < HEADER_ROWS + 1) return map;

  const colsToRead = CALC_COL;  // читаем до CALC_COL включительно
  const data = sh.getRange(1, 1, lastRow, colsToRead).getValues();

  for (let b = 0; ; b++) {
    const topIdx = HEADER_ROWS + b * BLOCK_H;
    if (topIdx >= data.length) break;
    const sku = String(data[topIdx][0] || '').trim();
    if (!sku) break;

    // Пробуем CALC_COL, при пустом — CALC_COL-1 (миграция со старой позиции)
    const newVal = String(data[topIdx][CALC_COL - 1] || '').trim();
    const oldVal = CALC_COL >= 2 ? String(data[topIdx][CALC_COL - 2] || '').trim() : '';
    const val    = parseDateRange_(newVal) ? newVal
                 : parseDateRange_(oldVal) ? oldVal
                 : '';
    if (val) map[sku] = val;
  }
  return map;
}


// ── Прошлый период ────────────────────────────────────────────────

/**
 * Вычисляет «прошлый период» той же длины, что заканчивается за день до from.
 * Возвращает {from, to} в формате 'yyyy-MM-dd'.
 */
function computePrevPeriod_(from, to) {
  const MS = 24 * 60 * 60 * 1000;
  const a  = new Date(from + 'T00:00:00');
  const b  = new Date(to   + 'T00:00:00');
  const dur = Math.round((b - a) / MS);  // длина периода - 1 (inclusive)

  const prevTo   = new Date(a.getTime() - MS);
  const prevFrom = new Date(prevTo.getTime() - dur * MS);

  function toKey(d) {
    return d.getFullYear() + '-'
      + String(d.getMonth() + 1).padStart(2, '0') + '-'
      + String(d.getDate()).padStart(2, '0');
  }
  return { from: toKey(prevFrom), to: toKey(prevTo) };
}

/**
 * Заполняет колонку PREV_COL (D) для всех блоков с заданным диапазоном в CALC_COL (E).
 * Строка «Обложка» PREV_COL получает строку вида "07.06–10.06", метрики — значения.
 */
function runPrevPeriodForSheet_(sh, prods) {
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < HEADER_ROWS + BLOCK_H || lastCol < FIRST_DATA_COL) return 0;

  const nDays   = lastCol - FIRST_DATA_COL + 1;
  if (nDays <= 0) return 0;

  const dateKeys = loadDateKeys_(sh, nDays);
  const allVals  = sh.getRange(1, 1, lastRow, lastCol).getValues();

  const writes = [];
  let updated  = 0;

  for (let b = 0; b < prods.length; b++) {
    const topIdx = HEADER_ROWS + b * BLOCK_H;
    if (topIdx >= allVals.length) break;

    // Читаем диапазон из CALC_COL (E) строки «Обложка»
    const rangeRaw = String(allVals[topIdx][CALC_COL - 1] || '').trim();
    if (!rangeRaw) continue;
    const range = parseDateRange_(rangeRaw);
    if (!range) continue;

    // Вычисляем прошлый период
    const prev = computePrevPeriod_(range.from, range.to);

    // Индексы дней прошлого периода
    const dayIdxs = [];
    for (let i = 0; i < nDays; i++) {
      const dk = dateKeys[i];
      if (dk && dk >= prev.from && dk <= prev.to) dayIdxs.push(i);
    }

    // Строка «Обложка» PREV_COL — показываем период
    const prevLabel = formatCalcRangeStr_(prev.from, prev.to);
    writes.push({ row: topIdx + 1, col: PREV_COL, value: prevLabel });

    if (!dayIdxs.length) continue;

    // Метрики
    for (let li = 1; li < BLOCK_H; li++) {
      const rowIdx = topIdx + li;
      const label  = String(allVals[rowIdx][FIRST_DATA_COL - 2] || '');
      if (MANUAL_LABELS.indexOf(label) >= 0) continue;
      if (!(label in METRIC_ROW_FIELD)) continue;

      const vals = dayIdxs
        .map(i => allVals[rowIdx][(FIRST_DATA_COL - 1) + i])
        .filter(v => v !== '' && v !== null && v !== undefined)
        .map(Number)
        .filter(n => !isNaN(n));

      if (!vals.length) { writes.push({ row: rowIdx + 1, col: PREV_COL, value: '' }); continue; }

      const sum    = vals.reduce((a, x) => a + x, 0);
      const result = SUM_LABELS.indexOf(label) >= 0 ? sum : sum / vals.length;
      writes.push({ row: rowIdx + 1, col: PREV_COL, value: Math.round(result * 100) / 100 });
    }
    updated++;
  }

  for (const w of writes) sh.getRange(w.row, w.col).setValue(w.value);
  return updated;
}

/**
 * Парсит строку диапазона дат в {from, to} в формате 'yyyy-MM-dd'.
 * Допустимые форматы:
 *   "01.06"              → один день
 *   "01.06-15.06"        → диапазон (текущий год)
 *   "01.06.2026-15.06.2026"
 *   Разделители: "-", "–", "—" (с пробелами или без)
 */
function parseDateRange_(raw) {
  const year = new Date().getFullYear();
  // Нормализуем разделители в один дефис
  const s = String(raw).replace(/\s*[–—]\s*/g, '-').replace(/\s*-\s*/g, '-').trim();

  // Ищем паттерны "дд.мм" или "дд.мм.гггг"
  const parts = s.match(/\d{1,2}\.\d{1,2}(?:\.\d{4})?/g);
  if (!parts || !parts.length) return null;

  function toKey(d) {
    const p = d.split('.');
    return `${p[2] || year}-${p[1].padStart(2, '0')}-${p[0].padStart(2, '0')}`;
  }

  const from = toKey(parts[0]);
  const to   = parts.length > 1 ? toKey(parts[1]) : from;
  return from <= to ? { from, to } : { from: to, to: from };
}
