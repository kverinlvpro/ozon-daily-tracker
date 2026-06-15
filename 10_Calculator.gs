/**
 * КАЛЬКУЛЯТОР ПЕРИОДА
 * ===================
 * Пользователь вводит диапазон дат в ячейку D строки «Обложка» каждого блока
 * (формат: "01.06-15.06" или "01.06–15.06") и нажимает меню.
 *
 * Функция считает по выбранным дням:
 *   — среднее для % метрик (CTR, Рекл.CTR, Кр в корзину, CR корзина→заказ)
 *   — сумму   для абсолютных (Показы, Заказы, Выкупы)
 *
 * Ячейки дней (F+) НЕ ТРОГАЮТСЯ.
 */

// Метрики, для которых считается сумма (остальные — среднее)
const SUM_LABELS = ['Показы', 'Заказы', 'Выкупы'];

/** Пункт меню: пересчитать калькулятор для всех кабинетных листов */
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
      const label  = String(allVals[rowIdx][FIRST_DATA_COL - 2] || ''); // col E = index 4

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
 * Читает сохранённые диапазоны из строк «Обложка» колонки D текущего листа.
 * Вызывается ДО sh.clear() в renderCabinet_, чтобы не потерять введённые даты.
 * Возвращает {sku → строка диапазона}.
 */
function readCalcRanges_(sh) {
  const map    = {};
  const lastRow = sh.getLastRow();
  if (lastRow < HEADER_ROWS + 1) return map;

  // Читаем только col A (SKU) и col D (диапазон)
  const data = sh.getRange(1, 1, lastRow, CALC_COL).getValues();

  for (let b = 0; ; b++) {
    const topIdx = HEADER_ROWS + b * BLOCK_H;
    if (topIdx >= data.length) break;
    const sku = String(data[topIdx][0] || '').trim();
    if (!sku) continue;
    const val = String(data[topIdx][CALC_COL - 1] || '').trim();
    if (val && parseDateRange_(val)) map[sku] = val;   // сохраняем только если парсится
  }
  return map;
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
