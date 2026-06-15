/**
 * РЕНДЕР ЛИСТОВ КАБИНЕТОВ
 * ======================
 * Каждый кабинет — отдельный лист. Раскладка как в макете:
 *   A = Артикул, B = Наименование (объединены на всю высоту блока артикула),
 *   C = название строки, D..N = дни (по колонке на день).
 * Блок артикула: строка обложки (IMAGE по дням) + строки метрик + ручные строки.
 *
 * Рендер идемпотентен: лист полностью перерисовывается из _Данные,
 * но ручные строки (Гипотеза / Вид продвижения / Комментарий) сохраняются.
 */

function renderAllCabinets_(accounts, products) {
  const settings = alertSettings_();
  const todayKey = dateKey_(new Date());
  const displayDays = Number(getSetting_(SETTING_KEYS.DISPLAY_DAYS)) || 60;
  const data = readAllData_();

  // data -> {accountId: [records]}
  const byAccount = {};
  for (const rec of data) {
    (byAccount[rec.accountId] = byAccount[rec.accountId] || []).push(rec);
  }

  for (const account of accounts) {
    try {
      renderCabinet_(account,
        getProductsForAccount_(products, account.id),
        byAccount[account.id] || [],
        settings, todayKey, displayDays);
    } catch (e) {
      log_('ERROR', 'renderCabinet_', `Кабинет ${account.name}: ${e.message}\n${e.stack}`);
    }
  }
}


function renderCabinet_(account, prods, accountData, settings, todayKey, displayDays) {
  const ss = SpreadsheetApp.getActive();
  const sh = ensureSheet_(ss, account.name);

  // 1) Сохраняем ручные строки и диапазоны калькулятора до очистки
  const manualMap  = readManualValues_(sh);
  const calcRanges = readCalcRanges_(sh);   // {sku → строка диапазона}

  // 2) Определяем колонки-дни: уникальные даты кабинета, последние displayDays
  let dateKeys = Array.from(new Set(accountData.map(r => r.date)))
    .sort((a, b) => a.localeCompare(b));
  if (dateKeys.length > displayDays) dateKeys = dateKeys.slice(-displayDays);

  // data -> {sku: {dateKey: record}}  и  {sku: [records]}
  const bySku = {};
  for (const rec of accountData) {
    (bySku[rec.sku] = bySku[rec.sku] || []).push(rec);
  }

  // 3) Чистим лист и снимаем все merge
  sh.clear();
  if (sh.getMaxRows() > 1 && sh.getMaxColumns() > 1) {
    sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  }

  const nDays = dateKeys.length;
  const nCols = (FIRST_DATA_COL - 1) + nDays;
  const nRows = HEADER_ROWS + prods.length * BLOCK_H;

  // 4) Строим матрицу значений
  const values = [];
  for (let r = 0; r < nRows; r++) values.push(new Array(nCols).fill(''));

  // Заголовок: A=Артикул,B=Имя,C=Наименование,D=Прошлый период,E=Калькулятор,F=Период тестирования,G+=дни
  values[0][0] = 'Артикул';
  values[0][1] = 'Имя';
  values[0][2] = 'Наименование';
  values[0][3] = 'Прошлый период';
  values[0][4] = 'Калькулятор';
  values[0][5] = 'Период тестирования';
  dateKeys.forEach((dk, i) => { values[0][(FIRST_DATA_COL - 1) + i] = dateLabel_(dk); });

  // Анализ по каждому SKU (для подсветки)
  const skuAnalysis = {};

  prods.forEach((p, b) => {
    const top = HEADER_ROWS + b * BLOCK_H;   // 0-based индекс верхней строки блока
    const recsByDate = {};
    (bySku[p.sku] || []).forEach(r => { recsByDate[r.date] = r; });
    const recs = bySku[p.sku] || [];
    skuAnalysis[p.sku] = analyzeSku_(recs, settings, todayKey);

    // Наименование берём из карточки Ozon (самая свежая запись с именем),
    // откат на ручное из листа «Артикулы», если данных ещё нет.
    let cardName = '', maxDate = '';
    recs.forEach(r => { if (r.name && String(r.date) >= maxDate) { maxDate = String(r.date); cardName = r.name; } });

    // A=Артикул, B=Имя (менеджер), C=Наименование товара, D=диапазон калькулятора
    values[top][0] = p.sku;
    values[top][1] = p.manager || '';
    values[top][2] = cardName || p.name;
    values[top][CALC_COL - 1] = calcRanges[p.sku] || '';  // E: восстанавливаем диапазон калькулятора

    for (let li = 0; li < BLOCK_H; li++) {
      const row = top + li;
      const label = ROW_LABELS[li];
      values[row][5] = label;  // col F (0-based index 5)

      for (let i = 0; i < nDays; i++) {
        const col = (FIRST_DATA_COL - 1) + i;
        const dk = dateKeys[i];
        const rec = recsByDate[dk];

        if (li === IMAGE_ROW_OFFSET) {
          // строка обложки
          // без второго аргумента: режим по умолчанию (вписать в ячейку),
          // и не зависит от локали (запятая/точка с запятой как разделитель)
          if (rec && rec.coverUrl) values[row][col] = `=IMAGE("${rec.coverUrl}")`;
        } else if (label in METRIC_ROW_FIELD) {
          if (rec) {
            const v = rec[METRIC_ROW_FIELD[label]];
            values[row][col] = (v === '' || v === null || v === undefined) ? '' : Number(v);
          }
        } else if (MANUAL_LABELS.indexOf(label) >= 0) {
          const key = p.sku + '||' + label + '||' + dk;
          if (manualMap[key] !== undefined && manualMap[key] !== '') values[row][col] = manualMap[key];
        }
      }
    }
  });

  // 5) Расширяем лист под нужный размер и пишем всё разом
  ensureDimensions_(sh, nRows, nCols);
  sh.getRange(1, 1, nRows, nCols).setValues(values);
  sh.getRange(1, 1, nRows, nCols).setFontFamily('Onest').setVerticalAlignment('middle');

  // Запоминаем dateKeys в невидимых метаданных листа (не в примечаниях — они видны при наведении)
  saveDateKeys_(sh, dateKeys);

  // Очищаем старые примечания с дат (если остались от предыдущей версии)
  if (nDays > 0) {
    sh.getRange(1, FIRST_DATA_COL, 1, nDays).clearNote();
  }

  // 6) Форматирование
  formatHeader_(sh, nCols);
  sh.setFrozenRows(HEADER_ROWS);
  // фиксируем A:C только если справа есть хотя бы одна колонка-день
  if (nCols > FIRST_DATA_COL - 1) sh.setFrozenColumns(FIRST_DATA_COL - 1);
  sh.setColumnWidth(1, 120);   // A: Артикул
  sh.setColumnWidth(2, 90);    // B: Имя (менеджер)
  sh.setColumnWidth(3, 200);   // C: Наименование
  sh.setColumnWidth(4, 110);   // D: Прошлый период
  sh.setColumnWidth(5, 110);   // E: Калькулятор
  sh.setColumnWidth(6, 150);   // F: Метрика
  for (let i = 0; i < nDays; i++) sh.setColumnWidth(FIRST_DATA_COL + i, DAY_COL_WIDTH);

  prods.forEach((p, b) => {
    const topRow1 = HEADER_ROWS + b * BLOCK_H + 1;   // 1-based строка обложки
    formatBlock_(sh, topRow1, nDays, skuAnalysis[p.sku], dateKeys);
  });

  // Группируем строки метрик каждого блока (строка обложки остаётся видимой)
  applyRowGroups_(sh, prods.length);

  // Градиентная заливка строк CTR и Рекламный CTR
  if (nDays > 0) applyMetricColorScales_(sh, prods, nDays);

  log_('INFO', 'renderCabinet_',
    `Кабинет «${account.name}»: артикулов ${prods.length}, дней ${nDays}`);
}


/**
 * Форматирует блок одного артикула: merge A/B, высота обложки,
 * форматы чисел, границы, подсветка смен обложки и «залежалости».
 */
function formatBlock_(sh, topRow, nDays, analysis, dateKeys) {
  const lastCol = (FIRST_DATA_COL - 1) + nDays;

  // merge A (Артикул), B (Имя), C (Наименование) на всю высоту блока
  sh.getRange(topRow, 1, BLOCK_H, 1).merge().setVerticalAlignment('middle').setWrap(true).setFontWeight('bold');
  sh.getRange(topRow, 2, BLOCK_H, 1).merge().setVerticalAlignment('middle').setHorizontalAlignment('center').setWrap(true);
  sh.getRange(topRow, 3, BLOCK_H, 1).merge().setVerticalAlignment('middle').setWrap(true);

  // Батч-форматирование D (Прошлый период) и E (Калькулятор) за 2 вызова
  const numFmts = ROW_LABELS.map(lbl => {
    if (PERCENT_LABELS.indexOf(lbl) >= 0) return ['0.00"%"'];
    if (lbl in METRIC_ROW_FIELD)          return ['#,##0'];
    return ['@'];   // текст для Гипотезы/Комментария/Обложки
  });
  numFmts[IMAGE_ROW_OFFSET] = ['@'];  // строка Обложки — всегда текст

  sh.getRange(topRow, PREV_COL, BLOCK_H, 1)
    .setBackground('#f1f3f4').setHorizontalAlignment('center')
    .setFontColor('#555555').setWrap(false)
    .setNumberFormats(numFmts);
  sh.getRange(topRow + IMAGE_ROW_OFFSET, PREV_COL).setFontStyle('italic');

  sh.getRange(topRow, CALC_COL, BLOCK_H, 1)
    .setBackground('#e8f0fe').setHorizontalAlignment('center')
    .setFontColor('#000000').setWrap(false)
    .setNumberFormats(numFmts);
  sh.getRange(topRow + IMAGE_ROW_OFFSET, CALC_COL).setFontStyle('italic');

  // колонка F — названия строк (метрик)
  sh.getRange(topRow, 6, BLOCK_H, 1).setFontWeight('bold').setBackground('#f3f3f3').setFontColor('#000000');

  // высота строки обложки
  sh.setRowHeight(topRow + IMAGE_ROW_OFFSET, IMAGE_ROW_HEIGHT);
  if (nDays > 0) {
    sh.getRange(topRow + IMAGE_ROW_OFFSET, FIRST_DATA_COL, 1, nDays).setHorizontalAlignment('center');
  }

  // форматы чисел по строкам метрик
  for (let li = 0; li < BLOCK_H; li++) {
    const label = ROW_LABELS[li];
    if (nDays === 0) break;
    const rng = sh.getRange(topRow + li, FIRST_DATA_COL, 1, nDays);
    if (PERCENT_LABELS.indexOf(label) >= 0) {
      rng.setNumberFormat('0.00"%"').setHorizontalAlignment('center');
    } else if (label in METRIC_ROW_FIELD) {
      rng.setNumberFormat('#,##0').setHorizontalAlignment('center');
    }
  }

  // нижняя граница блока
  sh.getRange(topRow + BLOCK_H - 1, 1, 1, lastCol).setBorder(
    null, null, true, null, null, null, '#666666', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Подсветка дней смены обложки: рамка вокруг ячейки обложки + заголовок дня
  if (nDays > 0 && analysis) {
    dateKeys.forEach((dk, i) => {
      if (analysis.changeDateKeys[dk]) {
        const col = FIRST_DATA_COL + i;
        sh.getRange(topRow + IMAGE_ROW_OFFSET, col)
          .setBorder(true, true, true, true, false, false, '#e69138', SpreadsheetApp.BorderStyle.SOLID_THICK);
        sh.getRange(1, col).setBackground('#fce5cd').setFontColor('#000000'); // подсветка заголовка-даты
      }
    });

    // «Залежавшаяся» обложка — красим артикул
    if (analysis.stale) {
      sh.getRange(topRow, 1).setBackground('#f4cccc');
    }
  }
}


/**
 * Группирует строки метрик/ручного ввода каждого блока (всё, кроме строки обложки),
 * чтобы их можно было сворачивать. Строка обложки остаётся видимой как «шапка» блока.
 * Идемпотентно: если группа уже есть — не трогает (сохраняет свёрнутость).
 */
function applyRowGroups_(sh, nBlocks) {
  for (let b = 0; b < nBlocks; b++) {
    const blockStart = HEADER_ROWS + b * BLOCK_H + 1;           // 1-based: строка обложки
    const groupStart = blockStart + GROUP_START_OFFSET;         // первая сворачиваемая строка
    const count = BLOCK_H - GROUP_START_OFFSET;

    // Уже есть группа ровно на нужной границе? — не трогаем (сохраняем свёрнутость)
    let ok = false;
    try {
      const g = sh.getRowGroup(groupStart, 1);
      if (g) {
        const rng = g.getRange();
        if (rng.getRow() === groupStart && rng.getNumRows() === count) ok = true;
      }
    } catch (e) { /* нет группы */ }

    if (!ok) {
      // Снимаем любые старые группы в пределах блока (по строке, безопасно)
      for (let r = blockStart; r < blockStart + BLOCK_H; r++) {
        try { sh.getRange(r, 1).shiftRowGroupDepth(-1); } catch (_) {}
      }
      try { sh.getRange(groupStart, 1, count, 1).shiftRowGroupDepth(1); } catch (_) {}
    }
  }
}

/**
 * Свернуть (collapse=true) или развернуть все блоки на всех листах кабинетов.
 * Возвращает число обработанных блоков.
 */
function setAllGroups_(collapse) {
  const ss = SpreadsheetApp.getActive();
  let count = 0;
  ss.getSheets().forEach(sh => {
    if (!isCabinetSheet_(sh.getName())) return; // только листы кабинетов
    const lastRow = sh.getLastRow();
    for (let b = 0; ; b++) {
      const groupStart = HEADER_ROWS + b * BLOCK_H + 1 + GROUP_START_OFFSET;
      if (groupStart > lastRow) break;
      try {
        const g = sh.getRowGroup(groupStart, 1);
        if (g) { collapse ? g.collapse() : g.expand(); count++; }
      } catch (e) { /* нет группы в этой строке — пропускаем */ }
    }
  });
  return count;
}


/** Гарантирует, что на листе хватает строк и колонок под запись nRows x nCols. */
function ensureDimensions_(sh, nRows, nCols) {
  const maxR = sh.getMaxRows();
  const maxC = sh.getMaxColumns();
  if (nRows > maxR) sh.insertRowsAfter(maxR, nRows - maxR);
  if (nCols > maxC) sh.insertColumnsAfter(maxC, nCols - maxC);
}


/**
 * Читает ручные строки (Гипотеза / Вид продвижения / Комментарий) из текущего листа.
 * Ключ: sku || label || dateKey. dateKey берём из примечаний к заголовкам-датам.
 */
function readManualValues_(sh) {
  const map = {};
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < HEADER_ROWS + 1 || lastCol < FIRST_DATA_COL) return map;

  const nDays = lastCol - FIRST_DATA_COL + 1;
  const dayKeys = loadDateKeys_(sh, nDays); // dateKeys из метаданных листа
  const values  = sh.getRange(1, 1, lastRow, lastCol).getValues();

  for (let top = HEADER_ROWS; top + BLOCK_H - 1 < lastRow + 1 && top + BLOCK_H <= values.length; top += BLOCK_H) {
    const sku = String(values[top][0] || '').trim();
    if (!sku) continue;
    for (let li = 0; li < BLOCK_H; li++) {
      const label = values[top + li][5];  // col F (0-based index 5) — названия строк
      if (MANUAL_LABELS.indexOf(label) < 0) continue;
      for (let i = 0; i < nDays; i++) {
        const dk = dayKeys[i];
        if (!dk) continue;
        const val = values[top + li][(FIRST_DATA_COL - 1) + i];
        if (val !== '' && val !== null && val !== undefined) {
          map[sku + '||' + label + '||' + dk] = val;
        }
      }
    }
  }
  return map;
}


// ── Хранение dateKeys в невидимых метаданных листа ──────────────────────────
// Developer Metadata не показывается в UI (нет всплывающих подсказок).

const DATE_KEYS_META = 'ozon_dateKeys';

/** Сохраняет массив dateKeys в метаданных листа. */
function saveDateKeys_(sh, dateKeys) {
  // Удаляем старую запись (если есть), затем пишем новую
  sh.getDeveloperMetadata()
    .filter(m => m.getKey() === DATE_KEYS_META)
    .forEach(m => m.remove());
  sh.addDeveloperMetadata(
    DATE_KEYS_META,
    JSON.stringify(dateKeys),
    SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT
  );
}

/**
 * Градиентная заливка строк CTR и Рекламный CTR по всем блокам.
 * Красный → жёлтый → зелёный по percentile внутри каждой метрики.
 * Одно правило на метрику = все блоки сравниваются между собой.
 */
function applyMetricColorScales_(sh, prods, nDays) {
  const ctrRanges   = [];
  const adCtrRanges = [];

  prods.forEach((p, b) => {
    const topRow = HEADER_ROWS + b * BLOCK_H + 1;  // 1-based
    ROW_LABELS.forEach((label, li) => {
      const rng = sh.getRange(topRow + li, FIRST_DATA_COL, 1, nDays);
      if (label === 'CTR')           ctrRanges.push(rng);
      else if (label === 'Рекламный CTR') adCtrRanges.push(rng);
    });
  });

  const P = SpreadsheetApp.InterpolationType.PERCENTILE;
  const rules = [];

  [ctrRanges, adCtrRanges].forEach(ranges => {
    if (!ranges.length) return;
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .setGradientMinpointWithValue('#ea4335', P, '10')   // красный  — нижние 10%
        .setGradientMidpointWithValue('#fbbc04', P, '50')   // жёлтый   — медиана
        .setGradientMaxpointWithValue('#34a853', P, '90')   // зелёный  — верхние 10%
        .setRanges(ranges)
        .build()
    );
  });

  sh.setConditionalFormatRules(rules);
}


/** Читает массив dateKeys из метаданных листа. При отсутствии — возвращает массив пустых строк. */
function loadDateKeys_(sh, nDays) {
  const meta = sh.getDeveloperMetadata().find(m => m.getKey() === DATE_KEYS_META);
  if (meta) {
    try {
      const keys = JSON.parse(meta.getValue());
      if (Array.isArray(keys)) return keys;
    } catch (e) { /* повреждённые метаданные — игнорируем */ }
  }
  return new Array(nDays).fill('');
}
