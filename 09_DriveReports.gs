/**
 * ЧТЕНИЕ РЕКЛАМНЫХ ОТЧЁТОВ ИЗ GOOGLE DRIVE
 * ==========================================
 * Рекламный CTR берётся из ежедневных Excel-отчётов,
 * которые загружаются в Google Drive вместо нестабильного async-ZIP Performance API.
 *
 * Ожидаемая структура папки DRIVE_REPORTS_FOLDER_ID:
 *   └── 04.06/          (подпапка на каждый день, формат dd.MM)
 *         ├── *ИП*.xlsx  (отчёт кабинета ИП)
 *         └── *ООО*.xlsx (отчёт кабинета ООО)
 *
 * В каждом xlsx: строка заголовка (ищем "Артикул"/"offer_id" и "CTR"),
 * далее строки по товарам. CTR в %, привязанный к offer_id.
 *
 * Интеграция: fetchAllAdStatsForAccount_() сначала пробует Drive,
 * при неудаче — падает обратно на Performance API.
 */

// ID главной папки Google Drive с рекламными отчётами.
// Взят из URL: drive.google.com/drive/folders/<ID>
const DRIVE_REPORTS_FOLDER_ID = '1bVNYhoQnClt3bzCd41dJKiNCNJjkjfx4';


/**
 * Читает рекламный CTR из Excel-отчёта в Drive для указанного кабинета и даты.
 * Возвращает объект {offer_id → {adViews, adClicks, ctr}} совместимый с Performance API.
 *
 * Возвращает пустой объект {} если:
 *  - папка для даты не найдена
 *  - файл кабинета не найден
 *  - файл пустой или не содержит нужных колонок
 *
 * @param {Object} account    — объект аккаунта (нужны account.name)
 * @param {Date}   metricDate — дата, за которую берём CTR
 */
function readAdCtrFromDriveReport_(account, metricDate) {
  const tz       = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const dayMonth = Utilities.formatDate(metricDate, tz, 'dd.MM');      // «04.06»
  const dateStr  = Utilities.formatDate(metricDate, tz, 'yyyy-MM-dd'); // «2026-06-04»

  try {
    // 1. Главная папка Drive
    const mainFolder = DriveApp.getFolderById(DRIVE_REPORTS_FOLDER_ID);

    // 2. Подпапка с датой
    const dateFolder = findFolderByDate_(mainFolder, dayMonth, dateStr);
    if (!dateFolder) {
      log_('DEBUG', 'readAdCtrFromDrive_',
        `Папка для даты ${dayMonth} / ${dateStr} не найдена в Drive`);
      return {};
    }

    // 3. Файл кабинета в подпапке
    const xlsxFile = findCabinetFile_(dateFolder, account.name);
    if (!xlsxFile) {
      log_('DEBUG', 'readAdCtrFromDrive_',
        `Файл кабинета «${account.name}» не найден в папке ${dateFolder.getName()}`);
      return {};
    }
    log_('DEBUG', 'readAdCtrFromDrive_',
      `Кабинет «${account.name}»: файл ${xlsxFile.getName()}`);

    // 4. Читаем содержимое (xlsx → Google Sheet → 2D-массив)
    const data = readFileAsArray_(xlsxFile);
    if (!data || !data.length) return {};

    // 5. Определяем колонки по заголовкам
    const { skuCol, viewsCol, clicksCol, startRow } = detectColumns_(data);
    log_('DEBUG', 'readAdCtrFromDrive_',
      `Колонки: артикул=${colLetter_(skuCol)}, просмотры=${colLetter_(viewsCol)}, ` +
      `клики=${colLetter_(clicksCol)}, данные с строки ${startRow + 1}`);

    // 6. Агрегируем просмотры и клики по SKU (один SKU может быть в нескольких кампаниях)
    // CTR = сумма(Клики) / сумма(Просмотры) × 100  — точнее чем среднее по строкам
    const agg = {};  // sku → {views, clicks}
    for (let r = startRow; r < data.length; r++) {
      const row = data[r];
      if (!row) continue;
      const sku = String(row[skuCol] || '').trim();
      if (!sku) continue;

      const views  = parseNum_(row[viewsCol]);
      const clicks = parseNum_(row[clicksCol]);

      if (!agg[sku]) agg[sku] = { views: 0, clicks: 0 };
      agg[sku].views  += views;
      agg[sku].clicks += clicks;
    }

    // 7. Вычисляем CTR из агрегатов
    const result = {};
    for (const sku of Object.keys(agg)) {
      const { views, clicks } = agg[sku];
      const ctr = views > 0 ? Math.round((clicks / views) * 10000) / 100 : 0;  // 2 знака
      result[sku] = { adViews: views, adClicks: clicks, ctr };
    }

    log_('DEBUG', 'readAdCtrFromDrive_',
      `Итого SKU с рекл. CTR: ${Object.keys(result).length}`);
    return result;

  } catch (e) {
    log_('WARN', 'readAdCtrFromDrive_',
      `Ошибка чтения Drive отчёта (кабинет ${account.name}, дата ${dayMonth}): ${e.message}`);
    return {};
  }
}


// ════════════════════════════════════════════════════════════════
//  РЕТРОСПЕКТИВНОЕ ЗАПОЛНЕНИЕ РЕКЛАМНОГО CTR
// ════════════════════════════════════════════════════════════════

/**
 * Заполняет поле «Рекламный CTR» в _Данные из Drive-отчётов за последние nDays дней.
 * НЕ затрагивает другие метрики (CTR, показы и т.д.) — только adCtr.
 * Быстро: нет обращений к Seller API, только Drive.
 *
 * Для каждого дня ищет файл в Drive-папке (папка dd.MM в DRIVE_REPORTS_FOLDER_ID),
 * читает CTR per SKU и делает частичный upsert в _Данные.
 * После завершения перерисовывает листы кабинетов.
 */
function backfillAdCtr_(nDays) {
  const accounts = getAccounts_();
  const products  = getActiveProducts_();
  if (!accounts.length || !products.length) {
    SpreadsheetApp.getUi().alert('Нет активных аккаунтов или артикулов.');
    return;
  }

  let filled = 0, noFile = 0;

  for (let d = 1; d <= nDays; d++) {
    const metricDate = addDays_(new Date(), -d);
    const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
    const dayStr = Utilities.formatDate(metricDate, tz, 'dd.MM');

    // Читаем Drive-отчёт для каждого аккаунта
    const adByAccount = {};
    for (const acc of accounts) {
      adByAccount[acc.id] = readAdCtrFromDriveReport_(acc, metricDate);
    }

    // Формируем частичные записи: только accountId + sku + date + adCtr
    const records = [];
    for (const product of products) {
      const account = accounts.find(a => a.id === product.accountId);
      if (!account) continue;
      const adStats = (adByAccount[account.id] || {})[product.sku] || null;
      if (!adStats) {
        noFile++;
        continue;
      }
      records.push({
        accountId: account.id,
        sku:       product.sku,
        date:      dateKey_(metricDate),
        adCtr:     adStats.ctr
      });
      filled++;
      log_('DEBUG', 'backfillAdCtr_',
        `${dayStr} SKU=${product.sku}: Рекл.CTR=${adStats.ctr}%`);
    }

    if (records.length) mergeUpsertRecords_(records);
  }

  renderAllCabinets_(accounts, products);
  flushLog_();
  SpreadsheetApp.getActive().toast(
    `Рекл. CTR: заполнено ${filled} записей за ${nDays} дн., ` +
    `не найдено в Drive: ${noFile} SKU×дней`,
    '📊 Готово', 6
  );
}

/**
 * Заполняет рекламный CTR за вчера из Drive — вызывается ТРИГГЕРОМ автоматически.
 * Без UI-вызовов (alert/toast), только логирование — иначе триггер упадёт.
 */
function fillAdCtrYesterdayTrigger() {
  try {
    const accounts = getAccounts_();
    const products  = getActiveProducts_();
    if (!accounts.length || !products.length) {
      log_('WARN', 'fillAdCtrYesterdayTrigger', 'Нет активных аккаунтов или артикулов');
      return;
    }

    const metricDate = addDays_(new Date(), -1);
    const records    = [];

    for (const account of accounts) {
      const adMap      = readAdCtrFromDriveReport_(account, metricDate);
      const accProducts = products.filter(p => p.accountId === account.id);
      for (const product of accProducts) {
        const adStats = adMap[product.sku] || null;
        if (!adStats) continue;
        records.push({
          accountId: account.id,
          sku:       product.sku,
          date:      dateKey_(metricDate),
          adCtr:     adStats.ctr
        });
      }
    }

    if (records.length) {
      mergeUpsertRecords_(records);
      renderAllCabinets_(accounts, products);
      log_('INFO', 'fillAdCtrYesterdayTrigger',
        `Рекл. CTR: заполнено ${records.length} записей за ${dateKey_(metricDate)}`);
    } else {
      log_('WARN', 'fillAdCtrYesterdayTrigger',
        `Рекл. CTR: нет данных — Drive-папка ещё не готова или файл не найден`);
    }
  } catch (e) {
    log_('FATAL', 'fillAdCtrYesterdayTrigger', e.message + '\n' + e.stack);
  } finally {
    flushLog_();
  }
}

function backfillAdCtrMenu() {
  const ui  = SpreadsheetApp.getUi();
  const res = ui.prompt(
    '📊 Заполнить Рекл. CTR из Drive',
    'За сколько последних дней заполнить? (1–60)\n' +
    'Для каждого дня нужен соответствующий файл в папке Drive.\n' +
    'Другие метрики (CTR, показы и т.д.) не изменятся.',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const n = parseInt(res.getResponseText(), 10);
  if (!n || n < 1 || n > 60) { ui.alert('Введи число от 1 до 60.'); return; }
  backfillAdCtr_(n);
}


// ════════════════════════════════════════════════════════════════
//  ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
// ════════════════════════════════════════════════════════════════

/**
 * Ищет подпапку для даты в родительской папке.
 * Пробует несколько форматов названия: «04.06», «2026-06-04», «04-06», «04.06.2026».
 */
function findFolderByDate_(parentFolder, dayMonth, dateStr) {
  const [year, mon, day] = dateStr.split('-');

  // Точные варианты названия папки
  const exact = [
    dayMonth,                          // «04.06»
    dateStr,                           // «2026-06-04»
    `${day}-${mon}`,                   // «04-06»
    `${day}.${mon}.${year}`,           // «04.06.2026»
    `${day} ${mon}`,                   // «04 06»
  ];

  for (const name of exact) {
    const iter = parentFolder.getFoldersByName(name);
    if (iter.hasNext()) return iter.next();
  }

  // Мягкий поиск: имя папки содержит dayMonth или dateStr
  const iter = parentFolder.getFolders();
  while (iter.hasNext()) {
    const f = iter.next();
    const fn = f.getName();
    if (fn.includes(dayMonth) || fn.includes(dateStr) || fn.includes(`${day}.${mon}`)) {
      return f;
    }
  }
  return null;
}


/**
 * Находит Excel/Sheets-файл для кабинета в папке.
 * Выбирает файл по ключевому слову ИП/ООО из названия кабинета.
 */
function findCabinetFile_(folder, accountName) {
  const CABINET_KEYWORDS = ['ООО', 'ИП'];  // порядок важен: ООО проверяем первым
  const upperName = accountName.toUpperCase();

  // Определяем ключевое слово кабинета
  let keyword = '';
  for (const kw of CABINET_KEYWORDS) {
    if (upperName.includes(kw)) { keyword = kw; break; }
  }

  const ALLOWED_MIME = [
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // .xlsx
    'application/vnd.ms-excel',                                           // .xls
    'text/csv',                                                            // .csv
    'application/vnd.google-apps.spreadsheet'                             // Google Sheet
  ];

  const fileIter = folder.getFiles();
  let fallback = null;
  while (fileIter.hasNext()) {
    const f = fileIter.next();
    const mime  = f.getMimeType();
    const fname = f.getName();
    const isReport = ALLOWED_MIME.includes(mime) ||
                     fname.match(/\.(xlsx|xls|csv)$/i);
    if (!isReport) continue;

    if (keyword && fname.toUpperCase().includes(keyword)) return f;
    if (!fallback) fallback = f;  // первый подходящий файл — запасной вариант
  }
  return fallback;  // null если ничего не нашлось
}


/**
 * Читает файл (xlsx, xls, csv или Google Sheet) из Drive как 2D-массив.
 *
 * Для xlsx/xls: конвертирует в Google Sheet через Drive REST API,
 * читает первый лист, удаляет временный файл.
 * Для Google Sheet и CSV: читает напрямую.
 */
function readFileAsArray_(driveFile) {
  const mime = driveFile.getMimeType();

  // Google Sheet — читаем напрямую
  if (mime === 'application/vnd.google-apps.spreadsheet') {
    return SpreadsheetApp.openById(driveFile.getId())
      .getSheets()[0].getDataRange().getValues();
  }

  // CSV — читаем как текст и парсим
  if (mime === 'text/csv' || driveFile.getName().endsWith('.csv')) {
    return Utilities.parseCsv(driveFile.getBlob().getDataAsString('UTF-8'));
  }

  // xlsx / xls — конвертируем в Google Sheet через Drive REST API.
  // ScriptApp.getOAuthToken() даёт токен с полными Drive-правами (DriveApp используется выше).
  const token   = ScriptApp.getOAuthToken();
  const copyUrl = 'https://www.googleapis.com/drive/v3/files/' +
                  driveFile.getId() + '/copy';

  log_('DEBUG', 'readFileAsArray_',
    `Конвертация xlsx→Sheet: ${driveFile.getName()} (${driveFile.getId()})`);

  const copyResp = UrlFetchApp.fetch(copyUrl, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + token },
    payload: JSON.stringify({
      name:     '_tmp_ctr_' + driveFile.getId(),
      mimeType: 'application/vnd.google-apps.spreadsheet'
    }),
    muteHttpExceptions: true
  });

  log_('DEBUG', 'readFileAsArray_',
    `Drive copy HTTP ${copyResp.getResponseCode()}: ${copyResp.getContentText().slice(0, 200)}`);

  if (copyResp.getResponseCode() !== 200) {
    throw new Error(
      `Drive copy HTTP ${copyResp.getResponseCode()}: ${copyResp.getContentText().slice(0, 200)}`
    );
  }

  const tempId = JSON.parse(copyResp.getContentText()).id;
  log_('DEBUG', 'readFileAsArray_', `Временный лист создан, ID=${tempId}`);

  try {
    // Даём Drive 4 с на конвертацию (xlsx→Sheets — асинхронно на стороне Google)
    Utilities.sleep(4000);

    // Открываем с 3 попытками на случай медленной конвертации
    let ss = null;
    let lastErr = '';
    for (let attempt = 1; attempt <= 3; attempt++) {
      try {
        ss = SpreadsheetApp.openById(tempId);
        log_('DEBUG', 'readFileAsArray_', `Открыт с попытки ${attempt}`);
        break;
      } catch (e) {
        lastErr = e.message;
        log_('DEBUG', 'readFileAsArray_', `Попытка ${attempt} открыть: ${lastErr}`);
        if (attempt < 3) Utilities.sleep(3000);
      }
    }
    if (!ss) throw new Error(`Не удалось открыть временный лист: ${lastErr}`);

    const sheet = ss.getSheets()[0];
    const data  = sheet.getDataRange().getValues();
    log_('DEBUG', 'readFileAsArray_', `Прочитано строк: ${data.length}, столбцов: ${data[0] ? data[0].length : 0}`);
    return data;

  } finally {
    // Всегда удаляем временный лист
    try {
      UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + tempId, {
        method:  'delete',
        headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
        muteHttpExceptions: true
      });
      log_('DEBUG', 'readFileAsArray_', `Временный лист удалён (${tempId})`);
    } catch (e) {
      log_('DEBUG', 'readFileAsArray_', `Удаление временного листа: ${e.message}`);
    }
  }
}


/**
 * Определяет позиции ключевых колонок по заголовочным строкам файла.
 * Возвращает {skuCol, viewsCol, clicksCol, startRow} (0-indexed).
 *
 * Заголовки Ozon Performance-отчёта (ru):
 *   A:Фото | B:Артикул продавца | C:Кампании | D:Типы |
 *   E:Стоимость 1000 показ | F:Просмотры | G:CTR | H:Клики | ...
 *
 * Умолчания соответствуют этой структуре:
 *   skuCol=1 (B), viewsCol=5 (F), clicksCol=7 (H).
 */
function detectColumns_(data) {
  let skuCol    = 1;   // B — «Артикул продавца»
  let viewsCol  = 5;   // F — «Просмотры»
  let clicksCol = 7;   // H — «Клики»
  let startRow  = 1;   // строка 2 по умолчанию

  for (let r = 0; r < Math.min(10, data.length); r++) {
    const row = data[r];
    let foundSku = false;

    for (let c = 0; c < row.length; c++) {
      const cell = String(row[c] || '').toLowerCase().trim();
      if (!cell) continue;

      // Артикул / SKU
      if (
        cell === 'артикул' || cell === 'offer_id' || cell === 'sku' ||
        cell.startsWith('артикул') || cell === 'артикул товара'
      ) {
        skuCol   = c;
        startRow = r + 1;
        foundSku = true;
      }

      // Просмотры / Показы / Impressions / Views
      if (
        cell === 'просмотры' || cell === 'показы' ||
        cell === 'impressions' || cell === 'views' ||
        cell.startsWith('просмотр')
      ) {
        viewsCol = c;
      }

      // Клики / Clicks
      if (cell === 'клики' || cell === 'clicks' || cell === 'click') {
        clicksCol = c;
      }
    }

    if (foundSku) break;
  }

  return { skuCol, viewsCol, clicksCol, startRow };
}


/** Парсит число из ячейки (уже число, строка с запятой/точкой, пустое → 0). */
function parseNum_(val) {
  if (typeof val === 'number') return val;
  return parseFloat(String(val || '0').replace(',', '.').trim()) || 0;
}


/** Буква столбца по 0-based индексу: 0→A, 6→G и т.д. */
function colLetter_(idx) {
  let s = '';
  let n = idx;
  do {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return s;
}
