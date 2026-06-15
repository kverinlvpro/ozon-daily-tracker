/**
 * ИНИЦИАЛИЗАЦИЯ ТАБЛИЦЫ
 * =====================
 * Запускать ОДИН РАЗ при первом развёртывании (меню → 🆕 Инициализировать).
 * Повторный запуск безопасен: сохраняет уже введённые настройки/аккаунты/артикулы.
 */

function initSpreadsheet() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  // ===== Настройки =====
  let s = ensureSheet_(ss, SHEET.SETTINGS);
  const savedSettings = {};
  if (s.getLastRow() > 1) {
    s.getRange(2, 1, s.getLastRow() - 1, 2).getValues()
      .forEach(([k, v]) => { if (k) savedSettings[k] = v; });
  }
  s.clear();
  s.getRange(1, 1, 1, 3).setValues([['Ключ', 'Значение', 'Описание']]);
  const defaults = [
    [SETTING_KEYS.LAG_DAYS,     1,  'За какой день брать метрики (1 = вчера; аналитика Ozon с задержкой ~1 день)'],
    [SETTING_KEYS.UPDATE_HOUR,  14, 'Час ежедневного автообновления (0–23, таймзона Екатеринбург UTC+5; 14 = 12:00 МСК)'],
    [SETTING_KEYS.DISPLAY_DAYS, 60, 'Сколько последних дней показывать колонками в листах кабинетов'],
    [SETTING_KEYS.STALE_DAYS,   14, 'Предупреждать, если обложка не менялась дольше стольких дней'],
    [SETTING_KEYS.ALERT_PCT,    10, 'Порог изменения конверсии после смены обложки для предупреждения, %'],
    [SETTING_KEYS.BEFORE_DAYS,  7,  'Окно ДО смены обложки для сравнения, дней'],
    [SETTING_KEYS.AFTER_DAYS,   7,  'Окно ПОСЛЕ смены обложки для сравнения, дней']
  ];
  const toWrite = defaults.map(([k, def, desc]) => [k, k in savedSettings ? savedSettings[k] : def, desc]);
  s.getRange(2, 1, toWrite.length, 3).setValues(toWrite);
  formatHeader_(s, 3);
  s.setColumnWidth(1, 200);
  s.setColumnWidth(2, 90);
  s.setColumnWidth(3, 560);

  // ===== Аккаунты =====
  s = ensureSheet_(ss, SHEET.ACCOUNTS);
  const savedAccounts = s.getLastRow() > 1 ? s.getRange(2, 1, s.getLastRow() - 1, 6).getValues() : null;
  s.clear();
  s.getRange(1, 1, 1, 6).setValues([[
    'ID аккаунта', 'Название', 'Seller Client-Id', 'Seller Api-Key',
    'Performance: client_id|secret', 'Активен'
  ]]);
  if (savedAccounts && savedAccounts.some(r => r[0])) {
    s.getRange(2, 1, savedAccounts.length, 6).setValues(savedAccounts);
  } else {
    s.getRange(2, 1, 1, 6).setValues([['acc1', 'Кабинет 1', '', '', '', true]]);
  }
  formatHeader_(s, 6);
  s.setColumnWidth(2, 160);
  s.setColumnWidth(4, 280);
  s.setColumnWidth(5, 280);

  // ===== Листы артикулов (по одному на кабинет) =====
  // Читаем аккаунты из только что сохранённого листа (savedAccounts может быть null при первом запуске)
  const currentAccounts = getAccounts_();

  // Миграция: если есть старый лист «Артикулы» с данными → перекладываем в новые листы
  const oldProductSheet = ss.getSheetByName(SHEET.PRODUCTS);
  const oldRows = oldProductSheet && oldProductSheet.getLastRow() > 1
    ? oldProductSheet.getRange(2, 1, oldProductSheet.getLastRow() - 1, 6).getValues()
    : [];
  // Старые колонки: A=SKU, B=Название, C=ID аккаунта, D=Менеджер, E=Статус
  const migrationByAccount = {};  // accountId → [[sku, name, manager, status]]
  for (const r of oldRows) {
    if (!r[0]) continue;
    const accId = String(r[2]);
    if (!migrationByAccount[accId]) migrationByAccount[accId] = [];
    migrationByAccount[accId].push([String(r[0]), String(r[1]), String(r[3]), String(r[4])]);
  }

  // Создаём/обновляем лист для каждого кабинета
  for (const acc of currentAccounts) {
    const sheetName = PRODUCTS_PREFIX + acc.name;
    s = ensureSheet_(ss, sheetName);

    // Сохраняем уже введённые данные
    const existingRows = s.getLastRow() > 1
      ? s.getRange(2, 1, s.getLastRow() - 1, 4).getValues()
      : [];

    s.clear();
    s.getRange(1, 1, 1, 4).setValues([['SKU (offer_id)', 'Название', 'Менеджер', 'Статус']]);

    // Приоритет данных: 1) уже были в новом листе, 2) миграция из старого листа
    let dataToWrite = existingRows.filter(r => r[0]);
    if (!dataToWrite.length && migrationByAccount[acc.id] && migrationByAccount[acc.id].length) {
      dataToWrite = migrationByAccount[acc.id];
    }

    if (dataToWrite.length) {
      s.getRange(2, 1, dataToWrite.length, 4).setValues(dataToWrite);
    }

    formatHeader_(s, 4);
    s.setColumnWidth(1, 140);
    s.setColumnWidth(2, 260);
    s.setColumnWidth(3, 120);
    s.setColumnWidth(4, 90);

    // Валидация статуса
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['активен', 'пауза'], true).build();
    s.getRange(2, 4, 1000, 1).setDataValidation(statusRule);

    // Замораживаем шапку
    s.setFrozenRows(1);
  }

  // Скрываем (не удаляем!) старый лист «Артикулы» если он ещё есть
  if (oldProductSheet) oldProductSheet.hideSheet();

  // ===== _Данные (служебный, скрытый) =====
  s = getDataSheet_();
  s.hideSheet();

  // ===== ⚠️ Предупреждения =====
  s = ensureSheet_(ss, SHEET.ALERTS);
  s.clear();
  s.getRange(1, 1, 1, 7).setValues([[
    'Кабинет', 'Артикул', 'Название', 'Менеджер', 'Тип', 'Предупреждение', 'Дата'
  ]]);
  formatHeader_(s, 7);

  // ===== Лог =====
  s = ensureSheet_(ss, SHEET.LOG);
  s.clear();
  s.getRange(1, 1, 1, 4).setValues([['Время', 'Уровень', 'Функция', 'Сообщение']]);
  formatHeader_(s, 4);
  s.setColumnWidth(4, 600);

  ui.alert('Готово',
    'Таблица инициализирована. Дальше:\n\n' +
    '1. Лист «Аккаунты» — впиши Client-Id, Api-Key и (опц.) Performance client_id|secret. Название кабинета = имя его листа.\n' +
    '2. Лист «Артикулы» — добавь SKU (offer_id), укажи ID аккаунта и менеджера, статус «активен».\n' +
    '3. Меню → 🔧 Тест подключения к Ozon API.\n' +
    '4. Меню → ▶️ Обновить данные сейчас (создаст листы кабинетов).\n' +
    '5. Меню → ⏰ Установить ежедневный триггер.',
    ui.ButtonSet.OK);
}


// ====== ОБЩИЕ ХЕЛПЕРЫ ======
function ensureSheet_(ss, name) {
  let s = ss.getSheetByName(name);
  if (!s) s = ss.insertSheet(name);
  return s;
}

function formatHeader_(sheet, cols) {
  sheet.getRange(1, 1, 1, cols)
    .setFontFamily('Onest')
    .setFontWeight('bold')
    .setBackground('#4a86e8')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setWrap(true);
  sheet.setFrozenRows(1);
}


// ====== ТЕСТ DRIVE-ОТЧЁТОВ ======
/**
 * Диагностика: читает Drive-отчёт за вчера и логирует:
 *  - список подпапок в главной папке
 *  - найденные файлы по каждому кабинету
 *  - первые 3 строки каждого файла (чтобы видеть заголовки)
 *  - определённые колонки Артикул и CTR
 *  - первые 5 строк данных (sku → ctr)
 *
 * Запусти ОДИН РАЗ через меню → 🗂️ Тест Drive-отчётов.
 * Смотри лист «Лог» — там полный вывод.
 */
function testDriveReport() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  try {
    const accounts = getAccounts_();
    if (!accounts.length) { ui.alert('Нет активных аккаунтов'); return; }

    // Спрашиваем дату (по умолчанию — вчера)
    const tz = ss.getSpreadsheetTimeZone();
    const defaultDate = Utilities.formatDate(addDays_(new Date(), -1), tz, 'dd.MM.yyyy');
    const res = ui.prompt(
      '🗂️ Тест Drive-отчётов',
      `За какую дату проверить? (формат дд.мм.гггг)\nПо умолчанию: ${defaultDate}`,
      ui.ButtonSet.OK_CANCEL
    );
    if (res.getSelectedButton() !== ui.Button.OK) return;
    const inputRaw = res.getResponseText().trim();

    let metricDate;
    if (!inputRaw) {
      metricDate = addDays_(new Date(), -1);
    } else {
      // Парсим дд.мм.гггг или дд.мм (текущий год)
      const parts = inputRaw.split('.');
      const d = parseInt(parts[0], 10);
      const m = parseInt(parts[1], 10) - 1; // месяц 0-based
      const y = parts[2] ? parseInt(parts[2], 10) : new Date().getFullYear();
      metricDate = new Date(y, m, d);
      if (isNaN(metricDate.getTime())) {
        ui.alert('Неверный формат даты. Введи дд.мм.гггг, например 04.06.2026');
        return;
      }
    }

    const dayMonth = Utilities.formatDate(metricDate, tz, 'dd.MM');
    const dateStr  = Utilities.formatDate(metricDate, tz, 'yyyy-MM-dd');

    log_('INFO', 'testDriveReport',
      `=== Тест Drive-отчётов за ${dayMonth} (${dateStr}) ===`);

    // Главная папка
    const mainFolder = DriveApp.getFolderById(DRIVE_REPORTS_FOLDER_ID);
    log_('INFO', 'testDriveReport',
      `Папка Drive: «${mainFolder.getName()}» (ID=${DRIVE_REPORTS_FOLDER_ID})`);

    // Список всех подпапок
    const subfolderNames = [];
    const fi = mainFolder.getFolders();
    while (fi.hasNext()) subfolderNames.push(fi.next().getName());
    log_('INFO', 'testDriveReport',
      `Подпапок всего: ${subfolderNames.length}. Названия: ${subfolderNames.slice(0, 20).join(', ')}`);

    // Поиск папки для вчера
    const dateFolder = findFolderByDate_(mainFolder, dayMonth, dateStr);
    if (!dateFolder) {
      flushLog_();
      ui.alert(
        `Папка для ${dayMonth} не найдена.\n\n` +
        `Есть папки: ${subfolderNames.slice(0, 15).join(', ')}\n\n` +
        `Подробности в листе «Лог».`
      );
      return;
    }
    log_('INFO', 'testDriveReport', `Найдена папка для даты: «${dateFolder.getName()}»`);

    // Файлы в папке
    const fileIter = dateFolder.getFiles();
    let fileCount = 0;
    while (fileIter.hasNext()) {
      const f = fileIter.next();
      log_('INFO', 'testDriveReport',
        `Файл[${fileCount}]: ${f.getName()}  (mime=${f.getMimeType()})`);
      fileCount++;
    }
    if (!fileCount) {
      flushLog_();
      ui.alert(`Папка «${dateFolder.getName()}» пуста. Подробности в «Лог».`);
      return;
    }

    // Читаем каждый аккаунт
    for (const acc of accounts) {
      log_('INFO', 'testDriveReport', `─── Кабинет: «${acc.name}» ───`);
      const xlsxFile = findCabinetFile_(dateFolder, acc.name);
      if (!xlsxFile) {
        log_('WARN', 'testDriveReport', `Файл кабинета не найден!`);
        continue;
      }
      log_('INFO', 'testDriveReport', `Файл: ${xlsxFile.getName()}`);

      // Читаем массив
      const data = readFileAsArray_(xlsxFile);
      log_('INFO', 'testDriveReport', `Строк в файле: ${data.length}`);

      // Выводим первые 4 строки чтобы видеть заголовки
      for (let r = 0; r < Math.min(4, data.length); r++) {
        const preview = data[r].slice(0, 12).map((v, i) =>
          `${colLetter_(i)}:«${String(v).slice(0, 20)}»`
        ).join('  ');
        log_('INFO', 'testDriveReport', `Строка ${r + 1}: ${preview}`);
      }

      // Определяем колонки
      const { skuCol, viewsCol, clicksCol, startRow } = detectColumns_(data);
      log_('INFO', 'testDriveReport',
        `Определено: Артикул → ${colLetter_(skuCol)}, ` +
        `Просмотры → ${colLetter_(viewsCol)}, ` +
        `Клики → ${colLetter_(clicksCol)}, данные с строки ${startRow + 1}`);

      // Агрегируем и выводим первые 5 SKU
      const agg = {};
      for (let r = startRow; r < data.length; r++) {
        const row = data[r];
        const sku = String(row[skuCol] || '').trim();
        if (!sku) continue;
        if (!agg[sku]) agg[sku] = { views: 0, clicks: 0 };
        agg[sku].views  += parseNum_(row[viewsCol]);
        agg[sku].clicks += parseNum_(row[clicksCol]);
      }
      let shown = 0;
      for (const sku of Object.keys(agg)) {
        if (shown >= 5) break;
        const { views, clicks } = agg[sku];
        const ctr = views > 0 ? Math.round(clicks / views * 10000) / 100 : 0;
        log_('INFO', 'testDriveReport',
          `  SKU=${sku}  Просмотры=${views}  Клики=${clicks}  Рекл.CTR=${ctr}%`);
        shown++;
      }
    }

    flushLog_();
    ui.alert(
      '✅ Drive тест завершён.\n\n' +
      'Посмотри лист «Лог» — там заголовки файла, найденные колонки и первые SKU с CTR.\n\n' +
      'Если колонки определились неправильно — сообщи, скорректируем detectColumns_().'
    );
  } catch (e) {
    flushLog_();
    ui.alert('❌ Ошибка: ' + e.message + '\n\nПодробности в листе «Лог».');
  }
}


// ====== ТЕСТ РЕКЛАМНОГО API ======
/**
 * Показывает диагностику Performance API:
 * токен, кол-во кампаний, сырой ответ statistics/products/json.
 * После запуска смотри лист «Лог» — там полный вывод.
 */
function testPerfApi() {
  const ui = SpreadsheetApp.getUi();
  try {
    const accounts = getAccounts_();
    if (!accounts.length) { ui.alert('Нет активных аккаунтов'); return; }
    const products = getActiveProducts_();
    const account  = accounts[0];
    const sku      = products.length ? products[0].sku : 'TEST';
    const metricDate = addDays_(new Date(), -1);

    log_('INFO', 'testPerfApi', `Аккаунт: ${account.id}, SKU: ${sku}`);

    // 1. Токен
    const token = getPerfToken_(account);
    log_('INFO', 'testPerfApi', `Токен получен: ${token.slice(0, 12)}...`);

    // 2. Кампании
    const campResp = UrlFetchApp.fetch(PERF_BASE + '/api/client/campaign', {
      method: 'get', headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true
    });
    const campBody = campResp.getContentText();
    log_('INFO', 'testPerfApi', `Кампании HTTP ${campResp.getResponseCode()}: ${campBody.slice(0, 400)}`);
    const campaigns = JSON.parse(campBody).list || [];

    if (!campaigns.length) {
      flushLog_();
      ui.alert('Performance API', 'Токен OK, но кампаний не найдено (list пустой).\nПодробности в листе «Лог».', ui.ButtonSet.OK);
      return;
    }

    // 3. Статистика за вчера
    const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
    const df = Utilities.formatDate(metricDate, tz, 'yyyy-MM-dd');
    // Берём только первые 10 кампаний для теста (лимит API = 10 за запрос)
    const allIds = campaigns.map(c => String(c.id));
    const chunkIds = allIds.slice(0, 10);
    log_('INFO', 'testPerfApi', `Кампаний всего: ${allIds.length}, в тесте первые ${chunkIds.length}`);
    const statResp = UrlFetchApp.fetch(PERF_BASE + '/api/client/statistics', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + token },
      payload: JSON.stringify({ campaigns: chunkIds, groupBy: 'GROUP_BY_SKU', dateFrom: df, dateTo: df }),
      muteHttpExceptions: true
    });
    const statBody = statResp.getContentText();
    log_('INFO', 'testPerfApi', `Statistics HTTP ${statResp.getResponseCode()}`);
    log_('INFO', 'testPerfApi', `Statistics body (first 800): ${statBody.slice(0, 800)}`);

    flushLog_();
    ui.alert('Performance API — результат',
      `Токен: OK\nКампаний: ${campaigns.length}\n` +
      `Statistics HTTP: ${statResp.getResponseCode()}\n\n` +
      `Подробный ответ API — в листе «Лог» (последние строки).`,
      ui.ButtonSet.OK);
  } catch (e) {
    flushLog_();
    ui.alert('Ошибка', e.message + '\n\nПодробности в листе «Лог»', ui.ButtonSet.OK);
  }
}


// ====== ТЕСТ ПОДКЛЮЧЕНИЯ ======
function testApiConnection() {
  const ui = SpreadsheetApp.getUi();
  try {
    const accounts = getAccounts_();
    if (!accounts.length) { ui.alert('Нет активных аккаунтов в листе «Аккаунты»'); return; }
    const products = getActiveProducts_();
    if (!products.length) { ui.alert('Нет активных артикулов в листе «Артикулы»'); return; }

    const product = products[0];
    const account = accounts.find(a => a.id === product.accountId) || accounts[0];
    const metricDate = addDays_(new Date(), -(Number(getSetting_(SETTING_KEYS.LAG_DAYS)) || 1));

    const details = fetchProductDetails_(account, product.sku);
    const cover = coverRecord_(product, account, details, new Date());
    const rec   = metricsRecord_(product, account, details, metricDate);
    flushLog_();

    ui.alert('OK',
      `Кабинет: ${account.name}\nАртикул: ${rec.sku} (${rec.name})\n` +
      `Дата метрик: ${rec.date}\n\n` +
      `Обложка (сегодня): ${cover.coverUrl ? 'есть' : 'нет'}\n\n` +
      `Все показы (hits_view):             ${rec.views}\n` +
      `Клики в карточку (hits_view_pdp):   ${rec.pdpViews}\n` +
      `Добавлений в корзину:               ${rec.toCart}\n` +
      `Заказы:                             ${rec.orders}\n\n` +
      `CTR (клики/показы):                 ${rec.ctr}%\n` +
      `Конв. в корзину (корзины/клики):    ${rec.cartConvPct}%\n` +
      `CR корзина→заказ (заказы/корзины):  ${rec.cartToOrderPct}%\n` +
      `Выкупы: ${rec.buyouts} (${rec.buyoutPct}%)`,
      ui.ButtonSet.OK);
  } catch (e) {
    flushLog_();
    ui.alert('Ошибка', e.message + '\n\n' + e.stack, ui.ButtonSet.OK);
  }
}
