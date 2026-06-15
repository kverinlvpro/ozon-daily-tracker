/**
 * OZON DAILY COVER TRACKER
 * ========================
 * Ежедневный трекер обложек товаров Ozon для менеджеров маркетплейса.
 *
 * Что делает:
 *  - Раз в день (или по кнопке) для каждого артикула снимает текущую обложку
 *    и метрики: CTR, показы, конверсию в корзину, CR корзина→заказ, заказы, выкупы.
 *  - Складывает «сырые» данные в служебный лист и рисует по кабинетам наглядные
 *    листы: каждый артикул — блок строк, каждый день — колонка, в верхней строке
 *    блока видно как менялась обложка день ото дня.
 *  - Предупреждает: (1) если после смены обложки конверсия упала/выросла,
 *    (2) если обложка не менялась более N дней.
 *
 * Структура листов:
 *  - Настройки        — глобальные параметры
 *  - Аккаунты         — кабинеты Ozon (Client-Id, Api-Key, Performance creds)
 *  - Артикулы         — SKU под трекингом (привязаны к кабинету)
 *  - <Кабинет 1..N>   — наглядные листы по кабинетам (блочная раскладка)
 *  - ⚠️ Предупреждения — сводка алертов
 *  - _Данные          — служебный append-only лог метрик по дням (можно скрыть)
 *  - Лог              — отладочный журнал
 */

// ====== СЛУЖЕБНЫЕ ИМЕНА ЛИСТОВ ======
// (листы кабинетов называются по полю «Название» аккаунта)
const SHEET = {
  SETTINGS: 'Настройки',
  ACCOUNTS: 'Аккаунты',
  PRODUCTS: 'Артикулы',
  ALERTS:   '⚠️ Предупреждения',
  DATA:     '_Данные',
  LOG:      'Лог'
};

// Префикс листов артикулов (по одному на кабинет: «Артикулы — ИП Иванова» и т.д.)
const PRODUCTS_PREFIX = 'Артикулы — ';

// Имена листов, которые НЕ являются листами кабинетов
// Листы, начинающиеся с PRODUCTS_PREFIX, тоже исключаются — см. isCabinetSheet_()
const RESERVED_SHEETS = [
  SHEET.SETTINGS, SHEET.ACCOUNTS, SHEET.PRODUCTS,
  SHEET.ALERTS, SHEET.DATA, SHEET.LOG
];

// ====== КЛЮЧИ НАСТРОЕК ======
const SETTING_KEYS = {
  LAG_DAYS:        'lag_days',          // за какой день брать метрики (1 = вчера; аналитика Ozon с задержкой)
  UPDATE_HOUR:     'update_hour',       // час ежедневного триггера (0-23)
  DISPLAY_DAYS:    'display_days',      // сколько последних дней показывать колонками
  STALE_DAYS:      'stale_cover_days',  // обложка не менялась дольше — предупреждение
  ALERT_PCT:       'alert_threshold_pct', // порог изменения конверсии для алерта, %
  BEFORE_DAYS:     'before_window_days',  // окно ДО смены обложки для сравнения
  AFTER_DAYS:      'after_window_days'    // окно ПОСЛЕ смены обложки для сравнения
};

// ====== ФОРМУЛЫ КОНВЕРСИЙ ======
// Все считаются из стабильных абсолютных метрик Seller API:
//   CTR               = hits_view_pdp  / hits_view_search  (клики в карточку / показы в поиске)
//   Конверсия в корзину = hits_tocart_pdp / hits_view_pdp  (корзины / клики)
//   CR корзина→заказ  = ordered_units  / hits_tocart_pdp  (заказы / корзины)
// Никаких процентных метрик от Ozon — только деление абсолютных чисел.

// ====== СТРОКИ БЛОКА АРТИКУЛА (в порядке сверху вниз) ======
// Первая строка — обложка (картинка по дням). Остальные — метрики/ручной ввод.
// Порядок важен: строки 0..(GROUP_START_OFFSET-1) остаются видимыми при сворачивании,
// остальное группируется. Сейчас видны «Обложка», «CTR» и «Рекламный CTR».
const ROW_LABELS = [
  'Обложка',            // 0 — IMAGE по дням (видна при сворачивании)
  'CTR',                // 1 — общий CTR, % (видна при сворачивании)
  'Рекламный CTR',      // 2 — CTR по рекламным кампаниям, % (видна при сворачивании)
  'Гипотеза',           // 3 — ручной ввод
  'Показы',             // 4 — метрика
  'Кр в корзину',       // 5 — метрика (конверсия в корзину), %
  'CR корзина→заказ',   // 6 — метрика, %
  'Заказы',             // 7 — метрика
  'Выкупы',             // 8 — метрика
  'Вид продвижения',    // 9 — ручной ввод
  'Комментарий'         // 10 — ручной ввод
];
const BLOCK_H = ROW_LABELS.length;   // высота блока одного артикула в строках
const IMAGE_ROW_OFFSET = 0;          // смещение строки обложки внутри блока
// С какого смещения внутри блока начинается группируемая (сворачиваемая) часть.
// Всё до него (Обложка + CTR + Рекламный CTR) остаётся видимым в свёрнутом виде.
const GROUP_START_OFFSET = 3;

// Строки, которые менеджер заполняет вручную (рендер их сохраняет при перерисовке)
const MANUAL_LABELS = ['Гипотеза', 'Вид продвижения', 'Комментарий'];

// Какую метрику класть в какую строку: label -> ключ поля записи
const METRIC_ROW_FIELD = {
  'CTR':              'ctr',
  'Рекламный CTR':    'adCtr',
  'Показы':           'views',
  'Кр в корзину':     'cartConvPct',
  'CR корзина→заказ': 'cartToOrderPct',
  'Заказы':           'orders',
  'Выкупы':           'buyouts'
};
// Строки с процентным форматом
const PERCENT_LABELS = ['CTR', 'Рекламный CTR', 'Кр в корзину', 'CR корзина→заказ'];

// Геометрия листа кабинета
const FIRST_DATA_COL = 7;     // колонка G — первый день (A=Артикул,B=Имя,C=Наименование,D=Прошлый период,E=Калькулятор,F=Метрика)
const PREV_COL       = 4;     // колонка D — прошлый период (авто)
const CALC_COL       = 5;     // колонка E — калькулятор периода (пользователь)
const HEADER_ROWS = 1;        // строка 1 — заголовок с датами
const IMAGE_ROW_HEIGHT = 120; // высота строки обложки, px
const DAY_COL_WIDTH = 110;    // ширина колонки одного дня, px


/**
 * ТОЧКА ВХОДА 1: ежедневное обновление (триггер / кнопка).
 *  1) Снимает обложку и метрики по каждому артикулу за нужный день
 *  2) Пишет записи в служебный лист _Данные (upsert по дате+SKU)
 *  3) Перерисовывает листы кабинетов
 *  4) Пересчитывает предупреждения
 */
function dailyUpdate() {
  const startedAt = new Date();
  log_('INFO', 'dailyUpdate', 'Старт ежедневного обновления');

  try {
    const accounts = getAccounts_();
    const products = getActiveProducts_();
    log_('INFO', 'dailyUpdate', `Аккаунтов: ${accounts.length}, артикулов: ${products.length}`);

    const lagDays    = Number(getSetting_(SETTING_KEYS.LAG_DAYS)) || 1;
    const todayKey   = dateKey_(new Date());
    const metricDate = addDays_(new Date(), -lagDays);
    const dayFrom    = startOfDay_(metricDate);
    const dayTo      = endOfDay_(metricDate);

    const records = [];

    for (const account of accounts) {
      const accProducts = products.filter(p => p.accountId === account.id);
      if (!accProducts.length) continue;
      const offerIds = accProducts.map(p => p.sku);

      log_('INFO', 'dailyUpdate', `Кабинет «${account.name}»: ${accProducts.length} артикулов`);

      // ── 4 пакетных вызова на кабинет (не зависит от числа SKU) ──────────
      const detailsMap  = fetchAllProductDetailsBatch_(account, offerIds);
      const analyticsMap = fetchAllAnalyticsBatch_(account, dayFrom, dayTo);
      const buyoutsMap  = fetchAllBuyoutsBatch_(account, offerIds, dayFrom, dayTo);
      const adMap       = fetchAllAdStatsForAccount_(account, metricDate);

      log_('INFO', 'dailyUpdate',
        `${account.name}: деталей=${Object.keys(detailsMap).length}, ` +
        `аналитики=${Object.keys(analyticsMap).length}, рекл.CTR=${Object.keys(adMap).length}`);

      for (const product of accProducts) {
        try {
          const details = detailsMap[product.sku];
          if (!details) {
            log_('WARN', 'dailyUpdate', `SKU ${product.sku} не найден в кабинете`);
            continue;
          }
          const an      = analyticsMap[String(details.sku)] || null;  // по числовому SKU
          const buy     = buyoutsMap[product.sku] || null;
          const adStats = adMap[product.sku] || null;

          records.push(coverRecord_(product, account, details, todayKey));
          records.push(metricsRecord_(product, account, details, metricDate, adStats, an, buy));
        } catch (e) {
          log_('ERROR', 'dailyUpdate', `SKU ${product.sku}: ${e.message}`);
        }
      }
    }

    if (records.length) mergeUpsertRecords_(records);
    renderAllCabinets_(accounts, products);
    rebuildAlerts_(accounts, products);

    log_('INFO', 'dailyUpdate',
      `Готово. Записей: ${records.length}, время: ${(new Date() - startedAt) / 1000}s`);
  } catch (e) {
    log_('FATAL', 'dailyUpdate', e.message + '\n' + e.stack);
  } finally {
    flushLog_();
  }
}


/**
 * ТОЧКА ВХОДА 2: установить ежедневный триггер. Запустить один раз вручную.
 */
function installDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'dailyUpdate') ScriptApp.deleteTrigger(t);
  });
  const hour = Number(getSetting_(SETTING_KEYS.UPDATE_HOUR));
  const atHour = (hour >= 0 && hour <= 23) ? hour : 7;
  ScriptApp.newTrigger('dailyUpdate').timeBased().everyDays(1).atHour(atHour).create();
  // atHour работает в таймзоне скрипта (Екатеринбург UTC+5).
  // 14:00 Екб = 12:00 МСК = 09:00 UTC
  const mskHour = ((atHour - 2 + 24) % 24); // UTC+5 → UTC+3
  SpreadsheetApp.getUi().alert(`Ежедневный триггер установлен на ~${atHour}:00 Екб (~${mskHour}:00 МСК).`);
}


/**
 * ТОЧКА ВХОДА 2б: установить отдельный триггер для рекламного CTR (~13:00 МСК).
 * Запускает fillAdCtrYesterdayTrigger() через 1 ч после основного обновления,
 * когда Drive-отчёты гарантированно появляются.
 */
function installAdCtrTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'fillAdCtrYesterdayTrigger') ScriptApp.deleteTrigger(t);
  });
  // 15:00 Екб (UTC+5) = 13:00 МСК (UTC+3)
  ScriptApp.newTrigger('fillAdCtrYesterdayTrigger').timeBased().everyDays(1).atHour(15).create();
  SpreadsheetApp.getUi().alert(
    'Триггер рекл. CTR установлен на ~15:00 Екб (~13:00 МСК).\n\n' +
    'Каждый день в 13:00 МСК таблица автоматически подтянет рекламный CTR из Drive за вчера.'
  );
}

/**
 * ТОЧКА ВХОДА 3: меню в таблице.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📸 Трекер обложек Ozon')
    .addItem('▶️ Обновить данные сейчас', 'dailyUpdate')
    .addItem('⏰ Установить триггер обновления (~12:00 МСК)', 'installDailyTrigger')
    .addItem('⏰ Установить триггер рекл. CTR (~13:00 МСК)', 'installAdCtrTrigger')
    .addSeparator()
    .addItem('🔄 Перерисовать листы кабинетов', 'renderAllCabinetsMenu')
    .addItem('⚠️ Пересчитать предупреждения', 'rebuildAlertsMenu')
    .addSeparator()
    .addItem('🔁 Пересчитать метрики за 7 дней', 'recalculate7Days')
    .addItem('🔁 Пересчитать метрики за N дней…', 'recalculateNDaysMenu')
    .addItem('📊 Заполнить Рекл. CTR из Drive за N дней…', 'backfillAdCtrMenu')
    .addSeparator()
    .addItem('📅 Калькулятор периода (календарь)', 'showCalendarSidebar')
    .addItem('🧮 Пересчитать калькулятор (текущий период)', 'runCalculatorMenu')
    .addSeparator()
    .addItem('🔽 Развернуть все блоки', 'expandAllBlocks')
    .addItem('🔼 Свернуть все блоки', 'collapseAllBlocks')
    .addSeparator()
    .addItem('🆕 Инициализировать таблицу', 'initSpreadsheet')
    .addItem('🔄 Подтянуть названия артикулов из Ozon', 'lookupProductNamesMenu')
    .addItem('🔧 Тест подключения к Ozon API', 'testApiConnection')
    .addItem('🗂️ Тест Drive-отчётов (рекл. CTR)', 'testDriveReport')
    .addItem('📣 Тест рекламного API (Performance)', 'testPerfApi')
    .addToUi();
}

// Обёртки для пунктов меню (без аргументов)
function renderAllCabinetsMenu() {
  renderAllCabinets_(getAccounts_(), getActiveProducts_());
  flushLog_();
  SpreadsheetApp.getUi().alert('Листы кабинетов перерисованы.');
}
function rebuildAlertsMenu() {
  rebuildAlerts_(getAccounts_(), getActiveProducts_());
  flushLog_();
  SpreadsheetApp.getUi().alert('Предупреждения пересчитаны. См. лист «⚠️ Предупреждения».');
}
/**
 * Пересчитывает метрики за последние nDays дней для всех активных артикулов.
 * Обложки НЕ трогает — только метрики (CTR, показы, корзина, заказы, выкупы).
 * Используется для исправления исторических данных после смены формул или фикса багов.
 */
function recalculateDays_(nDays) {
  const accounts = getAccounts_();
  const products  = getActiveProducts_();
  if (!accounts.length || !products.length) {
    SpreadsheetApp.getUi().alert('Нет активных аккаунтов или артикулов.');
    return;
  }

  let ok = 0, errors = 0;
  for (let d = 1; d <= nDays; d++) {
    const metricDate = addDays_(new Date(), -d);
    const dayFrom    = startOfDay_(metricDate);
    const dayTo      = endOfDay_(metricDate);

    const records = [];
    for (const account of accounts) {
      const accProducts = products.filter(p => p.accountId === account.id);
      if (!accProducts.length) continue;
      const offerIds = accProducts.map(p => p.sku);

      // Пакетные вызовы за день
      const detailsMap   = fetchAllProductDetailsBatch_(account, offerIds);
      const analyticsMap = fetchAllAnalyticsBatch_(account, dayFrom, dayTo);
      const buyoutsMap   = fetchAllBuyoutsBatch_(account, offerIds, dayFrom, dayTo);
      const adMap        = fetchAllAdStatsForAccount_(account, metricDate);

      for (const product of accProducts) {
        try {
          const details = detailsMap[product.sku];
          if (!details) continue;
          const an      = analyticsMap[String(details.sku)] || null;
          const buy     = buyoutsMap[product.sku] || null;
          const adStats = adMap[product.sku] || null;
          records.push(metricsRecord_(product, account, details, metricDate, adStats, an, buy));
          ok++;
        } catch (e) {
          log_('ERROR', 'recalculateDays_', `SKU ${product.sku} день -${d}: ${e.message}`);
          errors++;
        }
      }
    }
    if (records.length) mergeUpsertRecords_(records);
  }

  renderAllCabinets_(accounts, products);
  flushLog_();
  SpreadsheetApp.getActive().toast(
    `Пересчитано: ${ok} записей за ${nDays} дн., ошибок: ${errors}`, '🔁 Готово', 5);
}

function recalculate7Days()     { recalculateDays_(7); }
function recalculateNDaysMenu() {
  const ui  = SpreadsheetApp.getUi();
  const res = ui.prompt('Пересчитать метрики', 'За сколько последних дней пересчитать? (1–30)', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const n = parseInt(res.getResponseText(), 10);
  if (!n || n < 1 || n > 30) { ui.alert('Введи число от 1 до 30.'); return; }
  recalculateDays_(n);
}

/** Возвращает true, если лист является листом кабинета (не служебным и не артикулами). */
function isCabinetSheet_(name) {
  if (RESERVED_SHEETS.indexOf(name) >= 0) return false;
  if (name.startsWith(PRODUCTS_PREFIX)) return false;
  return true;
}

function expandAllBlocks() {
  const n = setAllGroups_(false);
  SpreadsheetApp.getActive().toast(`Развёрнуто блоков: ${n}`);
}
function collapseAllBlocks() {
  const n = setAllGroups_(true);
  SpreadsheetApp.getActive().toast(`Свёрнуто блоков: ${n}`);
}
