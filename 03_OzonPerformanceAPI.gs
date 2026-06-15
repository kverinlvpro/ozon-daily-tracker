/**
 * OZON PERFORMANCE API
 * ====================
 * Документация: https://docs.ozon.ru/api/performance/
 *
 * Используем для CTR рекламы и показов в рекламных кампаниях.
 *
 * Авторизация: OAuth 2.0 (client_credentials).
 * Токен живёт 30 минут — кэшируем в CacheService.
 *
 * Скопировано из проекта ozon-ab без изменений.
 */

const PERF_BASE = 'https://api-performance.ozon.ru';

function getPerfToken_(account) {
  if (!account.performanceClientId || !account.performanceSecret) {
    throw new Error('Performance API креды не заданы в листе Аккаунты');
  }

  const cache = CacheService.getScriptCache();
  const cacheKey = 'perf_token_' + account.id;
  const cached = cache.get(cacheKey);
  if (cached) return cached;

  const resp = UrlFetchApp.fetch(PERF_BASE + '/api/client/token', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      client_id:     account.performanceClientId,
      client_secret: account.performanceSecret,
      grant_type:    'client_credentials'
    }),
    muteHttpExceptions: true
  });

  if (resp.getResponseCode() !== 200) {
    throw new Error('Performance auth: ' + resp.getContentText());
  }
  const json = JSON.parse(resp.getContentText());
  const token = json.access_token;
  // Токен на 30 мин, кэшируем на 25 для надёжности
  cache.put(cacheKey, token, 1500);
  return token;
}

/**
 * Получает рекламную статистику по SKU за период.
 * Возвращает {adViews, adClicks, ctr}.
 *
 * NB: Performance API отдаёт статистику по кампаниям, а не по SKU напрямую.
 * Здесь — упрощённая логика: получаем все активные кампании и агрегируем
 * показы/клики по нужному SKU. Если рекламы у SKU нет, вернётся нули.
 */
/**
 * Получает рекламную статистику по ВСЕМ SKU кабинета за один день.
 * Возвращает объект {offer_id → {adViews, adClicks, ctr}}.
 *
 * Вызывается ОДИН РАЗ на аккаунт (не на каждый SKU), чтобы не множить
 * async-запросы при большом числе кампаний (может быть 500+).
 *
 * Фильтрация: только кампании в статусе RUNNING с бюджетом > 0.
 * Этим отсекаем нулевые и незапущенные кампании без данных.
 */
function fetchAllAdStatsForAccount_(account, metricDate) {
  // ── Источник 1: Excel-отчёты из Google Drive ──────────────────────────────
  // Быстро, надёжно, без асинхронных ZIP-архивов.
  // Если папка/файл для этой даты не загружены — тихо переходим к API.
  try {
    const driveResult = readAdCtrFromDriveReport_(account, metricDate);
    if (Object.keys(driveResult).length > 0) {
      log_('INFO', 'fetchAllAdStats_',
        `Drive отчёт: кабинет «${account.name}», SKU с CTR: ${Object.keys(driveResult).length}`);
      return driveResult;
    }
  } catch (e) {
    log_('WARN', 'fetchAllAdStats_',
      `Drive отчёт недоступен (${e.message}), переключаемся на Performance API`);
  }

  // ── Источник 2: Ozon Performance API (fallback) ───────────────────────────
  // Используется если Drive-файл за нужную дату не найден.
  // Известное ограничение: API возвращает ZIP-архив, парсинг которого
  // пока не реализован → результат всегда пустой (adCtr = 0).
  const result = {};  // offer_id → {adViews, adClicks, ctr}
  try {
    if (!account.performanceClientId || !account.performanceSecret) return result;
    const token = getPerfToken_(account);
    const tz    = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
    const dateStr = Utilities.formatDate(metricDate, tz, 'yyyy-MM-dd');

    // 1. Список кампаний
    const campResp = UrlFetchApp.fetch(PERF_BASE + '/api/client/campaign', {
      method: 'get',
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });
    if (campResp.getResponseCode() !== 200) {
      log_('WARN', 'fetchAllAdStats_', `Кампании HTTP ${campResp.getResponseCode()}`);
      return result;
    }
    const allCampaigns = JSON.parse(campResp.getContentText()).list || [];

    // Только RUNNING-кампании с бюджетом > 0 и стартовавшие до нужной даты
    const active = allCampaigns.filter(c =>
      c.state === 'CAMPAIGN_STATE_RUNNING' &&
      (parseFloat(c.budget || '0') > 0 || parseFloat(c.dailyBudget || '0') > 0) &&
      c.fromDate <= dateStr
    );
    log_('DEBUG', 'fetchAllAdStats_',
      `Аккаунт ${account.id}: кампаний всего=${allCampaigns.length}, активных с бюджетом=${active.length}`);
    if (!active.length) return result;

    // 2. Запросы батчами по 10, собираем все строки
    const CHUNK = 10;
    const allIds = active.map(c => String(c.id));
    const allRows = [];

    for (let i = 0; i < allIds.length; i += CHUNK) {
      const chunk = allIds.slice(i, i + CHUNK);
      const statResp = UrlFetchApp.fetch(PERF_BASE + '/api/client/statistics', {
        method: 'post',
        contentType: 'application/json',
        headers: { 'Authorization': 'Bearer ' + token },
        payload: JSON.stringify({ campaigns: chunk, groupBy: 'GROUP_BY_SKU', dateFrom: dateStr, dateTo: dateStr }),
        muteHttpExceptions: true
      });

      if (statResp.getResponseCode() !== 200) {
        log_('WARN', 'fetchAllAdStats_', `Батч ${i/CHUNK+1}: HTTP ${statResp.getResponseCode()}`);
        continue;
      }

      let parsed = JSON.parse(statResp.getContentText());
      const batchN = Math.floor(i / CHUNK) + 1;

      if (parsed.UUID) {
        log_('DEBUG', 'fetchAllAdStats_', `Батч ${batchN}: async UUID=${parsed.UUID}, поллим...`);
        parsed = pollReport_(token, parsed.UUID);
        if (!parsed) continue;
      }

      const rows = extractAdRows_(parsed);
      log_('DEBUG', 'fetchAllAdStats_', `Батч ${batchN}: строк ${rows.length}` +
        (rows.length ? ', поля: ' + Object.keys(rows[0]).join(', ') : ''));
      rows.forEach(r => allRows.push(r));
    }

    // 3. Строим map offer_id → aggregated stats
    for (const r of allRows) {
      const sku = String(r.sku || r.sku_id || r.article || r.offer_id || r.ozonId || r.itemId || '');
      if (!sku) continue;
      if (!result[sku]) result[sku] = { adViews: 0, adClicks: 0, ctr: 0 };
      result[sku].adViews  += Number(r.views || r.impressions || r.shows || r.view || 0);
      result[sku].adClicks += Number(r.clicks || r.click || 0);
    }
    for (const sku of Object.keys(result)) {
      const s = result[sku];
      s.ctr = s.adViews > 0 ? round2_((s.adClicks / s.adViews) * 100) : 0;
    }
    log_('DEBUG', 'fetchAllAdStats_', `Итого SKU с рекламой: ${Object.keys(result).length}`);

  } catch (e) {
    log_('WARN', 'fetchAllAdStats_', `Performance API, аккаунт ${account.id}: ${e.message}`);
  }
  return result;
}


// Старая per-SKU функция — оставлена для совместимости с тестовыми вызовами,
// но из ежедневного обновления больше не вызывается.
function fetchAdStats_(account, sku, dateFrom, dateTo) {
  try {
    const token = getPerfToken_(account);
    const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
    const dateFromStr = Utilities.formatDate(dateFrom, tz, 'yyyy-MM-dd');
    const dateToStr   = Utilities.formatDate(dateTo,   tz, 'yyyy-MM-dd');

    // 1. Список кампаний
    const campResp = UrlFetchApp.fetch(PERF_BASE + '/api/client/campaign', {
      method: 'get',
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });
    if (campResp.getResponseCode() !== 200) {
      log_('WARN', 'fetchAdStats_', `Кампании HTTP ${campResp.getResponseCode()}: ${campResp.getContentText().slice(0, 200)}`);
      return { adViews: 0, adClicks: 0, ctr: 0 };
    }
    const campaigns = JSON.parse(campResp.getContentText()).list || [];
    log_('DEBUG', 'fetchAdStats_', `SKU ${sku}: кампаний ${campaigns.length}, период ${dateFromStr}–${dateToStr}`);
    if (!campaigns.length) return { adViews: 0, adClicks: 0, ctr: 0 };

    // 2. Запрос статистики товаров батчами по 10 кампаний (лимит API).
    // groupBy: GROUP_BY_SKU — разбивка по артикулам.
    // Ответ либо синхронный {rows:[...]}, либо async {UUID:"..."} → поллим.
    const CHUNK = 10;
    const allIds = campaigns.map(c => String(c.id));
    log_('DEBUG', 'fetchAdStats_', `SKU ${sku}: всего кампаний ${allIds.length}, батчей ${Math.ceil(allIds.length / CHUNK)}`);

    let allRows = [];
    for (let i = 0; i < allIds.length; i += CHUNK) {
      const chunk = allIds.slice(i, i + CHUNK);
      const statResp = UrlFetchApp.fetch(PERF_BASE + '/api/client/statistics', {
        method: 'post',
        contentType: 'application/json',
        headers: { 'Authorization': 'Bearer ' + token },
        payload: JSON.stringify({
          campaigns: chunk,
          groupBy:   'GROUP_BY_SKU',
          dateFrom:  dateFromStr,
          dateTo:    dateToStr
        }),
        muteHttpExceptions: true
      });

      const rawBody = statResp.getContentText();
      const batchN = Math.floor(i / CHUNK) + 1;
      log_('DEBUG', 'fetchAdStats_', `Батч ${batchN}: HTTP ${statResp.getResponseCode()} body[0:300]: ${rawBody.slice(0, 300)}`);

      if (statResp.getResponseCode() !== 200) continue;

      let parsed = JSON.parse(rawBody);
      if (parsed.UUID) {
        log_('DEBUG', 'fetchAdStats_', `Батч ${batchN}: async UUID=${parsed.UUID}, поллим...`);
        parsed = pollReport_(token, parsed.UUID);
        if (!parsed) continue;
      }

      const rows = extractAdRows_(parsed);
      log_('DEBUG', 'fetchAdStats_', `Батч ${batchN}: строк ${rows.length}${rows.length ? ', первая: ' + JSON.stringify(rows[0]).slice(0, 150) : ''}`);
      allRows = allRows.concat(rows);
    }

    // 3. Ищем строки с нашим SKU
    let totalViews = 0, totalClicks = 0;
    for (const r of allRows) {
      const rowSku = String(r.sku || r.sku_id || r.article || r.offer_id || r.ozonId || r.itemId || '');
      if (rowSku === String(sku)) {
        totalViews  += Number(r.views || r.impressions || r.shows || r.view || 0);
        totalClicks += Number(r.clicks || r.click || 0);
      }
    }
    if (allRows.length && !totalViews && !totalClicks) {
      log_('WARN', 'fetchAdStats_',
        `SKU ${sku} не найден среди ${allRows.length} строк. Поля первой строки: ${Object.keys(allRows[0]).join(', ')}`);
    }

    const ctr = totalViews > 0 ? (totalClicks / totalViews) * 100 : 0;
    return { adViews: totalViews, adClicks: totalClicks, ctr };
  } catch (e) {
    log_('WARN', 'fetchAdStats_', `SKU ${sku}: ${e.message}`);
    return { adViews: 0, adClicks: 0, ctr: 0 };
  }
}

/**
 * Опрашивает async-отчёт Performance API до готовности (до 10 раз, пауза 3 с).
 * Возвращает распарсенный JSON отчёта или null при таймауте.
 */
function pollReport_(token, uuid) {
  for (let i = 0; i < 10; i++) {
    Utilities.sleep(3000);
    const r = UrlFetchApp.fetch(PERF_BASE + '/api/client/statistics/report?UUID=' + uuid, {
      method: 'get',
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });
    const body = r.getContentText();
    log_('DEBUG', 'pollReport_', `Попытка ${i + 1}, HTTP ${r.getResponseCode()}, body[0:200]: ${body.slice(0, 200)}`);
    if (r.getResponseCode() !== 200) continue;
    const parsed = JSON.parse(body);
    // Отчёт готов, если нет поля state или state === 'OK'/'READY'/'DONE'
    const state = String(parsed.state || parsed.status || 'READY').toUpperCase();
    if (state === 'IN_PROGRESS' || state === 'PENDING' || state === 'PROCESSING') continue;
    return parsed;
  }
  log_('WARN', 'pollReport_', `UUID ${uuid}: таймаут (30 с), отчёт не готов`);
  return null;
}

/**
 * Пробует извлечь плоский массив строк из разных форматов ответа Performance API.
 * Форматы: {rows:[...]}, {items:[{rows:[...]}]}, {result:{rows:[...]}}, ...
 */
function extractAdRows_(obj) {
  if (!obj || typeof obj !== 'object') return [];
  // Плоский массив на верхнем уровне
  if (Array.isArray(obj.rows))   return obj.rows;
  if (Array.isArray(obj.items))  return obj.items;
  if (Array.isArray(obj.data))   return obj.data;
  // Вложенная структура: [{rows:[...]}, ...]
  if (obj.result) return extractAdRows_(obj.result);
  // Массив объектов, у которых внутри rows (по кампании)
  const keys = Object.keys(obj);
  for (const k of keys) {
    if (Array.isArray(obj[k]) && obj[k].length && typeof obj[k][0] === 'object') {
      const inner = [];
      obj[k].forEach(item => {
        const sub = extractAdRows_(item);
        sub.forEach(r => inner.push(r));
      });
      if (inner.length) return inner;
    }
  }
  return [];
}
