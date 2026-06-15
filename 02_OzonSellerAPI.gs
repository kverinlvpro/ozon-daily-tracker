/**
 * OZON SELLER API
 * ===============
 * Документация: https://docs.ozon.ru/api/seller/
 *
 * Эндпоинты, которые мы используем:
 *  - POST /v3/product/info/list  — детали товара (включая массив images)
 *  - POST /v1/analytics/data     — аналитика (показы, корзина, заказы)
 *  - POST /v2/posting/fbo/list   — заказы FBO для расчёта выкупов
 *  - POST /v3/posting/fbs/list   — заказы FBS
 *
 * Скопировано из проекта ozon-ab без изменений.
 */

const OZON_BASE = 'https://api-seller.ozon.ru';

/**
 * Универсальный POST к Seller API с обработкой ошибок и ретраем.
 */
function ozonSellerRequest_(account, path, payload) {
  const url = OZON_BASE + path;
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Client-Id': account.sellerClientId,
      'Api-Key':   account.sellerApiKey
    },
    payload: JSON.stringify(payload || {}),
    muteHttpExceptions: true
  };

  for (let attempt = 1; attempt <= 3; attempt++) {
    const resp = UrlFetchApp.fetch(url, options);
    const code = resp.getResponseCode();
    const body = resp.getContentText();

    if (code === 200) {
      return JSON.parse(body);
    }
    if (code === 429 || code >= 500) {
      // Rate limit или серверная ошибка — ждём и ретраим
      Utilities.sleep(2000 * attempt);
      continue;
    }
    throw new Error(`Ozon API ${path}: HTTP ${code} — ${body}`);
  }
  throw new Error(`Ozon API ${path}: 3 попытки исчерпаны`);
}

/**
 * Получает детали товара по offer_id (артикулу продавца).
 * Возвращает {images, primaryImage, name, description, sku, productId}
 */
function fetchProductDetails_(account, identifier) {
  identifier = String(identifier);
  const isSku = /^\d+$/.test(identifier) && identifier.length > 8;

  function extractItems_(data) {
    const items = (data.result && data.result.items) || data.items;
    return items && items.length > 0 ? items : null;
  }

  const primaryPayload  = isSku ? { sku: [Number(identifier)] } : { offer_id: [identifier] };
  const primaryData     = ozonSellerRequest_(account, '/v3/product/info/list', primaryPayload);
  let items             = extractItems_(primaryData);

  let fallbackPayload = null;
  let fallbackData    = null;
  if (!items) {
    if (isSku) {
      fallbackPayload = { offer_id: [identifier] };
    } else if (/^\d+$/.test(identifier)) {
      fallbackPayload = { sku: [Number(identifier)] };
    }
    if (fallbackPayload) {
      fallbackData = ozonSellerRequest_(account, '/v3/product/info/list', fallbackPayload);
      items        = extractItems_(fallbackData);
    }
  }

  if (!items) {
    log_('DEBUG', 'fetchProductDetails_', `primary payload: ${JSON.stringify(primaryPayload)}`);
    log_('DEBUG', 'fetchProductDetails_', `primary response: ${JSON.stringify(primaryData)}`);
    if (fallbackPayload) {
      log_('DEBUG', 'fetchProductDetails_', `fallback payload: ${JSON.stringify(fallbackPayload)}`);
      log_('DEBUG', 'fetchProductDetails_', `fallback response: ${JSON.stringify(fallbackData)}`);
    }
    throw new Error(`${identifier} не найден в кабинете ни по offer_id, ни по sku`);
  }

  const item = items[0];
  return {
    productId:    item.id,
    sku:          item.sku,
    offerId:      item.offer_id,
    name:         item.name,
    images:       item.images || [],
    primaryImage: (item.primary_image && item.primary_image[0]) || (item.images && item.images[0]) || '',
    images360:    item.images360 || [],
    color:        item.color_image || ''
  };
}

/**
 * Аналитика Seller API: показы, корзина, заказы по SKU за период.
 * dateFrom, dateTo — объекты Date.
 *
 * ВАЖНО: эндпоинт /v1/analytics/data ограничен 60 запросов/минуту,
 * данные доступны с задержкой ~1 день.
 *
 * numericSku — числовой Ozon-SKU (из fetchProductDetails_.sku),
 * не offer_id продавца. Фильтр по key:'sku' работает только с ним.
 */
// Порядок метрик в запросе аналитики (индексы строго соответствуют ответу row.metrics)
const ANALYTICS_METRICS = [
  'hits_view',           // 0 — все показы товара (поиск + рекомендации + похожие) = знаменатель CTR
  'hits_view_search',    // 1 — показы только в поиске/каталоге (диагностика)
  'hits_view_pdp',       // 2 — переходы в карточку = «Клики» = числитель CTR
  'session_view_search', // 3 — показы в поиске (по сессиям, диагностика)
  'session_view_pdp',    // 4 — переходы в карточку (по сессиям, диагностика)
  'hits_tocart_pdp',     // 5 — добавления в корзину с карточки
  'ordered_units',       // 6 — заказано единиц
  'revenue'              // 7 — выручка
  // adv_view_pdp / adv_click_pdp не существуют в Seller Analytics — рекламный CTR берётся из Performance API
];

function fetchAnalytics_(account, numericSku, dateFrom, dateTo) {
  // ВАЖНО: форматируем в таймзоне таблицы, а НЕ в UTC.
  // Если форматировать в UTC, то полночь по Екатеринбургу (UTC+5) = 19:00 пред.дня UTC
  // → date_from сдвигается на сутки назад → Ozon возвращает данные за два дня сразу.
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const payload = {
    date_from:  Utilities.formatDate(dateFrom, tz, 'yyyy-MM-dd'),
    date_to:    Utilities.formatDate(dateTo,   tz, 'yyyy-MM-dd'),
    metrics:    ANALYTICS_METRICS,
    dimension:  ['sku'],
    filters:    [{ key: 'sku', value: String(numericSku) }],
    sort:       [{ key: 'ordered_units', order: 'DESC' }],
    limit:      1,
    offset:     0
  };

  const data = ozonSellerRequest_(account, '/v1/analytics/data', payload);
  const row = data.result && data.result.data && data.result.data[0];

  if (!row) {
    return { totalViews: 0, searchViews: 0, pdpViews: 0, sessSearch: 0, sessPdp: 0, toCart: 0, orders: 0, revenue: 0 };
  }
  const m = row.metrics;
  const named = ANALYTICS_METRICS.map((name, i) => `${name}=${m[i]}`).join(', ');
  log_('DEBUG', 'fetchAnalytics_', `SKU ${numericSku}: ${named}`);

  return {
    totalViews:  m[0] || 0,
    searchViews: m[1] || 0,
    pdpViews:    m[2] || 0,
    sessSearch:  m[3] || 0,
    sessPdp:     m[4] || 0,
    toCart:      m[5] || 0,
    orders:      m[6] || 0,
    revenue:     m[7] || 0
  };
}

// ══════════════════════════════════════════════════════════════════
//  ПАКЕТНЫЕ ЗАПРОСЫ — 4 вызова на кабинет вместо 4×N на каждый SKU
// ══════════════════════════════════════════════════════════════════

/**
 * BATCH: детали товаров для всех offer_id кабинета за 1 запрос (до 200 SKU в чанке).
 * Возвращает {offerId → detailsObject} — тот же объект что и fetchProductDetails_.
 */
function fetchAllProductDetailsBatch_(account, offerIds) {
  const result = {};
  if (!offerIds.length) return result;
  const CHUNK = 200;
  for (let i = 0; i < offerIds.length; i += CHUNK) {
    const chunk = offerIds.slice(i, i + CHUNK);
    try {
      const data  = ozonSellerRequest_(account, '/v3/product/info/list', { offer_id: chunk });
      const items = (data.result && data.result.items) || data.items || [];
      for (const item of items) {
        result[item.offer_id] = {
          productId:    item.id,
          sku:          item.sku,          // числовой Ozon-SKU
          offerId:      item.offer_id,
          name:         item.name,
          images:       item.images || [],
          primaryImage: (item.primary_image && item.primary_image[0])
                          || (item.images && item.images[0]) || '',
          images360:    item.images360 || [],
          color:        item.color_image  || ''
        };
      }
    } catch (e) {
      log_('WARN', 'fetchAllProductDetailsBatch_', `chunk[${i}]: ${e.message}`);
    }
  }
  log_('DEBUG', 'fetchAllProductDetailsBatch_',
    `Получено деталей: ${Object.keys(result).length} / ${offerIds.length}`);
  return result;
}

/**
 * BATCH: аналитика за день для ВСЕХ товаров кабинета — 1 запрос вместо N.
 * Возвращает {numericSku (string) → analyticsObject}.
 *
 * Ключ — числовой Ozon-SKU (item.sku из product/info/list).
 * При > 1000 товаров добавлено автоматическое листание (offset).
 */
function fetchAllAnalyticsBatch_(account, dateFrom, dateTo) {
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const dfStr = Utilities.formatDate(dateFrom, tz, 'yyyy-MM-dd');
  const dtStr = Utilities.formatDate(dateTo,   tz, 'yyyy-MM-dd');

  const result = {};
  let offset = 0;
  const LIMIT = 1000;

  do {
    const payload = {
      date_from: dfStr,
      date_to:   dtStr,
      metrics:   ANALYTICS_METRICS,
      dimension: ['sku'],
      filters:   [],          // без фильтра → все товары кабинета
      sort:      [{ key: 'hits_view', order: 'DESC' }],
      limit:     LIMIT,
      offset:    offset
    };
    const data = ozonSellerRequest_(account, '/v1/analytics/data', payload);
    const rows = (data.result && data.result.data) || [];

    for (const row of rows) {
      const dim        = row.dimensions && row.dimensions[0];
      const numericSku = dim && String(dim.id);
      if (!numericSku) continue;
      const m = row.metrics;
      result[numericSku] = {
        totalViews:  m[0] || 0,
        searchViews: m[1] || 0,
        pdpViews:    m[2] || 0,
        sessSearch:  m[3] || 0,
        sessPdp:     m[4] || 0,
        toCart:      m[5] || 0,
        orders:      m[6] || 0,
        revenue:     m[7] || 0
      };
    }

    if (rows.length < LIMIT) break;   // последняя страница
    offset += LIMIT;
    Utilities.sleep(500);             // пауза между страницами
  } while (true);

  log_('DEBUG', 'fetchAllAnalyticsBatch_',
    `Получено аналитики по ${Object.keys(result).length} SKU за ${dfStr}`);
  return result;
}

/**
 * BATCH: выкупы за день для всех отслеживаемых offer_id — 2 вызова (FBO + FBS) вместо 2×N.
 * Возвращает {offerId → {buyouts, buyoutRevenue}}.
 */
function fetchAllBuyoutsBatch_(account, offerIds, dateFrom, dateTo) {
  // Инициализируем нулями для всех отслеживаемых SKU
  const result = {};
  for (const id of offerIds) result[id] = { buyouts: 0, buyoutRevenue: 0 };

  const trackSet   = new Set(offerIds.map(String));
  const baseFilter = {
    since:  dateFrom.toISOString(),
    to:     dateTo.toISOString(),
    status: 'delivered'
  };

  // FBO
  try {
    const fbo = ozonSellerRequest_(account, '/v2/posting/fbo/list', {
      dir: 'ASC', filter: baseFilter, limit: 1000, offset: 0,
      with: { financial_data: true }
    });
    for (const p of (fbo.result || [])) {
      for (const prod of (p.products || [])) {
        const id = String(prod.offer_id);
        if (!trackSet.has(id)) continue;
        result[id].buyouts       += Number(prod.quantity) || 0;
        result[id].buyoutRevenue += (Number(prod.price) || 0) * (Number(prod.quantity) || 0);
      }
    }
  } catch (e) { log_('WARN', 'fetchAllBuyoutsBatch_', `FBO: ${e.message}`); }

  // FBS
  try {
    const fbs = ozonSellerRequest_(account, '/v3/posting/fbs/list', {
      dir: 'ASC', filter: baseFilter, limit: 1000, offset: 0,
      with: { financial_data: true }
    });
    for (const p of ((fbs.result && fbs.result.postings) || [])) {
      for (const prod of (p.products || [])) {
        const id = String(prod.offer_id);
        if (!trackSet.has(id)) continue;
        result[id].buyouts       += Number(prod.quantity) || 0;
        result[id].buyoutRevenue += (Number(prod.price) || 0) * (Number(prod.quantity) || 0);
      }
    }
  } catch (e) { log_('WARN', 'fetchAllBuyoutsBatch_', `FBS: ${e.message}`); }

  return result;
}


/**
 * FBO+FBS заказы — нужны для подсчёта выкупов (статус delivered) и суммы выкупов.
 * Возвращает {buyouts: число, buyoutRevenue: сумма}.
 *
 * Тонкость: Ozon считает выкупом доставленный заказ, не отменённый покупателем после доставки.
 * Поэтому фильтруем по статусу 'delivered' в окне.
 */
function fetchBuyouts_(account, offerId, dateFrom, dateTo) {
  let buyouts = 0;
  let buyoutRevenue = 0;

  const baseFilter = {
    since:  dateFrom.toISOString(),
    to:     dateTo.toISOString(),
    status: 'delivered'
  };

  // FBO
  try {
    const fbo = ozonSellerRequest_(account, '/v2/posting/fbo/list', {
      dir:    'ASC',
      filter: baseFilter,
      limit:  1000,
      offset: 0,
      with:   { financial_data: true }
    });
    for (const p of (fbo.result || [])) {
      for (const prod of (p.products || [])) {
        // Сравниваем по offer_id — это артикул продавца ("T2-00001541")
        if (String(prod.offer_id) === String(offerId)) {
          buyouts       += Number(prod.quantity) || 0;
          buyoutRevenue += (Number(prod.price) || 0) * (Number(prod.quantity) || 0);
        }
      }
    }
  } catch (e) {
    log_('WARN', 'fetchBuyouts_', `FBO для ${offerId}: ${e.message}`);
  }

  // FBS
  try {
    const fbs = ozonSellerRequest_(account, '/v3/posting/fbs/list', {
      dir:    'ASC',
      filter: baseFilter,
      limit:  1000,
      offset: 0,
      with:   { financial_data: true }
    });
    for (const p of (fbs.result && fbs.result.postings || [])) {
      for (const prod of (p.products || [])) {
        if (String(prod.offer_id) === String(offerId)) {
          buyouts       += Number(prod.quantity) || 0;
          buyoutRevenue += (Number(prod.price) || 0) * (Number(prod.quantity) || 0);
        }
      }
    }
  } catch (e) {
    log_('WARN', 'fetchBuyouts_', `FBS для ${offerId}: ${e.message}`);
  }

  return { buyouts, buyoutRevenue };
}
