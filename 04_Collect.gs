/**
 * СБОР МЕТРИК ЗА ДЕНЬ
 * ===================
 * coverRecord_   — частичная запись с обложкой за указанный день.
 * metricsRecord_ — частичная запись с метриками за указанный день.
 * Обе сливаются в _Данные через mergeUpsertRecords_ (не затирая друг друга).
 */

const DAY_MS = 24 * 60 * 60 * 1000;

/** Идентификационные поля записи (общие для обложки и метрик). */
function identity_(product, account, details) {
  return {
    accountId:   account.id,
    accountName: account.name,
    sku:         product.sku,
    numericSku:  details.sku,
    name:        details.name || product.name
  };
}

/**
 * Частичная запись «обложка за сегодня».
 * Содержит только идентификацию + обложку — метрики этого дня дозреют завтра.
 */
function coverRecord_(product, account, details, dateForCover) {
  const coverUrl = details.primaryImage || '';
  const rec = identity_(product, account, details);
  rec.date      = dateKey_(dateForCover);
  rec.coverUrl  = coverUrl;
  rec.coverHash = hashStr_(coverUrl);
  return rec;
}

/**
 * Частичная запись «метрики за день metricDate».
 * Обложку НЕ трогает (она уже записана в этот день ранее).
 *
 * Формулы конверсий — из стабильных абсолютных метрик (не процентных от Ozon):
 *   CTR               = hits_view_pdp  / hits_view_search  × 100
 *   Конверсия в корзину = hits_tocart  / hits_view_pdp     × 100
 *   CR корзина→заказ  = ordered_units  / hits_tocart       × 100
 */
/**
 * Собирает запись метрик за день.
 * @param adStats       — CTR из Drive/Performance API для этого SKU (или null)
 * @param prefetchedAn  — из fetchAllAnalyticsBatch_() по numericSku (undefined → fallback)
 * @param prefetchedBuy — из fetchAllBuyoutsBatch_() по offer_id (undefined → fallback)
 */
function metricsRecord_(product, account, details, metricDate, adStats, prefetchedAn, prefetchedBuy) {
  let an, buy;
  if (prefetchedAn !== undefined) {
    // Batch-режим: данные пришли пакетным запросом — API не вызываем
    const zero = { totalViews: 0, searchViews: 0, pdpViews: 0,
                   sessSearch: 0, sessPdp: 0, toCart: 0, orders: 0, revenue: 0 };
    an  = prefetchedAn  || zero;
    buy = prefetchedBuy || { buyouts: 0, buyoutRevenue: 0 };
  } else {
    // Fallback: индивидуальные вызовы (testApiConnection и обратная совместимость)
    const dayFrom = startOfDay_(metricDate);
    const dayTo   = endOfDay_(metricDate);
    an  = fetchAnalytics_(account, details.sku, dayFrom, dayTo);
    buy = fetchBuyouts_(account, product.sku, dayFrom, dayTo);
  }

  // CTR = клики в карточку (hits_view_pdp) / все показы (hits_view)
  const ctr            = an.totalViews  > 0 ? (an.pdpViews  / an.totalViews)  * 100 : 0;
  // Рекламный CTR — из Performance API, передаётся готовым (один запрос на аккаунт)
  const adCtr          = adStats ? (adStats.ctr || 0) : 0;
  // Конверсия в корзину = корзины / клики в карточку
  const cartConvPct    = an.pdpViews    > 0 ? (an.toCart    / an.pdpViews)    * 100 : 0;
  // CR корзина→заказ = заказы / корзины
  const cartToOrderPct = an.toCart      > 0 ? (an.orders    / an.toCart)      * 100 : 0;
  const buyoutPct      = an.orders      > 0 ? (buy.buyouts  / an.orders)      * 100 : 0;

  const rec = identity_(product, account, details);
  rec.date           = dateKey_(metricDate);
  rec.ctr            = round2_(ctr);
  rec.adCtr          = round2_(adCtr);
  rec.views          = an.totalViews;   // «Показы» = все показы товара (знаменатель CTR)
  rec.pdpViews       = an.pdpViews;     // «Клики» = переходы в карточку (числитель CTR)
  rec.sessSearch     = an.sessSearch;   // диагностика
  rec.sessPdp        = an.sessPdp;      // диагностика
  rec.toCart         = an.toCart;
  rec.cartConvPct    = round2_(cartConvPct);
  rec.orders         = an.orders;
  rec.cartToOrderPct = round2_(cartToOrderPct);
  rec.buyouts        = buy.buyouts;
  rec.buyoutPct      = round2_(buyoutPct);
  rec.revenue        = an.revenue;
  return rec;
}


// ====== УТИЛИТЫ ======
function hashStr_(str) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_1, String(str || ''));
  return bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

function round2_(n) {
  return Math.round((Number(n) || 0) * 100) / 100;
}

function addDays_(date, days) {
  const d = new Date(date);
  d.setDate(d.getDate() + days);
  return d;
}

function startOfDay_(date) {
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);
  return d;
}

function endOfDay_(date) {
  const d = new Date(date);
  d.setHours(23, 59, 59, 999);
  return d;
}

/** Ключ дня 'yyyy-MM-dd' в часовом поясе таблицы — стабилен для группировки. */
function dateKey_(date) {
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  return Utilities.formatDate(new Date(date), tz, 'yyyy-MM-dd');
}

/** Короткая подпись колонки дня 'dd.MM'. */
function dateLabel_(dateKey) {
  // dateKey = 'yyyy-MM-dd'
  const [y, m, d] = String(dateKey).split('-');
  return `${d}.${m}`;
}

/** Разница в днях между двумя ключами дней 'yyyy-MM-dd'. */
function daysBetweenKeys_(fromKey, toKey) {
  const a = new Date(fromKey + 'T00:00:00');
  const b = new Date(toKey + 'T00:00:00');
  return Math.round((b - a) / DAY_MS);
}
