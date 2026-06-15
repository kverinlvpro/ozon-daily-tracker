/**
 * НАСТРОЙКИ, АККАУНТЫ, АРТИКУЛЫ
 * ==============================
 *
 * Листы артикулов — по одному на кабинет:
 *   «Артикулы — ИП Иванова»  (PRODUCTS_PREFIX + account.name)
 *   «Артикулы — ООО ТД Краски Дизайн»
 *
 * Колонки каждого листа артикулов:
 *   A: SKU (offer_id) — менеджер вводит
 *   B: Название       — заполняется автоматически кнопкой «Подтянуть названия»
 *   C: Менеджер       — менеджер вводит своё имя
 *   D: Статус         — «активен» / «пауза»
 */

function getSetting_(key) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET.SETTINGS);
  if (!sh || sh.getLastRow() < 2) return null;
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  for (const [k, v] of data) {
    if (k === key) return v;
  }
  return null;
}

/**
 * Активные кабинеты Ozon.
 * Колонки листа «Аккаунты»:
 *  A ID | B Название | C Seller Client-Id | D Seller Api-Key |
 *  E Performance: client_id|secret | F Активен (TRUE/FALSE)
 */
function getAccounts_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET.ACCOUNTS);
  if (!sh || sh.getLastRow() < 2) return [];
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues();
  return rows
    .filter(r => r[0] && r[5] === true)
    .map(r => ({
      id:                  String(r[0]),
      name:                String(r[1]),
      sellerClientId:      String(r[2]),
      sellerApiKey:        String(r[3]),
      performanceClientId: String(r[4]).split('|')[0] || '',
      performanceSecret:   String(r[4]).split('|')[1] || ''
    }));
}

/**
 * Активные артикулы под трекингом.
 *
 * Читает из листов «Артикулы — [AccountName]» (по одному на кабинет).
 * accountId определяется по имени листа → имени аккаунта.
 *
 * Обратная совместимость: если новые листы не найдены, читает старый
 * единый лист «Артикулы» (колонки A=SKU, B=Название, C=ID аккаунта, D=Менеджер, E=Статус).
 */
function getActiveProducts_() {
  const ss       = SpreadsheetApp.getActive();
  const accounts = getAccounts_();

  // Карта: имя кабинета → account.id
  const nameToId = {};
  for (const acc of accounts) nameToId[acc.name] = acc.id;

  const products = [];

  // ── Новый формат: листы «Артикулы — [AccountName]» ─────────────────────
  for (const sh of ss.getSheets()) {
    const shName = sh.getName();
    if (!shName.startsWith(PRODUCTS_PREFIX)) continue;

    const accountName = shName.slice(PRODUCTS_PREFIX.length);
    const accountId   = nameToId[accountName];
    if (!accountId) continue;   // лист есть, но аккаунта нет → пропускаем

    if (sh.getLastRow() < 2) continue;
    // Колонки: A=SKU, B=Название, C=Менеджер, D=Статус
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
    for (const r of rows) {
      if (!r[0] || r[3] !== 'активен') continue;
      products.push({
        sku:         String(r[0]).trim(),
        name:        String(r[1]),
        manager:     String(r[2]),
        status:      String(r[3]),
        accountId:   accountId,
        accountName: accountName
      });
    }
  }

  if (products.length) return products;

  // ── Fallback: старый единый лист «Артикулы» ────────────────────────────
  const sh = ss.getSheetByName(SHEET.PRODUCTS);
  if (sh && sh.getLastRow() >= 2) {
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues();
    for (const r of rows) {
      if (!r[0] || r[4] !== 'активен') continue;
      products.push({
        sku:       String(r[0]).trim(),
        name:      String(r[1]),
        manager:   String(r[3]),
        status:    String(r[4]),
        accountId: String(r[2]),
        addedAt:   r[5]
      });
    }
  }
  return products;
}

/** Артикулы конкретного кабинета (в порядке листа «Артикулы — ...»). */
function getProductsForAccount_(products, accountId) {
  return products.filter(p => p.accountId === accountId);
}


// ════════════════════════════════════════════════════════════════
//  АВТОЗАПОЛНЕНИЕ НАЗВАНИЙ АРТИКУЛОВ ИЗ OZON API
// ════════════════════════════════════════════════════════════════

/**
 * Проходит по всем листам «Артикулы — [AccountName]»,
 * находит строки с SKU но без названия и запрашивает Ozon API.
 * Заполненные названия записывает обратно в столбец B.
 *
 * Безопасно: не трогает строки, у которых название уже есть.
 */
function lookupProductNames_() {
  const ss       = SpreadsheetApp.getActive();
  const accounts = getAccounts_();
  const nameToAcc = {};
  for (const acc of accounts) nameToAcc[acc.name] = acc;

  let filled = 0, errors = 0;

  for (const sh of ss.getSheets()) {
    const shName = sh.getName();
    if (!shName.startsWith(PRODUCTS_PREFIX)) continue;

    const accountName = shName.slice(PRODUCTS_PREFIX.length);
    const account = nameToAcc[accountName];
    if (!account) continue;

    if (sh.getLastRow() < 2) continue;
    const nRows = sh.getLastRow() - 1;
    const rows  = sh.getRange(2, 1, nRows, 2).getValues();  // A=SKU, B=Название

    for (let i = 0; i < rows.length; i++) {
      const sku  = String(rows[i][0] || '').trim();
      const name = String(rows[i][1] || '').trim();
      if (!sku || name) continue;  // нет SKU или название уже есть — пропускаем

      try {
        const details = fetchProductDetails_(account, sku);
        const prodName = details.name || '';
        if (prodName) {
          sh.getRange(i + 2, 2).setValue(prodName);  // столбец B
          filled++;
          log_('INFO', 'lookupProductNames_',
            `${shName}: SKU ${sku} → «${prodName.slice(0, 50)}»`);
        }
        Utilities.sleep(300);  // щадим rate limit
      } catch (e) {
        log_('WARN', 'lookupProductNames_',
          `${shName}: SKU ${sku} ошибка: ${e.message}`);
        errors++;
      }
    }
  }

  flushLog_();
  SpreadsheetApp.getActive().toast(
    `Названия: заполнено ${filled}, ошибок ${errors}`,
    '🔄 Готово', 5
  );
}

function lookupProductNamesMenu() {
  lookupProductNames_();
}
