/**
 * АНАЛИЗ И ПРЕДУПРЕЖДЕНИЯ
 * ======================
 * analyzeSku_ — единая логика, которую используют и рендер (подсветка),
 * и лист «⚠️ Предупреждения».
 *
 * Два типа предупреждений:
 *  1) Эффект смены обложки: после последней смены ключевые конверсии
 *     (CTR, в корзину, корзина→заказ, % выкупа) упали или выросли больше порога.
 *  2) Залежавшаяся обложка: не менялась дольше stale_cover_days.
 */

// Метрики, по которым оцениваем эффект смены обложки
const EFFECT_METRICS = [
  { key: 'ctr',            name: 'CTR' },
  { key: 'cartConvPct',    name: 'конв. в корзину' },
  { key: 'cartToOrderPct', name: 'CR корзина→заказ' },
  { key: 'buyoutPct',      name: '% выкупа' }
];

function alertSettings_() {
  return {
    staleDays:  Number(getSetting_(SETTING_KEYS.STALE_DAYS))  || 14,
    alertPct:   Number(getSetting_(SETTING_KEYS.ALERT_PCT))   || 10,
    beforeDays: Number(getSetting_(SETTING_KEYS.BEFORE_DAYS)) || 7,
    afterDays:  Number(getSetting_(SETTING_KEYS.AFTER_DAYS))  || 7
  };
}

/**
 * Анализирует историю одного артикула.
 * records — записи этого SKU (любой порядок), todayKey — 'yyyy-MM-dd'.
 */
function analyzeSku_(records, settings, todayKey) {
  const recs = records.slice().sort((a, b) => String(a.date).localeCompare(String(b.date)));
  const result = {
    changeDateKeys: {},   // dateKey -> true (дни смены обложки)
    lastChangeKey:  null,
    daysSinceChange: null,
    stale: false,
    alerts: []            // [{type, direction, message, dateKey}]
  };
  if (!recs.length) return result;

  // 1) Находим дни смены обложки
  const changes = [];
  let prev = null;
  for (const rec of recs) {
    if (prev && rec.coverHash && prev.coverHash && rec.coverHash !== prev.coverHash) {
      changes.push(rec);
      result.changeDateKeys[rec.date] = true;
    }
    prev = rec;
  }

  const firstKey = recs[0].date;
  result.lastChangeKey = changes.length ? changes[changes.length - 1].date : firstKey;
  result.daysSinceChange = daysBetweenKeys_(result.lastChangeKey, todayKey);

  // 2) Залежавшаяся обложка
  if (result.daysSinceChange > settings.staleDays) {
    result.stale = true;
    const since = changes.length ? `последней смены (${dateLabel_(result.lastChangeKey)})`
                                 : `начала наблюдения (${dateLabel_(firstKey)})`;
    result.alerts.push({
      type: 'stale',
      direction: 'warn',
      dateKey: result.lastChangeKey,
      message: `Обложка не менялась ${result.daysSinceChange} дн. с ${since}`
    });
  }

  // 3) Эффект последней смены обложки
  if (changes.length) {
    const ev = changes[changes.length - 1];
    const evIdx = recs.findIndex(r => r.date === ev.date);
    const before = recs.slice(Math.max(0, evIdx - settings.beforeDays), evIdx);
    const after  = recs.slice(evIdx, evIdx + settings.afterDays);

    if (before.length && after.length) {
      const moved = [];
      let anyDown = false, anyUp = false;
      for (const m of EFFECT_METRICS) {
        const b = avg_(before, m.key);
        const a = avg_(after, m.key);
        if (b <= 0) continue;
        const delta = ((a - b) / b) * 100;
        if (Math.abs(delta) >= settings.alertPct) {
          const arrow = delta > 0 ? '↑' : '↓';
          moved.push(`${m.name} ${arrow}${Math.abs(delta).toFixed(0)}% (${b.toFixed(2)}→${a.toFixed(2)})`);
          if (delta > 0) anyUp = true; else anyDown = true;
        }
      }
      if (moved.length) {
        const direction = anyDown && !anyUp ? 'down' : (anyUp && !anyDown ? 'up' : 'mixed');
        const head = direction === 'down' ? '📉 После смены обложки упало: '
                   : direction === 'up'   ? '📈 После смены обложки выросло: '
                   : '🔀 После смены обложки изменилось: ';
        result.alerts.push({
          type: 'effect',
          direction: direction,
          dateKey: ev.date,
          message: head + moved.join('; ')
        });
      }
    }
  }

  return result;
}

function avg_(records, key) {
  if (!records.length) return 0;
  let sum = 0, n = 0;
  for (const r of records) {
    const v = Number(r[key]);
    if (!isNaN(v)) { sum += v; n++; }
  }
  return n ? sum / n : 0;
}


/**
 * Перестраивает лист «⚠️ Предупреждения» по всем артикулам.
 */
function rebuildAlerts_(accounts, products) {
  const ss = SpreadsheetApp.getActive();
  const sh = ensureSheet_(ss, SHEET.ALERTS);
  sh.clear();

  const header = ['Кабинет', 'Артикул', 'Название', 'Менеджер', 'Тип', 'Предупреждение', 'Дата'];
  sh.getRange(1, 1, 1, header.length).setValues([header]);
  formatHeader_(sh, header.length);

  const settings = alertSettings_();
  const todayKey = dateKey_(new Date());
  const data = readAllData_();

  // Группируем данные по SKU (в рамках аккаунта)
  const bySku = {};
  for (const rec of data) {
    const k = rec.accountId + '|' + rec.sku;
    (bySku[k] = bySku[k] || []).push(rec);
  }

  const rows = [];
  for (const product of products) {
    const account = accounts.find(a => a.id === product.accountId);
    const recs = bySku[product.accountId + '|' + product.sku] || [];
    if (!recs.length) continue;

    const an = analyzeSku_(recs, settings, todayKey);
    for (const al of an.alerts) {
      rows.push([
        account ? account.name : product.accountId,
        product.sku,
        product.name || (recs[recs.length - 1] && recs[recs.length - 1].name) || '',
        product.manager || '',
        al.type === 'stale' ? '🕒 залежалась' : (al.direction === 'down' ? '📉 падение' : al.direction === 'up' ? '📈 рост' : '🔀 смешанно'),
        al.message,
        dateLabel_(al.dateKey)
      ]);
    }
  }

  if (rows.length) {
    sh.getRange(2, 1, rows.length, header.length).setValues(rows);
    // Подсветка по типу
    for (let i = 0; i < rows.length; i++) {
      const type = rows[i][4];
      let color = '#fff2cc'; // жёлтый по умолчанию (залежалась/смешанно)
      if (type.indexOf('падение') >= 0) color = '#f4cccc';      // красноватый
      else if (type.indexOf('рост') >= 0) color = '#d9ead3';    // зеленоватый
      sh.getRange(i + 2, 1, 1, header.length).setBackground(color);
    }
  } else {
    sh.getRange(2, 1).setValue('Предупреждений нет 🎉');
  }

  sh.setColumnWidth(1, 140);
  sh.setColumnWidth(2, 120);
  sh.setColumnWidth(3, 220);
  sh.setColumnWidth(4, 120);
  sh.setColumnWidth(5, 120);
  sh.setColumnWidth(6, 520);
  sh.setColumnWidth(7, 80);
  sh.setFrozenRows(1);

  log_('INFO', 'rebuildAlerts_', `Предупреждений: ${rows.length}`);
}
