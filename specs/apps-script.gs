/**
 * Спецификации менеджеров — хранилище для страницы «Спецификации» и Telegram-бот.
 * Google Apps Script в аккаунте Аиды. Как установить — specs/README.md.
 *
 * Что делает:
 *  - хранит спецификации в Google Таблице «Спецификации менеджеров (сайт)», файлы — в папке на Google Диске;
 *  - отдаёт и принимает данные для страницы https://damombaeva-dev.github.io/sobaka-korgi/specs/;
 *  - пишет менеджеру в Telegram, когда добавлена спецификация;
 *  - каждый понедельник в 9:00 (Москва) присылает план: просрочено, отгрузить на неделе, дальше, брони из Остатков.
 */

/* ===== НАСТРОЙКИ (заполнить один раз) ===== */
const BOT_TOKEN = '';   // токен бота от @BotFather, например '1234567890:AAE...'
const PIN = '';         // код для изменений на сайте; пусто — менять может любой, у кого есть ссылка
/* ========================================== */

const OSTATKI_ID = '18f2ag6IcCpKuBrG67i9qDSIGqBbz7B8UwWpDvYFDvNg'; // Google Таблица остатков
const MANAGERS = { olga: 'Ольга', anna: 'Анна' };                   // ключ ссылки → имя менеджера
const ALL = 'all';                                                   // ключ для «копии всего» (руководитель)
const TZ = 'Europe/Moscow';
const SITE = 'https://damombaeva-dev.github.io/sobaka-korgi/specs/';
const COLS = ['id', 'mgr', 'client', 'product', 'qty', 'supplier', 'supSpec', 'ourSpec', 'signed', 'from', 'to',
  'note', 'fileId', 'fileName', 'fileUrl', 'ostKey', 'status', 'createdAt', 'doneAt'];
const MONTHS = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь'];
const START = { y: 2026, m: 11 };

/* ---------- установка: запустить один раз вручную ---------- */
function setup() {
  const sh = sheet_();
  folder_();
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('pollTelegram').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('weeklyDigest').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(9).inTimezone(TZ).create();
  let bot = '(токен бота не указан)';
  if (BOT_TOKEN) {
    tg_('deleteWebhook', {});
    const me = tg_('getMe', {});
    if (!me || !me.ok) throw new Error('Telegram не принял токен. Проверьте BOT_TOKEN.');
    props_().setProperty('BOT_NAME', me.result.username);
    bot = '@' + me.result.username;
  }
  Logger.log('Готово. Таблица: ' + sh.getParent().getUrl());
  Logger.log('Бот: ' + bot);
  Logger.log('Теперь: Развернуть → Новое развертывание → Веб-приложение (Запуск от: Я, Доступ: Все).');
}

/* ---------- веб-приложение ---------- */
function doGet() {
  try { pollTelegram(); } catch (e) { }
  return json_(state_());
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad' }); }
  if (PIN && String(body.pin || '') !== PIN) return json_({ ok: false, error: 'pin' });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const a = body.action;
    if (a === 'save') save_(body.item || {}, body.file);
    else if (a === 'done') setStatus_(body.id, 'done');
    else if (a === 'restore') setStatus_(body.id, 'active');
    else if (a === 'remove') remove_(body.id);
    else if (a === 'test') testMessage_(body.mgr);
    else if (a !== 'list') return json_({ ok: false, error: 'bad' });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
  return json_(state_());
}

function state_() {
  const ch = chats_();
  const linked = {};
  Object.keys(MANAGERS).concat([ALL]).forEach(k => linked[k] = (ch[k] || []).length);
  return { ok: true, items: readAll_().map(x => { delete x._row; return x; }), bot: props_().getProperty('BOT_NAME') || '', linked: linked, pin: !!PIN };
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- хранилище ---------- */
function props_() { return PropertiesService.getScriptProperties(); }

function sheet_() {
  const p = props_();
  let ss = null;
  const id = p.getProperty('SHEET_ID');
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) {
    ss = SpreadsheetApp.create('Спецификации менеджеров (сайт)');
    p.setProperty('SHEET_ID', ss.getId());
    const sh = ss.getSheets()[0];
    sh.setName('Спецификации');
    sh.getRange(1, 1, 1000, COLS.length).setNumberFormat('@');
    sh.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return ss.getSheetByName('Спецификации') || ss.getSheets()[0];
}

function folder_() {
  const p = props_();
  const id = p.getProperty('FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { } }
  const f = DriveApp.createFolder('Спецификации менеджеров (файлы)');
  p.setProperty('FOLDER_ID', f.getId());
  return f;
}

function readAll_() {
  const sh = sheet_();
  const v = sh.getDataRange().getValues();
  const h = v.shift() || [];
  const out = [];
  v.forEach((r, i) => {
    if (!r[0]) return;
    const o = { _row: i + 2 };
    h.forEach((k, j) => {
      let x = r[j];
      if (x instanceof Date) x = Utilities.formatDate(x, TZ, 'yyyy-MM-dd');
      o[k] = x === null || x === undefined ? '' : String(x);
    });
    out.push(o);
  });
  return out;
}

function write_(row, item) {
  const sh = sheet_();
  sh.getRange(row, 1, 1, COLS.length).setNumberFormat('@').setValues([COLS.map(k => item[k] == null ? '' : String(item[k]))]);
}

function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'); }

function save_(item, file) {
  const all = readAll_();
  const old = item.id ? all.find(x => x.id === item.id) : null;
  const rec = {};
  COLS.forEach(k => rec[k] = old ? old[k] : '');
  ['mgr', 'client', 'product', 'qty', 'supplier', 'supSpec', 'ourSpec', 'signed', 'from', 'to', 'note', 'ostKey']
    .forEach(k => { if (item[k] !== undefined) rec[k] = String(item[k]).slice(0, 500); });
  if (!rec.mgr || !rec.client || !rec.product) throw new Error('Заполните менеджера, клиента и товар');
  if (file && file.b64) {
    const blob = Utilities.newBlob(Utilities.base64Decode(file.b64), file.mime || 'application/octet-stream', file.name || 'спецификация');
    const f = folder_().createFile(blob);
    f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    if (old && old.fileId) { try { DriveApp.getFileById(old.fileId).setTrashed(true); } catch (e) { } }
    rec.fileId = f.getId(); rec.fileName = f.getName(); rec.fileUrl = f.getUrl();
  }
  if (old) {
    write_(old._row, rec);
  } else {
    rec.id = Utilities.getUuid().slice(0, 8);
    rec.status = 'active';
    rec.createdAt = now_();
    write_(sheet_().getLastRow() + 1, rec);
    notifyNew_(rec);
  }
}

function setStatus_(id, status) {
  const it = readAll_().find(x => x.id === id);
  if (!it) throw new Error('Спецификация не найдена');
  it.status = status;
  it.doneAt = status === 'done' ? now_() : '';
  write_(it._row, it);
}

function remove_(id) {
  const it = readAll_().find(x => x.id === id);
  if (!it) return;
  if (it.fileId) { try { DriveApp.getFileById(it.fileId).setTrashed(true); } catch (e) { } }
  sheet_().deleteRow(it._row);
}

/* ---------- даты ---------- */
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function addDays_(s, n) { const d = new Date(s + 'T12:00:00Z'); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); }
function weekStart_(s) { const d = new Date(s + 'T12:00:00Z'); return addDays_(s, -((d.getUTCDay() + 6) % 7)); }
function ru_(s) { if (!s) return '—'; const p = String(s).split('-'); return p.length === 3 ? p[2] + '.' + p[1] + '.' + p[0] : s; }
function ddmm_(s) { const p = String(s).split('-'); return p.length === 3 ? p[2] + '.' + p[1] : s; }
function kg_(q) { const n = parseFloat(String(q).replace(/\s/g, '').replace(',', '.')); return isNaN(n) ? (q || '') : Math.round(n).toLocaleString('ru-RU').replace(/,/g, ' ') + ' кг'; }
function esc_(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

/** группа спецификации: overdue | week | soon | later | done */
function bucket_(it, today) {
  if (it.status === 'done') return 'done';
  const ws = weekStart_(today), we = addDays_(ws, 6);
  const end = it.to || it.from;
  if (end && end < today) return 'overdue';
  const sd = it.from || it.to;
  if (!sd || sd <= we) return 'week';
  if (sd <= addDays_(we, 14)) return 'soon';
  return 'later';
}

/* ---------- Telegram ---------- */
function tg_(method, payload) {
  if (!BOT_TOKEN) return null;
  const r = UrlFetchApp.fetch('https://api.telegram.org/bot' + BOT_TOKEN + '/' + method, {
    method: 'post', contentType: 'application/json', payload: JSON.stringify(payload || {}), muteHttpExceptions: true
  });
  try { return JSON.parse(r.getContentText()); } catch (e) { return null; }
}

function send_(chat, text, buttons) {
  if (buttons && text.length < 3800) {
    tg_('sendMessage', { chat_id: chat, text: text, parse_mode: 'HTML', disable_web_page_preview: true,
      reply_markup: { inline_keyboard: [buttons.map(b => ({ text: b[0], url: b[1] }))] } });
    return;
  }
  const lines = String(text).split('\n');
  let buf = '';
  lines.forEach(l => {
    if ((buf + '\n' + l).length > 3800) { tg_('sendMessage', { chat_id: chat, text: buf, parse_mode: 'HTML', disable_web_page_preview: true }); buf = ''; }
    buf = buf ? buf + '\n' + l : l;
  });
  if (buf) tg_('sendMessage', { chat_id: chat, text: buf, parse_mode: 'HTML', disable_web_page_preview: true });
}

function chats_() { try { return JSON.parse(props_().getProperty('CHATS') || '{}'); } catch (e) { return {}; } }
function mgrKey_(name) { const n = String(name || '').toLowerCase(); return Object.keys(MANAGERS).find(k => n.indexOf(MANAGERS[k].toLowerCase().slice(0, 3)) === 0) || ''; }
function chatsFor_(mgr) { const ch = chats_(); const k = mgrKey_(mgr); return [].concat(k ? ch[k] || [] : [], ch[ALL] || []).filter((x, i, a) => a.indexOf(x) === i); }

/** Читает новые сообщения боту: /start olga — подключить Ольгу, /week — план сейчас, /stop — отключить. Запускается по таймеру. */
function pollTelegram() {
  if (!BOT_TOKEN) return;
  try { checkExcelSpecs_(ostatkiRows_()); } catch (e) { Logger.log(e); }
  const p = props_();
  const off = +(p.getProperty('TG_OFFSET') || 0);
  const r = tg_('getUpdates', { offset: off, timeout: 0, allowed_updates: ['message'] });
  if (!r || !r.ok || !r.result.length) return;
  const ch = chats_();
  let last = off;
  r.result.forEach(u => {
    last = u.update_id + 1;
    const m = u.message;
    if (!m || !m.text) return;
    const id = m.chat.id, t = m.text.trim();
    const st = t.match(/^\/start(?:\s+(\S+))?/i);
    if (st) {
      const k = String(st[1] || '').toLowerCase();
      if (!MANAGERS[k] && k !== ALL) { send_(id, 'Здравствуйте! Чтобы получать напоминания, откройте ссылку «Подключить Telegram» на странице спецификаций:\n' + SITE); return; }
      ch[k] = ch[k] || [];
      if (ch[k].indexOf(id) < 0) ch[k].push(id);
      const who = k === ALL ? 'копии всех уведомлений (Ольга и Анна)' : 'напоминания для: ' + MANAGERS[k];
      send_(id, '✅ Подключено: ' + who + '.\n\nБуду писать, когда добавлена новая спецификация, и каждый понедельник в 9:00 присылать план отгрузок и броней.\n\n/week — план прямо сейчас\n/stop — отключить');
    } else if (/^\/stop/i.test(t)) {
      Object.keys(ch).forEach(k => ch[k] = (ch[k] || []).filter(x => x !== id));
      send_(id, 'Напоминания отключены. Подключить снова: ' + SITE);
    } else if (/^\/(week|plan|неделя)/i.test(t)) {
      const keys = Object.keys(ch).filter(k => (ch[k] || []).indexOf(id) >= 0);
      if (!keys.length) send_(id, 'Этот чат ещё не подключён. Откройте ссылку «Подключить Telegram» на ' + SITE);
      p.setProperty('CHATS', JSON.stringify(ch));
      keys.forEach(k => (k === ALL ? Object.keys(MANAGERS) : [k]).forEach(mk => send_(id, digest_(MANAGERS[mk]), [['Все спецификации', SITE]])));
    } else {
      send_(id, 'Команды: /week — план на неделю, /stop — отключить напоминания.\nСпецификации: ' + SITE);
    }
  });
  p.setProperty('CHATS', JSON.stringify(ch));
  p.setProperty('TG_OFFSET', String(last));
}

const MON_GEN = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'];
const MON_S = ['янв', 'фев', 'мар', 'апр', 'мая', 'июн', 'июл', 'авг', 'сен', 'окт', 'ноя', 'дек'];
function ymd_(s) { const p = String(s).split('-').map(Number); return { y: p[0], m: p[1], d: p[2] }; }
function lastDay_(y, m) { return new Date(Date.UTC(y, m, 0)).getUTCDate(); }

/** Срок словами: «14–18 декабря», «28 ноября – 5 декабря», «весь декабрь», «до 18 декабря» */
function winText_(from, to, withYear) {
  if (!from && !to) return 'срок не указан';
  const a = ymd_(from || to), b = ymd_(to || from);
  const yr = withYear ? ' ' + b.y : '';
  if (a.y === b.y && a.m === b.m && a.d === 1 && b.d === lastDay_(b.y, b.m)) return 'весь ' + MONTHS[b.m - 1] + yr;
  if (from === to || !from) return (from ? '' : 'до ') + b.d + ' ' + MON_GEN[b.m - 1] + yr;
  if (a.y === b.y && a.m === b.m) return a.d + '–' + b.d + ' ' + MON_GEN[b.m - 1] + yr;
  return a.d + ' ' + MON_GEN[a.m - 1] + ' – ' + b.d + ' ' + MON_GEN[b.m - 1] + yr;
}
function winShort_(from, to) {
  const a = ymd_(from || to), b = ymd_(to || from);
  if (a.y === b.y && a.m === b.m && a.d === 1 && b.d === lastDay_(b.y, b.m)) return MONTHS[b.m - 1];
  if (from === to || !from) return b.d + ' ' + MON_S[b.m - 1];
  if (a.m === b.m) return a.d + '–' + b.d + ' ' + MON_S[b.m - 1];
  return a.d + ' ' + MON_S[a.m - 1] + ' – ' + b.d + ' ' + MON_S[b.m - 1];
}
function daysText_(it) {
  const t = today_(), end = it.to || it.from, st = it.from || it.to;
  if (!end) return '';
  const dd = (a, b) => Math.round((new Date(b + 'T12:00:00Z') - new Date(a + 'T12:00:00Z')) / 864e5);
  if (end < t) return '⚠️ просрочено на ' + dd(end, t) + ' дн.';
  if (st > t) return 'начало через ' + dd(t, st) + ' дн.';
  const n = dd(t, end);
  return n === 0 ? '❗ последний день' : 'осталось ' + n + ' дн.';
}
function specText_(it) {
  const s = String(it.ourSpec || '').trim();
  if (!s && !it.signed) return '';
  if (/^подпис/i.test(s)) return 'подписана' + s.replace(/^подписана/i, '').replace(/^,\s*/, ', ');
  return (s ? (/^\s*№/.test(s) ? s : '№ ' + s) : 'подписана') + (it.signed ? ' от ' + ru_(it.signed) : '');
}

/** Одна спецификация в плане: клиент, товар и объём, срок — одной строкой */
function line_(it) {
  return '▫️ <b>' + esc_(it.client) + '</b> · ' + esc_(it.product) + (it.qty ? ', ' + kg_(it.qty) : '') + ' · ' + winShort_(it.from, it.to);
}

/** Уведомление: крупно товар+объём, срок, клиент; остальное одной строкой курсивом */
function notifyNew_(it) {
  const chats = chatsFor_(it.mgr);
  if (!chats.length) return;
  const days = daysText_(it);
  const small = [esc_(it.mgr), specText_(it) ? 'спец. ' + esc_(specText_(it)) : '',
    it.supplier ? esc_(it.supplier) + (it.supSpec ? ' (спец. ' + esc_(it.supSpec) + ')' : '') : '', it.note ? esc_(it.note) : '']
    .filter(Boolean).join(' · ');
  const t = '🏢 <b>' + esc_(it.client) + '</b>\n' +
    '📦 <b>' + esc_(it.product) + (it.qty ? ' — ' + kg_(it.qty) : '') + '</b>\n' +
    '🚚 <b>' + winText_(it.from, it.to, false) + '</b>' + (days ? '  <i>' + days + '</i>' : '') +
    (small ? '\n<blockquote expandable>' + small.split(' · ').join('\n') + '</blockquote>' : '');
  const btn = [];
  if (it.fileUrl) btn.push(['📎 Файл', it.fileUrl]);
  btn.push(['Все спецификации', SITE]);
  chats.forEach(c => send_(c, t, btn));
}

function testMessage_(mgr) {
  const chats = chatsFor_(mgr || '');
  if (!chats.length) throw new Error('Для ' + (mgr || 'этого менеджера') + ' Telegram ещё не подключён');
  chats.forEach(c => send_(c, '🔔 Проверка: напоминания по спецификациям работают' + (mgr ? ' (' + esc_(mgr) + ')' : '') + '.'));
}

/* ---------- Остатки: брони и подписанные спецификации (столбцы I, J, K) ---------- */
function dIso_(v, yr) {
  if (v === '' || v === null || v === undefined) return '';
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  const m = String(v).match(/(\d{1,2})[.\/](\d{1,2})(?:[.\/](\d{2,4}))?/);
  if (!m) return '';
  let y = m[3] ? +m[3] : yr; if (y < 100) y += 2000;
  return y + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
}
function isoMonday_(y, w) { const j4 = new Date(Date.UTC(y, 0, 4)); const m = new Date(j4); m.setUTCDate(4 - ((j4.getUTCDay() + 6) % 7) + (w - 1) * 7); return m.toISOString().slice(0, 10); }
/** номер недели (49, «нед 49», «49 неделя», «W49», «49/2027») → {from: понедельник, to: воскресенье} */
function weekOf_(v, yr, mon) {
  let w = null, y = null;
  if (typeof v === 'number') { if (v >= 1 && v <= 53 && Math.floor(v) === v) w = v; }
  else if (typeof v === 'string') {
    const rg = v.trim().match(/^(?:w|нед\.?|недели|неделя)?\s*(\d{1,2})\s*[-–—]\s*(\d{1,2})\s*(?:нед\.?|недели|неделя|н)?$/i);
    if (rg && +rg[1] >= 1 && +rg[2] <= 53 && +rg[1] <= +rg[2]) { const a = weekOf_(+rg[1], yr, mon), b = weekOf_(+rg[2], yr, mon); return a && b ? { from: a.from, to: b.to } : null; }
    const m = v.trim().match(/^(?:w|нед\.?|неделя)?\s*(\d{1,2})\s*(?:нед\.?|неделя|н)?\s*(?:[\/.]\s*(\d{2,4}))?$/i);
    if (m && +m[1] >= 1 && +m[1] <= 53) { w = +m[1]; if (m[2]) { y = +m[2]; if (y < 100) y += 2000; } }
  }
  if (w === null) return null;
  if (!y) { y = yr; if (isoMonday_(y, w) < addDays_(yr + '-' + ('0' + mon).slice(-2) + '-01', -60)) y++; }
  const a = isoMonday_(y, w);
  return { from: a, to: addDays_(a, 6) };
}
function monthNum_(v) { if (typeof v === 'number') return v >= 1 && v <= 12 && Math.floor(v) === v ? v : 0; const m = String(v === null || v === undefined ? '' : v).trim().match(/^(\d{1,2})(?:\s*мес\.?)?$/i); return m && +m[1] >= 1 && +m[1] <= 12 ? +m[1] : 0; }
function addMonthsG_(s, n) { const p = s.split('-').map(Number); return new Date(Date.UTC(p[0], p[1] - 1 + n, 1)).toISOString().slice(0, 10); }
function cellTxt_(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'dd.MM.yyyy') : String(v === null || v === undefined ? '' : v).trim(); }

/** Все строки броней из Остатков (вкладки «Месяц ГГГГ» от ноября 2026). Ключ строки совпадает с ключом на сайте. */
function ostatkiRows_() {
  const out = [];
  let ss;
  try { ss = SpreadsheetApp.openById(OSTATKI_ID); } catch (e) { return out; }
  ss.getSheets().forEach(sh => {
    if (sh.isSheetHidden()) return;
    const month = String(sh.getName()).trim();
    const mm = month.toLowerCase().match(/^([а-яё]+)\s+(\d{4})$/);
    if (!mm) return;
    const mi = MONTHS.indexOf(mm[1]), yr = +mm[2];
    if (mi < 0 || yr * 12 + mi + 1 < START.y * 12 + START.m) return;
    const last = sh.getLastRow();
    if (last < 1) return;
    const rows = sh.getRange(1, 1, last, Math.min(11, sh.getMaxColumns())).getValues();
    const isNum = v => typeof v === 'number' || /^-?\d+([.,]\d+)?$/.test(String(v).replace(/\s/g, ''));
    const blank = r => !r.slice(0, 8).some(v => v !== '' && v !== null && String(v).trim() !== '');
    const isHead = r => typeof r[0] === 'string' && r[0].trim() && !isNum(r[0]) && !r.slice(1, 8).some(v => v !== '' && v !== null && String(v).trim() !== '');
    const nextNonBlank = i => { for (let j = i + 1; j < rows.length; j++) if (!blank(rows[j])) return rows[j]; return null; };
    let sup = '', head = '';
    rows.forEach((r, i) => {
      if (blank(r)) return;
      if (isHead(r)) {
        const t = r[0].trim(), nx = nextNonBlank(i);
        if (/^\|?-{2,}/.test(t) || (nx && isHead(nx))) { sup = t.replace(/^\|?-+\s*|\s*-+\|?\s*$/g, '').trim() || t; head = ''; return; }
        head = t; return;
      }
      if (!head) return;
      const client = cellTxt_(r[1]);
      const qn = typeof r[2] === 'number' ? r[2] : (isNum(r[2]) ? parseFloat(String(r[2]).replace(/\s/g, '').replace(',', '.')) : null);
      if (!client && qn === null) return;
      const ship = r[4], res = r[3];
      const done = ship instanceof Date, cx = !done && /отмен/i.test(cellTxt_(ship) + ' ' + cellTxt_(res));
      const pm = head.match(/^(.*?)(?:\s*\(|\s+-{1,3}\s|\s+\/\/|$)/);
      /* I = месяц отгрузки цифрой 1–12 (= спецификация подписана); J = номер недели или дата, если срок точнее */
      const mo = monthNum_(r[8]);
      let sFrom = '', sTo = '', spec = '';
      if (mo) {
        let y = yr; if (mo < mi + 1 - 2) y++;
        const ms = y + '-' + ('0' + mo).slice(-2) + '-01';
        const me = addDays_(addMonthsG_(ms, 1), -1);
        const wj = weekOf_(r[9], y, mo), dj = dIso_(r[9], y);
        if (wj) {
          sFrom = wj.from; sTo = wj.to;
          if (sFrom > addDays_(me, 40)) { const a = isoMonday_(y - 1, +String(r[9]).match(/\d+/)[0]); sFrom = a; sTo = addDays_(a, 6); }
          spec = 'подписана, нед. ' + cellTxt_(r[9]).replace(/^\D+/, '');
        } else if (dj) { sFrom = sTo = dj; spec = 'подписана'; }
        else { sFrom = ms; sTo = me; spec = 'подписана, ' + MONTHS[mo - 1] + ' ' + y; }
      }
      out.push({ month: month, sup: sup.replace(/\s*\(.*$/, '').trim() || sup, product: (pm && pm[1].trim()) || head, client: client, q: qn === null ? '' : qn,
        res: res instanceof Date ? Utilities.formatDate(res, TZ, 'dd.MM') : String(res || '').trim(),
        terms: [cellTxt_(r[5]), cellTxt_(r[6])].filter(Boolean).join(' · '), mgr: cellTxt_(r[7]), done: done, cx: cx,
        spec: spec, from: sFrom, to: sTo || sFrom, key: [month, sup, head, client, qn === null ? '' : qn].join('|') });
    });
  });
  return out;
}

/** Подписанные спецификации из Остатков в том же виде, что и с сайта */
function excelSpecs_(rows) {
  return rows.filter(b => b.spec && !b.cx).map(b => ({ id: 'x', mgr: b.mgr, client: b.client, product: b.product, qty: b.q, supplier: b.sup,
    ourSpec: b.spec, from: b.from, to: b.to, note: b.terms, ostKey: b.key, status: b.done ? 'done' : 'active', fileUrl: '' }));
}

/** Новые подписанные спецификации в Остатках → сообщение менеджеру. Запускается вместе с опросом бота. */
function checkExcelSpecs_(rows) {
  const p = props_();
  const h = k => Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, k, Utilities.Charset.UTF_8)).slice(0, 10);
  const specs = excelSpecs_(rows).filter(x => x.status !== 'done');
  const cur = specs.map(x => h(x.ostKey + '|' + x.ourSpec));
  const raw = p.getProperty('SEEN_X');
  if (raw !== null) {
    const seen = new Set(raw ? raw.split(',') : []);
    specs.forEach((x, i) => { if (!seen.has(cur[i])) notifyNew_(Object.assign({}, x, { ourSpec: x.ourSpec, signed: '' })); });
  }
  p.setProperty('SEEN_X', cur.join(','));
}

/* ---------- еженедельный план ---------- */
function digest_(mgr) {
  const today = today_();
  const ws = weekStart_(today), we = addDays_(ws, 6);
  const rows = ostatkiRows_();
  const own = readAll_();
  const ownKeys = new Set(own.map(x => x.ostKey).filter(Boolean));
  const xs = excelSpecs_(rows).filter(x => !ownKeys.has(x.ostKey));
  const byKey = {}; excelSpecs_(rows).forEach(x => byKey[x.ostKey] = x);
  const all = own.map(x => { const b = x.ostKey && byKey[x.ostKey]; return b ? Object.assign({}, x, { from: b.from || x.from, to: b.to || x.to, status: b.status === 'done' ? 'done' : x.status, ourSpec: b.ourSpec }) : x; }).concat(xs);
  const mine = all.filter(x => x.status !== 'done' && mgrKey_(x.mgr) === mgrKey_(mgr))
    .sort((a, b) => String(a.to || a.from).localeCompare(String(b.to || b.from)));
  const g = { overdue: [], week: [], soon: [], later: [] };
  mine.forEach(x => g[bucket_(x, today)].push(x));
  const kgSum = L0 => { const n = L0.reduce((a, x) => a + (parseFloat(String(x.qty || x.q || 0).replace(/\s/g, '').replace(',', '.')) || 0), 0); return n ? ' · ' + kg_(n) : ''; };
  const L = ['📋 <b>' + esc_(mgr) + ', план на неделю</b>', winText_(ws, we, false), ''];
  const sec = (icon, title, list, empty) => {
    if (!list.length && !empty) return;
    L.push(icon + ' <b>' + title + '</b>' + (list.length ? ' (' + list.length + kgSum(list) + ')' : ''));
    if (!list.length) L.push('<i>нет</i>');
    list.forEach(x => L.push(line_(x)));
    L.push('');
  };
  sec('⚠️', 'Срок прошёл', g.overdue);
  sec('🚚', 'Отгрузить на этой неделе', g.week, true);
  sec('🗓', 'Следующие 2 недели', g.soon);
  sec('⏳', 'Позже', g.later);
  const pref = String(mgr).toLowerCase().slice(0, 3);
  const b = rows.filter(x => !x.done && !x.cx && !x.spec && x.mgr.toLowerCase().slice(0, 3) === pref);
  if (b.length) {
    L.push('📦 <b>Брони без спецификации</b> (' + b.length + ')');
    b.forEach(x => L.push('▫️ <b>' + esc_(x.client) + '</b> · ' + esc_(x.product) + (x.q !== '' ? ', ' + kg_(x.q) : '') + ' · ' + esc_(x.month.toLowerCase())));
    L.push('');
  }
  return L.join('\n').trim();
}
/** Понедельник 9:00: каждой — свой план, руководителю — оба. */
function weeklyDigest() {
  const ch = chats_();
  Object.keys(MANAGERS).forEach(k => {
    const text = digest_(MANAGERS[k]);
    (ch[k] || []).forEach(c => send_(c, text, [['Все спецификации', SITE]]));
    (ch[ALL] || []).forEach(c => { if ((ch[k] || []).indexOf(c) < 0) send_(c, text, [['Все спецификации', SITE]]); });
  });
}
