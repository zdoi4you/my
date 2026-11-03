/**
 * My ЗДОИ4You – приемане на „Заявки за съдействие“.
 *
 * Скриптът е свързан с таблицата на Сдружението (Разширения → Apps Script) и се публикува
 * като уеб приложение: „Изпълнение като: аз“, „Достъп: всеки“. Приложението изпраща заявката
 * заедно с Google токена на потребителя; скриптът проверява токена при Google и записва
 * имейла, с който е влязъл потребителят. Без валиден вход заявка не се приема.
 *
 * Настройка: попълнете CONFIG, изпълнете веднъж setup() и публикувайте (Внедряване → Ново внедряване).
 */
const CONFIG = {
  CLIENT_ID: '838853307992-dv64og2iil0q6fqjgu293006ulc50r5q.apps.googleusercontent.com',
  SHEET: 'Заявки',
  FOLDER_ID: '',                       // папка в Drive за приложенията; празно – setup() създава нова
  NOTIFY: 'legal.crisis@zdoi4you.eu',  // уведомление за нова заявка
  MAX_PER_DAY: 3,                      // заявки от един акаунт за 24 часа
  MAX_FILE: 8 * 1024 * 1024,           // байта на файл
  MAX_TOTAL: 20 * 1024 * 1024,         // байта общо
  TYPES: ['application/pdf', 'image/jpeg', 'image/png', 'application/msword',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'application/vnd.oasis.opendocument.text']
};

const HEAD = ['№', 'Получена', 'Имейл (Google)', 'Име', 'Телефон', 'Населено място', 'Задължен субект', 'Вид отказ',
  'Вх. № на заявлението', 'Дата на регистриране', 'Дата на решението', 'Решението е получено на', 'Срок за обжалване',
  'Описание', 'Приложения', 'Статус', 'Бележка към заявителя', 'Отговорник'];
const COL = Object.fromEntries(HEAD.map((h, i) => [h, i + 1]));
const STATUS = ['Нова', 'В преглед', 'Одобрена', 'Договор изпратен', 'Заведено дело', 'Отказана – изпратени насоки', 'Приключена'];
const KIND = { explicit: 'Изричен отказ', partial: 'Частичен отказ', silent: 'Мълчалив отказ' };

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CONFIG.SHEET) || ss.insertSheet(CONFIG.SHEET);
  sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.getRange(2, COL['Статус'], sh.getMaxRows() - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(STATUS, true).setAllowInvalid(false).build());
  const props = PropertiesService.getScriptProperties();
  if (!CONFIG.FOLDER_ID && !props.getProperty('FOLDER_ID')) {
    props.setProperty('FOLDER_ID', DriveApp.createFolder('ЗДОИ4You – заявки за съдействие').getId());
  }
  // уведомяване на заявителя при смяна на статуса
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'onStatusEdit').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('onStatusEdit').forSpreadsheet(ss).onEdit().create();
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const email = verify_(req.token);
    if (req.action === 'list') return out_({ ok: true, items: list_(email) });
    if (req.action === 'submit') return out_(submit_(email, req.data || {}, req.files || []));
    return out_({ ok: false, error: 'action' });
  } catch (err) {
    return out_({ ok: false, error: String(err.message || err) });
  }
}

function doGet() { return out_({ ok: true, service: 'zdoi4you-legal' }); }

/* Google access token → имейл на потребителя (само за нашия OAuth клиент) */
function verify_(token) {
  if (!token) throw new Error('auth');
  const r = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?access_token=' + encodeURIComponent(token), { muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error('auth');
  const t = JSON.parse(r.getContentText());
  if ((t.aud !== CONFIG.CLIENT_ID && t.azp !== CONFIG.CLIENT_ID) || !t.email || String(t.email_verified) !== 'true') throw new Error('auth');
  return String(t.email).toLowerCase();
}

function submit_(email, d, files) {
  const kind = KIND[d.kind] ? d.kind : '';
  if (!kind || !txt_(d.name, 120) || !txt_(d.body, 200) || !txt_(d.reg, 10)) return { ok: false, error: 'fields' };
  const need = kind === 'silent' ? ['app', 'reg'] : ['app', 'reg', 'dec'];
  if (need.some(k => !files.some(f => f.kind === k))) return { ok: false, error: 'files' };
  if (files.length > 10) return { ok: false, error: 'size' };
  let total = 0;
  const blobs = files.map(f => {
    if (CONFIG.TYPES.indexOf(f.type) < 0) throw new Error('type');
    const bytes = Utilities.base64Decode(String(f.b64 || ''));
    total += bytes.length;
    if (bytes.length > CONFIG.MAX_FILE || total > CONFIG.MAX_TOTAL) throw new Error('size');
    const pre = { app: 'Заявление', reg: 'Регистрация', dec: 'Решение', other: 'Друго' }[f.kind] || 'Друго';
    return Utilities.newBlob(bytes, f.type, pre + ' – ' + String(f.name || 'файл').replace(/[\\/:*?"<>|]/g, '_').slice(0, 100));
  });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET);
    const rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, HEAD.length).getValues() : [];
    const day = Date.now() - 864e5;
    if (rows.filter(r => r[COL['Имейл (Google)'] - 1] === email && new Date(r[1]).getTime() > day).length >= CONFIG.MAX_PER_DAY) {
      return { ok: false, error: 'limit' };
    }
    const year = new Date().getFullYear();
    const seq = rows.filter(r => String(r[0]).indexOf('ЗС-' + year + '-') === 0).length + 1;
    const id = 'ЗС-' + year + '-' + ('000' + seq).slice(-4);
    const root = DriveApp.getFolderById(CONFIG.FOLDER_ID || PropertiesService.getScriptProperties().getProperty('FOLDER_ID'));
    const folder = root.createFolder(id + ' – ' + txt_(d.name, 60));
    blobs.forEach(b => folder.createFile(b));
    const row = [id, new Date(), email, txt_(d.name, 120), txt_(d.phone, 40), txt_(d.city, 120), txt_(d.body, 200), KIND[kind],
      txt_(d.no, 60), txt_(d.reg, 10), txt_(d.dec, 10), txt_(d.recv, 10), txt_(d.appeal, 10),
      txt_(d.desc, 4000), folder.getUrl(), STATUS[0], '', ''];
    sh.appendRow(row.map(safe_));

    MailApp.sendEmail({
      to: CONFIG.NOTIFY, replyTo: email,
      subject: 'Нова заявка за съдействие ' + id + ' – ' + KIND[kind],
      body: HEAD.slice(0, 15).map((h, i) => h + ': ' + (row[i] instanceof Date ? row[i].toLocaleString('bg-BG') : row[i])).join('\n')
    });
    MailApp.sendEmail({
      to: email, replyTo: CONFIG.NOTIFY, name: 'ЗДОИ4You',
      subject: 'Получихме вашата заявка за съдействие ' + id,
      body: 'Здравейте,\n\nПолучихме заявката ви за съдействие с № ' + id + ' (' + KIND[kind] + ', ' + txt_(d.body, 200) + ').\n\n' +
        'Екипът ни ще прецени правната основателност на случая и наличния капацитет и ще ви пише на този имейл. ' +
        'Статуса на заявката виждате и в приложението My ЗДОИ4You, раздел „Съдействие“.\n\n' +
        'Важно: заявката не спира сроковете за обжалване. Ако срокът изтича скоро, обжалвайте сами, за да не го пропуснете.\n\n' +
        'Сдружение „Просвета за демокрация“\n' + CONFIG.NOTIFY
    });
    return { ok: true, id: id };
  } finally {
    lock.releaseLock();
  }
}

function list_(email) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET);
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, HEAD.length).getValues()
    .filter(r => r[COL['Имейл (Google)'] - 1] === email)
    .map(r => ({ id: r[0], at: new Date(r[1]).getTime(), body: r[COL['Задължен субект'] - 1], kind: r[COL['Вид отказ'] - 1],
      status: r[COL['Статус'] - 1], note: r[COL['Бележка към заявителя'] - 1] }))
    .reverse();
}

/* имейл до заявителя, когато екипът смени статуса */
function onStatusEdit(e) {
  const r = e.range, sh = r.getSheet();
  if (sh.getName() !== CONFIG.SHEET || r.getColumn() !== COL['Статус'] || r.getRow() < 2 || r.getNumRows() !== 1) return;
  const v = sh.getRange(r.getRow(), 1, 1, HEAD.length).getValues()[0];
  const to = v[COL['Имейл (Google)'] - 1], st = v[COL['Статус'] - 1], note = v[COL['Бележка към заявителя'] - 1];
  if (!to || !st || st === e.oldValue) return;
  MailApp.sendEmail({
    to: to, replyTo: CONFIG.NOTIFY, name: 'ЗДОИ4You',
    subject: 'Заявка ' + v[0] + ': ' + st,
    body: 'Здравейте,\n\nСтатусът на заявката ви за съдействие № ' + v[0] + ' е „' + st + '“.' + (note ? '\n\n' + note : '') +
      '\n\nСдружение „Просвета за демокрация“\n' + CONFIG.NOTIFY
  });
}

const txt_ = (v, n) => String(v == null ? '' : v).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '').trim().slice(0, n);
/* да не се изпълняват като формули в таблицата */
const safe_ = v => typeof v === 'string' && /^[=+\-@]/.test(v) ? "'" + v : v;
const out_ = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
