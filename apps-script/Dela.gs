/**
 * ЗДОИ4You – Дела (https://dela.zdoi4you.eu): платформа на екипа за „Заявки за съдействие“.
 *
 * Файлът е в същия проект, свързан с таблицата „ЗДОИ4You – заявки за съдействие“, до Code.gs.
 * doPost в Code.gs препраща тук всички действия, които започват с „dela.“. Затова Дела използва
 * същия адрес на уеб приложението като My ЗДОИ4You.
 *
 * Сигурност: всяка заявка носи Google access token. Той се проверява при Google (издаден за нашия
 * OAuth клиент, имейлът е потвърден). Допускат се само активните членове от лист „Екип“, а правата
 * се проверяват тук, на сървъра.
 *
 * Настройка: изпълнете веднъж setupDela(). Той създава листовете „Екип“, „Коментари“ и „Дневник“,
 * а акаунтът, който го изпълнява, става администратор.
 */
const Dela = (() => {
  const CFG = {
    CLIENT_ID: '838853307992-dv64og2iil0q6fqjgu293006ulc50r5q.apps.googleusercontent.com',
    NOTIFY: 'legal.crisis@zdoi4you.eu',
    APP_URL: 'https://dela.zdoi4you.eu/',
    TZ: 'Europe/Sofia'
  };
  const SH = { REQ: 'Заявки', TEAM: 'Екип', COMM: 'Коментари', LOG: 'Дневник' };
  const F = {
    id: '№', at: 'Получена', email: 'Имейл (Google)', name: 'Име', phone: 'Телефон', city: 'Населено място',
    body: 'Задължен субект', kind: 'Вид отказ', no: 'Вх. № на заявлението', reg: 'Дата на регистриране',
    dec: 'Дата на решението', recv: 'Решението е получено на', appeal: 'Срок за обжалване', desc: 'Описание',
    files: 'Приложения', status: 'Статус', note: 'Бележка към заявителя', owner: 'Отговорник'
  };
  const STATUSES = ['Нова', 'В преглед', 'Одобрена', 'Договор изпратен', 'Заведено дело', 'Отказана – изпратени насоки', 'Приключена'];
  const ROLES = { 'Администратор': 3, 'Юрист': 2, 'Наблюдател': 1 };
  const TEAM_HEAD = ['Имейл', 'Име', 'Роля', 'Активен', 'Добавен', 'Добавен от'];
  const COMM_HEAD = ['Време', '№', 'Автор', 'Текст'];
  const LOG_HEAD = ['Време', 'Потребител', 'Действие', '№', 'Подробности'];
  /* минимална роля за всяко действие */
  const NEED = { me: 1, list: 1, get: 1, update: 2, comment: 2, team: 3, teamSave: 3, log: 3 };

  const ss = () => SpreadsheetApp.getActiveSpreadsheet();

  function setup() {
    const team = sheet(SH.TEAM, TEAM_HEAD);
    sheet(SH.COMM, COMM_HEAD);
    const log = sheet(SH.LOG, LOG_HEAD);
    team.getRange(2, 3, team.getMaxRows() - 1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(Object.keys(ROLES), true).setAllowInvalid(false).build());
    team.getRange(2, 4, team.getMaxRows() - 1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(['Да', 'Не'], true).setAllowInvalid(false).build());
    const me = String(Session.getEffectiveUser().getEmail()).toLowerCase();
    if (me && !members().some(m => m.email === me)) team.appendRow([me, '', 'Администратор', 'Да', new Date(), 'setup']);
    // листовете с права и одит може да редактира само собственикът на таблицата (и скриптът)
    [team, log].forEach(s => {
      const p = s.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0] || s.protect();
      p.setDescription('Управлява се от платформата Дела');
      p.removeEditors(p.getEditors());
      if (p.canDomainEdit()) p.setDomainEdit(false);
    });
  }

  /* req = { action: 'dela.<действие>', token, data } → обект за отговор */
  function handle(req) {
    try {
      const action = String(req.action).slice(5);
      const email = verify(req.token);
      const me = members().find(m => m.email === email && m.active);
      if (!me) { denied(email); return { ok: false, error: 'forbidden', email: email }; }
      const need = NEED[action];
      if (!need) return { ok: false, error: 'action' };
      if (ROLES[me.role] < need) return { ok: false, error: 'role' };
      const d = req.data || {};
      switch (action) {
        case 'me':
          log_(me.email, 'Вход', '', '');
          return { ok: true, me: me, statuses: STATUSES, roles: Object.keys(ROLES) };
        case 'list': return { ok: true, items: items().map(brief), team: teamPublic(), comments: commentCounts() };
        case 'get': return get(String(d.id || ''));
        case 'update': return update(me, d);
        case 'comment': return comment(me, d);
        case 'team': return { ok: true, team: members() };
        case 'teamSave': return teamSave(me, d);
        case 'log': return { ok: true, log: logList(400) };
      }
    } catch (err) {
      return { ok: false, error: String(err.message || err) };
    }
  }

  /* ---------- достъп ---------- */

  /* Google access token → потвърден имейл; резултатът се кешира до 10 мин. */
  function verify(token) {
    if (!token || typeof token !== 'string' || token.length > 4096) throw new Error('auth');
    const key = 't' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, token));
    const cache = CacheService.getScriptCache();
    const hit = cache.get(key);
    if (hit) return hit;
    const r = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?access_token=' + encodeURIComponent(token), { muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) throw new Error('auth');
    const t = JSON.parse(r.getContentText());
    if ((t.aud !== CFG.CLIENT_ID && t.azp !== CFG.CLIENT_ID) || !t.email || String(t.email_verified) !== 'true') throw new Error('auth');
    const email = String(t.email).toLowerCase();
    const ttl = Math.min(600, Number(t.expires_in) || 0);
    if (ttl > 30) cache.put(key, email, ttl);
    return email;
  }

  /* отказан достъп – записва се в дневника най-много веднъж на 10 мин. за акаунт */
  function denied(email) {
    const cache = CacheService.getScriptCache(), k = 'd' + email;
    if (cache.get(k)) return;
    cache.put(k, '1', 600);
    log_(email, 'Отказан достъп', '', '');
  }

  function members() {
    const sh = ss().getSheetByName(SH.TEAM);
    if (!sh || sh.getLastRow() < 2) return [];
    return sh.getRange(2, 1, sh.getLastRow() - 1, TEAM_HEAD.length).getValues()
      .map((r, i) => ({
        email: String(r[0]).trim().toLowerCase(), name: String(r[1] || ''), role: ROLES[r[2]] ? r[2] : 'Наблюдател',
        active: r[3] !== 'Не', added: date(r[4], true), by: String(r[5] || ''), _row: i + 2
      }))
      .filter(m => m.email);
  }

  const teamPublic = () => members().filter(m => m.active).map(m => ({ email: m.email, name: m.name, role: m.role }));

  /* ---------- заявки ---------- */

  function items() {
    const sh = ss().getSheetByName(SH.REQ);
    if (!sh || sh.getLastRow() < 2) return [];
    const v = sh.getDataRange().getValues();
    const ix = index(v[0]);
    return v.slice(1).map((r, i) => {
      const o = { _row: i + 2 };
      Object.keys(F).forEach(k => {
        const x = ix[k] == null ? '' : r[ix[k]];
        o[k] = k === 'at' ? date(x, true) : ['reg', 'dec', 'recv', 'appeal'].indexOf(k) >= 0 ? date(x) : String(x == null ? '' : x);
      });
      return o;
    }).filter(o => o.id);
  }

  function brief(o) {
    const b = {};
    ['id', 'at', 'email', 'name', 'city', 'body', 'kind', 'appeal', 'status', 'owner'].forEach(k => b[k] = o[k]);
    return b;
  }

  function get(id) {
    const it = items().find(o => o.id === id);
    if (!it) return { ok: false, error: 'notfound' };
    delete it._row;
    return {
      ok: true, item: it,
      comments: rows(SH.COMM, COMM_HEAD).filter(r => r[1] === id).map(r => ({ at: date(r[0], true), author: r[2], text: r[3] })),
      history: rows(SH.LOG, LOG_HEAD).filter(r => r[3] === id).map(r => ({ at: date(r[0], true), user: r[1], action: r[2], details: r[4] })).reverse()
    };
  }

  function update(me, d) {
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const sh = ss().getSheetByName(SH.REQ);
      const ix = index(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]);
      const it = items().find(o => o.id === String(d.id || ''));
      if (!it) return { ok: false, error: 'notfound' };
      const ch = {};
      if (d.status != null && d.status !== it.status) {
        if (STATUSES.indexOf(d.status) < 0) return { ok: false, error: 'status' };
        ch.status = d.status;
      }
      if (d.owner != null && d.owner !== it.owner) {
        if (d.owner && !teamPublic().some(m => m.email === d.owner)) return { ok: false, error: 'owner' };
        ch.owner = d.owner;
      }
      if (d.note != null && txt(d.note, 2000) !== it.note) ch.note = txt(d.note, 2000);
      if (!Object.keys(ch).length) return { ok: true, changed: false };
      Object.keys(ch).forEach(k => {
        if (ix[k] == null) throw new Error('column');
        sh.getRange(it._row, ix[k] + 1).setValue(safe(ch[k]));
      });
      const label = { status: 'Статус', owner: 'Отговорник', note: 'Бележка към заявителя' };
      Object.keys(ch).forEach(k => log_(me.email, 'Промяна: ' + label[k], it.id, k === 'note' ? ch[k] : (it[k] || '—') + ' → ' + (ch[k] || '—')));

      // промените от скрипт не задействат onStatusEdit, затова имейлът до заявителя се праща оттук
      const now = Object.assign({}, it, ch);
      let mailed = false;
      if (d.notify && it.email && (ch.status || ch.note)) {
        MailApp.sendEmail({
          to: it.email, replyTo: CFG.NOTIFY, name: 'ЗДОИ4You',
          subject: 'Заявка ' + it.id + ': ' + now.status,
          body: 'Здравейте,\n\nСтатусът на заявката ви за съдействие № ' + it.id + ' е „' + now.status + '“.' +
            (now.note ? '\n\n' + now.note : '') + '\n\nСдружение „Просвета за демокрация“\n' + CFG.NOTIFY
        });
        log_(me.email, 'Имейл до заявителя', it.id, now.status);
        mailed = true;
      }
      if (ch.owner && ch.owner !== me.email) {
        MailApp.sendEmail({
          to: ch.owner, name: 'ЗДОИ4You – Дела',
          subject: 'Възложена ви е заявка ' + it.id,
          body: me.email + ' ви възложи заявка за съдействие ' + it.id + ' (' + it.kind + ', ' + it.body + ').\n' +
            (it.appeal ? 'Срок за обжалване: ' + it.appeal + '\n' : '') +
            '\n' + CFG.APP_URL + '#/r/' + encodeURIComponent(it.id) + '\n'
        });
      }
      return { ok: true, changed: true, mailed: mailed };
    } finally {
      lock.releaseLock();
    }
  }

  function comment(me, d) {
    const id = String(d.id || ''), text = txt(d.text, 4000);
    if (!text) return { ok: false, error: 'fields' };
    if (!items().some(o => o.id === id)) return { ok: false, error: 'notfound' };
    ss().getSheetByName(SH.COMM).appendRow([new Date(), id, me.email, safe(text)]);
    log_(me.email, 'Коментар', id, '');
    return { ok: true };
  }

  function commentCounts() {
    const c = {};
    rows(SH.COMM, COMM_HEAD).forEach(r => c[r[1]] = (c[r[1]] || 0) + 1);
    return c;
  }

  /* ---------- екип ---------- */

  function teamSave(me, d) {
    const email = String(d.email || '').trim().toLowerCase();
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || email.length > 200) return { ok: false, error: 'email' };
    if (!ROLES[d.role]) return { ok: false, error: 'role' };
    const active = d.active !== false;
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const list = members();
      const cur = list.find(m => m.email === email);
      // никой не си сваля сам правата и винаги остава поне един активен администратор
      if (email === me.email && (d.role !== 'Администратор' || !active)) return { ok: false, error: 'self' };
      const admins = list.filter(m => m.active && m.role === 'Администратор' && m.email !== email).length;
      if (!admins && (d.role !== 'Администратор' || !active)) return { ok: false, error: 'lastadmin' };
      const sh = ss().getSheetByName(SH.TEAM);
      const name = txt(d.name, 120);
      if (cur) {
        sh.getRange(cur._row, 2, 1, 3).setValues([[safe(name), d.role, active ? 'Да' : 'Не']]);
        log_(me.email, 'Екип: промяна', '', email + ' – ' + d.role + (active ? '' : ', деактивиран'));
      } else {
        sh.appendRow([email, safe(name), d.role, active ? 'Да' : 'Не', new Date(), me.email]);
        log_(me.email, 'Екип: добавен', '', email + ' – ' + d.role);
      }
      return { ok: true, team: members() };
    } finally {
      lock.releaseLock();
    }
  }

  /* ---------- дневник ---------- */

  function log_(user, action, id, details) {
    ss().getSheetByName(SH.LOG).appendRow([new Date(), user, action, id, safe(txt(details, 500))]);
  }

  function logList(n) {
    return rows(SH.LOG, LOG_HEAD).slice(-n).reverse()
      .map(r => ({ at: date(r[0], true), user: r[1], action: r[2], id: r[3], details: r[4] }));
  }

  /* ---------- помощни ---------- */

  function sheet(name, head) {
    const sh = ss().getSheetByName(name) || ss().insertSheet(name);
    if (sh.getLastRow() < 1) sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
    return sh;
  }

  function rows(name, head) {
    const sh = ss().getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return [];
    return sh.getRange(2, 1, sh.getLastRow() - 1, head.length).getValues();
  }

  /* колоните се търсят по заглавие – редът им в таблицата може да се променя */
  function index(head) {
    const ix = {};
    Object.keys(F).forEach(k => { const i = head.indexOf(F[k]); if (i >= 0) ix[k] = i; });
    return ix;
  }

  function date(v, time) {
    if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, CFG.TZ, time ? "yyyy-MM-dd'T'HH:mm" : 'yyyy-MM-dd');
    return String(v == null ? '' : v);
  }

  const txt = (v, n) => String(v == null ? '' : v).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '').trim().slice(0, n);
  /* да не се изпълняват като формули в таблицата */
  const safe = v => typeof v === 'string' && /^[=+\-@]/.test(v) ? "'" + v : v;

  return { handle: handle, setup: setup };
})();

/* Изпълнете веднъж от редактора: създава листовете на Дела и ви прави администратор. */
function setupDela() { Dela.setup(); }
