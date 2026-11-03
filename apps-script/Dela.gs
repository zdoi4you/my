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
 * Роли:
 *   Суперадминистратор – всичко; трие преписки и членове на екипа; управлява администраторите.
 *   Администратор      – администрира преписките: заключва/отключва, насочва към отговорник;
 *                        добавя и променя юристи и наблюдатели; вижда дневника.
 *   Юрист              – обработва незаключени преписки: статус, бележка, корекция на данните, коментари;
 *                        създава ръчно заявки (получени по имейл, телефон, на място).
 *   Наблюдател         – само преглед.
 * Всички членове пишат съобщения до администраторите (лист „Съобщения“); администраторите отговарят.
 *
 * Сесия: 30 минути от входа, проверява се на сървъра (CacheService). Удължава се с потвърждение;
 * след изтичане токенът не се приема повече и е нужен нов вход с Google.
 *
 * Настройка: изпълнете setupDela(). Тя създава листовете „Екип“, „Коментари“, „Дневник“ и колоната
 * „Заключена“. Ако още няма суперадминистратор, акаунтът, който я изпълнява, става суперадминистратор.
 */
const Dela = (() => {
  const CFG = {
    CLIENT_ID: '838853307992-dv64og2iil0q6fqjgu293006ulc50r5q.apps.googleusercontent.com',
    NOTIFY: 'legal.crisis@zdoi4you.eu',
    APP_URL: 'https://dela.zdoi4you.eu/',
    TZ: 'Europe/Sofia',
    SESSION_MIN: 30
  };
  const SH = { REQ: 'Заявки', TEAM: 'Екип', COMM: 'Коментари', LOG: 'Дневник', MSG: 'Съобщения' };
  const F = {
    id: '№', at: 'Получена', email: 'Имейл (Google)', name: 'Име', phone: 'Телефон', city: 'Населено място',
    body: 'Задължен субект', kind: 'Вид отказ', no: 'Вх. № на заявлението', reg: 'Дата на регистриране',
    dec: 'Дата на решението', recv: 'Решението е получено на', appeal: 'Срок за обжалване', desc: 'Описание',
    files: 'Приложения', status: 'Статус', note: 'Бележка към заявителя', owner: 'Отговорник', locked: 'Заключена', src: 'Източник'
  };
  /* данни от подадената заявка, които юристите могат да коригират: [макс. дължина, тип] */
  const DATA = {
    name: [120], phone: [40], city: [120], body: [200], kind: [0, 'kind'], no: [60],
    reg: [10, 'date'], dec: [10, 'date'], recv: [10, 'date'], appeal: [10, 'date'], desc: [4000]
  };
  const KINDS = ['Изричен отказ', 'Частичен отказ', 'Мълчалив отказ'];
  const DATES = ['reg', 'dec', 'recv', 'appeal'];
  const STATUSES = ['Нова', 'В преглед', 'Одобрена', 'Договор изпратен', 'Заведено дело', 'Отказана – изпратени насоки', 'Приключена'];
  const SUPER = 'Суперадминистратор', ADMIN = 'Администратор';
  const ROLES = { 'Суперадминистратор': 4, 'Администратор': 3, 'Юрист': 2, 'Наблюдател': 1 };
  const TEAM_HEAD = ['Имейл', 'Име', 'Роля', 'Активен', 'Добавен', 'Добавен от'];
  const COMM_HEAD = ['Време', '№', 'Автор', 'Текст'];
  const LOG_HEAD = ['Време', 'Потребител', 'Действие', '№', 'Подробности'];
  const MSG_HEAD = ['ID', 'Време', 'От', 'До', '№', 'Текст', 'Прочетено от'];
  const ADMINS = 'admins';
  /* минимална роля за всяко действие; по-фините проверки (заключване, отговорник, роли) са в самите действия */
  const NEED = {
    me: 1, extend: 1, logout: 1, list: 1, get: 1, msgList: 1, msgSend: 1, msgRead: 1,
    update: 2, comment: 2, create: 2, lock: 3, team: 3, teamSave: 3, log: 3, remove: 4, teamRemove: 4
  };

  const ss = () => SpreadsheetApp.getActiveSpreadsheet();
  const lvl = m => ROLES[m.role] || 0;

  function setup() {
    const team = sheet(SH.TEAM, TEAM_HEAD);
    sheet(SH.COMM, COMM_HEAD);
    sheet(SH.MSG, MSG_HEAD);
    const log = sheet(SH.LOG, LOG_HEAD);
    column(ss().getSheetByName(SH.REQ), F.locked);
    column(ss().getSheetByName(SH.REQ), F.src);
    team.getRange(2, 3, team.getMaxRows() - 1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(Object.keys(ROLES), true).setAllowInvalid(false).build());
    team.getRange(2, 4, team.getMaxRows() - 1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(['Да', 'Не'], true).setAllowInvalid(false).build());
    const me = String(Session.getEffectiveUser().getEmail()).toLowerCase();
    const list = members(), cur = list.find(m => m.email === me);
    if (!list.some(m => m.active && m.role === SUPER) && me) {
      if (cur) team.getRange(cur._row, 3, 1, 2).setValues([[SUPER, 'Да']]);
      else team.appendRow([me, '', SUPER, 'Да', new Date(), 'setup']);
    }
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
      if (lvl(me) < need) return { ok: false, error: 'role' };
      const d = req.data || {};
      const s = session(req.token), now = Date.now();
      if (action === 'me') {
        if (s.exp && s.exp < now) return { ok: false, error: 'session' };   // изтекла сесия не се подновява със същия токен
        if (!s.exp) {
          s.exp = now + CFG.SESSION_MIN * 60e3; s.save();
          log_(me.email, d.renew ? 'Удължена сесия (нов вход с Google)' : 'Вход', '', '');
        }
        return { ok: true, me: me, statuses: STATUSES, roles: Object.keys(ROLES), kinds: KINDS, session: s.exp, now: now };
      }
      if (!s.exp || s.exp < now) return { ok: false, error: 'session' };
      switch (action) {
        case 'extend':
          s.exp = now + CFG.SESSION_MIN * 60e3; s.save();
          log_(me.email, 'Удължена сесия', '', '');
          return { ok: true, session: s.exp, now: now };
        case 'logout':
          s.exp = 1; s.save();
          log_(me.email, d.auto ? 'Автоматичен изход (изтекла сесия)' : 'Изход', '', '');
          return { ok: true };
        case 'list':
          return { ok: true, items: items().map(brief), team: teamPublic(), comments: commentCounts(), unread: messages(me).filter(m => !m.read).length };
        case 'create': return create(me, d);
        case 'msgList': return { ok: true, messages: messages(me) };
        case 'msgSend': return msgSend(me, d);
        case 'msgRead': return msgRead(me, d);
        case 'get': return get(String(d.id || ''));
        case 'update': return update(me, d);
        case 'comment': return comment(me, d);
        case 'lock': return lock(me, d);
        case 'remove': return remove(me, d);
        case 'team': return { ok: true, team: members() };
        case 'teamSave': return teamSave(me, d);
        case 'teamRemove': return teamRemove(me, d);
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
    const key = 't' + tokenKey(token);
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

  const tokenKey = token => Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, token));

  /* сесия към токена: exp = край в ms (1 = прекратена); пази се 6 ч. – повече от живота на токена */
  function session(token) {
    const c = CacheService.getScriptCache(), k = 's' + tokenKey(token);
    const o = { exp: Number(c.get(k)) || 0 };
    o.save = () => c.put(k, String(o.exp), 21600);
    return o;
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
        o[k] = k === 'at' ? date(x, true) : DATES.indexOf(k) >= 0 ? date(x) : String(x == null ? '' : x);
      });
      o.locked = o.locked === 'Да';
      return o;
    }).filter(o => o.id);
  }

  function brief(o) {
    const b = {};
    ['id', 'at', 'email', 'name', 'city', 'body', 'kind', 'appeal', 'status', 'owner', 'locked'].forEach(k => b[k] = o[k]);
    return b;
  }

  function find(id) {
    const it = items().find(o => o.id === String(id || ''));
    if (!it) throw new Error('notfound');
    return it;
  }

  /* заключена преписка – само за администратори */
  const open = (me, it) => !it.locked || lvl(me) >= 3;

  function get(id) {
    const it = find(id);
    delete it._row;
    return {
      ok: true, item: it,
      comments: rows(SH.COMM, COMM_HEAD).filter(r => r[1] === id).map(r => ({ at: date(r[0], true), author: r[2], text: r[3] })),
      history: rows(SH.LOG, LOG_HEAD).filter(r => r[3] === id).map(r => ({ at: date(r[0], true), user: r[1], action: r[2], details: r[4] })).reverse()
    };
  }

  function update(me, d) {
    return locked_(() => {
      const sh = ss().getSheetByName(SH.REQ);
      const ix = index(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]);
      const it = find(d.id);
      if (!open(me, it)) return { ok: false, error: 'locked' };
      const ch = {};
      if (d.status != null && d.status !== it.status) {
        if (STATUSES.indexOf(d.status) < 0) return { ok: false, error: 'status' };
        ch.status = d.status;
      }
      if (d.owner != null && d.owner !== it.owner) {
        if (lvl(me) < 3) return { ok: false, error: 'role' };          // насочва само администратор
        if (d.owner && !teamPublic().some(m => m.email === d.owner && ROLES[m.role] >= 2)) return { ok: false, error: 'owner' };
        ch.owner = d.owner;
      }
      if (d.note != null && txt(d.note, 2000) !== it.note) ch.note = txt(d.note, 2000);
      // корекция на данните от подадената заявка
      const fix = d.fields && typeof d.fields === 'object' ? d.fields : {};
      for (const k of Object.keys(fix)) {
        if (!DATA[k]) return { ok: false, error: 'fields' };
        const v = DATA[k][1] === 'kind' ? String(fix[k]) : txt(fix[k], DATA[k][0]);
        if (DATA[k][1] === 'kind' && KINDS.indexOf(v) < 0) return { ok: false, error: 'fields' };
        if (DATA[k][1] === 'date' && v && !/^\d{4}-\d{2}-\d{2}$/.test(v)) return { ok: false, error: 'date' };
        if ((k === 'name' || k === 'body') && !v) return { ok: false, error: 'fields' };
        if (v !== it[k]) ch[k] = v;
      }
      if (!Object.keys(ch).length) return { ok: true, changed: false };
      Object.keys(ch).forEach(k => {
        if (ix[k] == null) throw new Error('column');
        const cell = sh.getRange(it._row, ix[k] + 1);
        if (DATES.indexOf(k) >= 0) cell.setNumberFormat('@');           // датата да остане като текст
        cell.setValue(safe(ch[k]));
      });
      Object.keys(ch).forEach(k => log_(me.email, (DATA[k] ? 'Корекция: ' : 'Промяна: ') + F[k], it.id,
        k === 'note' || k === 'desc' ? ch[k] : (it[k] || '—') + ' → ' + (ch[k] || '—')));

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
          body: me.email + ' ви възложи заявка за съдействие ' + it.id + ' (' + now.kind + ', ' + now.body + ').\n' +
            (now.appeal ? 'Срок за обжалване: ' + now.appeal + '\n' : '') +
            '\n' + CFG.APP_URL + '#/r/' + encodeURIComponent(it.id) + '\n'
        });
      }
      return { ok: true, changed: true, mailed: mailed };
    });
  }

  function lock(me, d) {
    return locked_(() => {
      const sh = ss().getSheetByName(SH.REQ);
      const col = column(sh, F.locked);
      const it = find(d.id), on = !!d.locked;
      if (it.locked === on) return { ok: true, changed: false };
      sh.getRange(it._row, col).setValue(on ? 'Да' : '');
      log_(me.email, on ? 'Заключена' : 'Отключена', it.id, '');
      return { ok: true, changed: true };
    });
  }

  function comment(me, d) {
    const text = txt(d.text, 4000);
    if (!text) return { ok: false, error: 'fields' };
    const it = find(d.id);
    if (!open(me, it)) return { ok: false, error: 'locked' };
    ss().getSheetByName(SH.COMM).appendRow([new Date(), it.id, me.email, safe(text)]);
    log_(me.email, 'Коментар', it.id, '');
    return { ok: true };
  }

  /* изтриване на преписка (суперадминистратор): редът, коментарите; папката с документи – в кошчето на Drive */
  function remove(me, d) {
    return locked_(() => {
      const it = find(d.id);
      if (String(d.confirm || '') !== it.id) return { ok: false, error: 'confirm' };
      let trashed = false;
      const m = /\/folders\/([\w-]+)/.exec(it.files);
      if (m) { try { DriveApp.getFolderById(m[1]).setTrashed(true); trashed = true; } catch (e) { /* папката липсва */ } }
      const cs = ss().getSheetByName(SH.COMM);
      const cv = rows(SH.COMM, COMM_HEAD);
      for (let i = cv.length - 1; i >= 0; i--) if (cv[i][1] === it.id) cs.deleteRow(i + 2);
      ss().getSheetByName(SH.REQ).deleteRow(it._row);
      // броячът на номерата да не слезе под изтрития № – номерът не се дава повторно
      const n = /^ЗС-(\d{4})-(\d+)$/.exec(it.id);
      if (n) {
        const props = PropertiesService.getScriptProperties(), ck = 'SEQ_' + n[1];
        if ((Number(props.getProperty(ck)) || 0) < +n[2]) props.setProperty(ck, n[2].replace(/^0+/, '') || '0');
      }
      log_(me.email, 'Изтрита преписка', it.id, it.name + ' · ' + it.email + ' · ' + it.body + (trashed ? ' · документите са в кошчето на Drive' : ''));
      return { ok: true };
    });
  }

  function commentCounts() {
    const c = {};
    rows(SH.COMM, COMM_HEAD).forEach(r => c[r[1]] = (c[r[1]] || 0) + 1);
    return c;
  }

  /* ръчно въведена заявка (получена по имейл, телефон, на място) */
  function create(me, d) {
    const v = {}, f = d.fields || {};
    for (const k of Object.keys(DATA)) {
      const x = DATA[k][1] === 'kind' ? String(f[k] || '') : txt(f[k], DATA[k][0]);
      if (DATA[k][1] === 'date' && x && !/^\d{4}-\d{2}-\d{2}$/.test(x)) return { ok: false, error: 'date' };
      v[k] = x;
    }
    if (!v.name || !v.body || KINDS.indexOf(v.kind) < 0) return { ok: false, error: 'fields' };
    const email = String(d.email || '').trim().toLowerCase();
    if (email && (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || email.length > 200)) return { ok: false, error: 'email' };
    return locked_(() => {
      const sh = ss().getSheetByName(SH.REQ);
      column(sh, F.locked); column(sh, F.src);
      const all = sh.getDataRange().getValues(), head = all[0], ix = index(head);
      const id = nextId_(all.slice(1));                                    // общият брояч от Code.gs
      let url = '';
      try { url = rootFolder_().createFolder(id + ' – ' + v.name.slice(0, 60)).getUrl(); } catch (e) { /* без папка */ }
      const o = Object.assign({}, v, {
        id: id, at: new Date(), email: email, files: url, status: STATUSES[0], note: '', owner: '', locked: '',
        src: 'Ръчно въведена от ' + me.email
      });
      const row = head.map(() => '');
      Object.keys(o).forEach(k => { if (ix[k] != null) row[ix[k]] = safe(o[k]); });
      sh.appendRow(row);
      // датите да останат като текст ГГГГ-ММ-ДД
      DATES.forEach(k => { if (ix[k] != null) sh.getRange(sh.getLastRow(), ix[k] + 1).setNumberFormat('@').setValue(v[k]); });
      log_(me.email, 'Ръчно създадена', id, v.kind + ' · ' + v.body + ' · ' + v.name);
      if (d.notify && email) {
        MailApp.sendEmail({
          to: email, replyTo: CFG.NOTIFY, name: 'ЗДОИ4You',
          subject: 'Регистрирахме вашата заявка за съдействие ' + id,
          body: 'Здравейте,\n\nРегистрирахме заявката ви за съдействие с № ' + id + ' (' + v.kind + ', ' + v.body + ').\n\n' +
            'Екипът ни ще прецени правната основателност на случая и ще ви пише на този имейл.\n\n' +
            'Важно: заявката не спира сроковете за обжалване. Ако срокът изтича скоро, обжалвайте сами, за да не го пропуснете.\n\n' +
            'Сдружение „Просвета за демокрация“\n' + CFG.NOTIFY
        });
        log_(me.email, 'Имейл до заявителя', id, 'потвърждение');
      }
      return { ok: true, id: id };
    });
  }

  /* ---------- съобщения ---------- */

  /* видими: мои, до мен и – за администраторите – до „администраторите“ */
  function messages(me) {
    return rows(SH.MSG, MSG_HEAD)
      .filter(r => r[2] === me.email || r[3] === me.email || (r[3] === ADMINS && lvl(me) >= 3))
      .slice(-300).reverse()
      .map(r => ({
        id: r[0], at: date(r[1], true), from: r[2], to: r[3], case: r[4], text: r[5],
        read: r[2] === me.email || String(r[6]).split(',').indexOf(me.email) >= 0
      }));
  }

  function msgSend(me, d) {
    const text = txt(d.text, 4000);
    if (!text) return { ok: false, error: 'fields' };
    const team = teamPublic();
    let to = ADMINS;
    if (lvl(me) >= 3 && d.to && d.to !== ADMINS) {                        // администраторът отговаря на конкретен член
      if (!team.some(m => m.email === d.to)) return { ok: false, error: 'owner' };
      to = d.to;
    }
    const cid = String(d.case || '');
    if (cid && !items().some(o => o.id === cid)) return { ok: false, error: 'notfound' };
    sheet(SH.MSG, MSG_HEAD).appendRow([Utilities.getUuid().slice(0, 8), new Date(), me.email, to, cid, safe(text), '']);
    log_(me.email, 'Съобщение', cid, to === ADMINS ? 'до администраторите' : 'до ' + to);
    const rcpt = to === ADMINS ? team.filter(m => ROLES[m.role] >= 3 && m.email !== me.email).map(m => m.email) : [to];
    if (rcpt.length) {
      MailApp.sendEmail({
        to: rcpt.join(','), replyTo: me.email, name: 'ЗДОИ4You – Дела',
        subject: 'Ново съобщение в Дела от ' + me.email + (cid ? ' (' + cid + ')' : ''),
        body: text + '\n\n—\nОтговорете в платформата: ' + CFG.APP_URL + '#/messages'
      });
    }
    return { ok: true };
  }

  function msgRead(me, d) {
    const ids = Array.isArray(d.ids) ? d.ids.map(String) : [];
    if (!ids.length) return { ok: true };
    const sh = ss().getSheetByName(SH.MSG);
    rows(SH.MSG, MSG_HEAD).forEach((r, i) => {
      if (ids.indexOf(String(r[0])) < 0 || r[2] === me.email) return;
      if (!(r[3] === me.email || (r[3] === ADMINS && lvl(me) >= 3))) return;
      const read = String(r[6] || '').split(',').filter(Boolean);
      if (read.indexOf(me.email) < 0) sh.getRange(i + 2, 7).setValue(read.concat(me.email).join(','));
    });
    return { ok: true };
  }

  /* ---------- екип ---------- */

  /* администраторът управлява само юристи и наблюдатели; суперадминистраторът – всички */
  const manages = (me, role) => lvl(me) >= 4 || (ROLES[role] || 0) < 3;

  function teamSave(me, d) {
    const email = String(d.email || '').trim().toLowerCase();
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || email.length > 200) return { ok: false, error: 'email' };
    if (!ROLES[d.role]) return { ok: false, error: 'role' };
    const active = d.active !== false;
    return locked_(() => {
      const list = members();
      const cur = list.find(m => m.email === email);
      if (!manages(me, d.role) || (cur && !manages(me, cur.role))) return { ok: false, error: 'role' };
      // никой не си сваля сам правата и винаги остава поне един активен суперадминистратор
      if (email === me.email && (d.role !== me.role || !active)) return { ok: false, error: 'self' };
      if (cur && cur.role === SUPER && (d.role !== SUPER || !active) &&
        !list.some(m => m.active && m.role === SUPER && m.email !== email)) return { ok: false, error: 'lastadmin' };
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
    });
  }

  function teamRemove(me, d) {
    const email = String(d.email || '').trim().toLowerCase();
    if (email === me.email) return { ok: false, error: 'self' };
    return locked_(() => {
      const list = members();
      const cur = list.find(m => m.email === email);
      if (!cur) return { ok: false, error: 'notfound' };
      if (cur.role === SUPER && !list.some(m => m.active && m.role === SUPER && m.email !== email)) return { ok: false, error: 'lastadmin' };
      ss().getSheetByName(SH.TEAM).deleteRow(cur._row);
      log_(me.email, 'Екип: изтрит', '', email + ' – ' + cur.role);
      return { ok: true, team: members() };
    });
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

  function locked_(fn) {
    const l = LockService.getScriptLock();
    l.waitLock(20000);
    try { return fn(); } finally { l.releaseLock(); }
  }

  function sheet(name, head) {
    const sh = ss().getSheetByName(name) || ss().insertSheet(name);
    if (sh.getLastRow() < 1) sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
    return sh;
  }

  /* номер на колоната с това заглавие; ако липсва – добавя се в края */
  function column(sh, title) {
    const last = Math.max(sh.getLastColumn(), 1);
    const head = sh.getRange(1, 1, 1, last).getValues()[0];
    const i = head.indexOf(title);
    if (i >= 0) return i + 1;
    sh.getRange(1, last + 1).setValue(title).setFontWeight('bold');
    return last + 1;
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
    if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, CFG.TZ, time ? "yyyy-MM-dd'T'HH:mm:ss" : 'yyyy-MM-dd');
    return String(v == null ? '' : v);
  }

  const txt = (v, n) => String(v == null ? '' : v).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '').trim().slice(0, n);
  /* да не се изпълняват като формули в таблицата */
  const safe = v => typeof v === 'string' && /^[=+\-@]/.test(v) ? "'" + v : v;

  return { handle: handle, setup: setup };
})();

/* Изпълнете от редактора: създава листовете и колоните на Дела; при нужда ви прави суперадминистратор. */
function setupDela() { Dela.setup(); }
