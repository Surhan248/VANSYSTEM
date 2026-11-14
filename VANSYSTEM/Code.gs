/**
 * บันทึกรายได้รถตู้ — Backend (Google Apps Script + Google Sheets)
 * 1) แก้ ADMIN_USER / ADMIN_PASS ด้านล่าง  2) รัน setup() หนึ่งครั้ง  3) Deploy เป็น Web app
 */
var ADMIN_USER = 'admin';
var ADMIN_PASS = 'ตั้งรหัสผ่านที่นี่ก่อนรัน setup';

var COLS = {
  Users: ['username', 'salt', 'hash', 'role', 'name', 'plate', 'pct', 'reserve', 'wStart', 'wEnd', 'active'],
  Entries: ['id', 'username', 'date', 'type', 'src', 'cat', 'note', 'amount', 'ts', 'weekId'],
  Weeks: ['id', 'username', 'start', 'end', 'i', 'o', 'net', 'res', 'pct', 'd', 'own', 'ib', 'bk', 'items', 'closedAt'],
  Withdrawals: ['id', 'username', 'date', 'amount', 'note', 'ts']
};
var TEXTK = { id: 1, username: 1, salt: 1, hash: 1, role: 1, name: 1, plate: 1, date: 1, start: 1, end: 1, wStart: 1, wEnd: 1, weekId: 1, src: 1, cat: 1, type: 1, note: 1, active: 1, ib: 1, bk: 1, items: 1 };
var NUMK = { amount: 1, ts: 1, i: 1, o: 1, net: 1, res: 1, pct: 1, d: 1, own: 1, reserve: 1, closedAt: 1 };
var SRCK = { grab: 1, bolt: 1, groundline: 1, agent: 1, other: 1 };
var CATK = { fuel: 1, repair: 1, park: 1, other: 1 };

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('บันทึกรายได้รถตู้')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ---------- ติดตั้งครั้งแรก ---------- */
function setup() {
  var ss = SpreadsheetApp.getActive();
  Object.keys(COLS).forEach(function (n) {
    var s = ss.getSheetByName(n) || ss.insertSheet(n);
    s.getRange(1, 1, 1, COLS[n].length).setValues([COLS[n]]);
    s.setFrozenRows(1);
    COLS[n].forEach(function (k, j) {
      if (TEXTK[k]) s.getRange(2, j + 1, s.getMaxRows() - 1, 1).setNumberFormat('@');
    });
  });
  if (!findUser(ADMIN_USER)) {
    var t = today();
    createUserRow({ username: ADMIN_USER, password: ADMIN_PASS, role: 'admin', name: 'ผู้ดูแล', plate: '', pct: 50, reserve: 0, wStart: t, wEnd: addDays(t, 6) });
  }
  Logger.log('setup เสร็จแล้ว เข้าระบบด้วย ' + ADMIN_USER);
}

/* ---------- ตัวช่วยข้อมูล ---------- */
function sh(n) { return SpreadsheetApp.getActive().getSheetByName(n); }
function today() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function addDays(s, n) { var d = new Date(s + 'T00:00:00'); d.setDate(d.getDate() + n); return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function norm(o) {
  Object.keys(o).forEach(function (k) {
    var v = o[k];
    if (v instanceof Date) o[k] = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    else if (NUMK[k]) o[k] = Number(v) || 0;
    else if (v === '' || v === null) o[k] = '';
    else if (typeof v !== 'string') o[k] = String(v);
  });
  return o;
}
function readAll(n) {
  var v = sh(n).getDataRange().getValues(), h = v[0], out = [];
  for (var r = 1; r < v.length; r++) {
    if (v[r][0] === '') continue;
    var o = { _row: r + 1 };
    h.forEach(function (k, j) { o[k] = v[r][j]; });
    out.push(norm(o));
  }
  return out;
}
function appendRow(n, o) { sh(n).appendRow(COLS[n].map(function (k) { return o[k] === undefined ? '' : o[k]; })); }
function updateRow(n, row, o) {
  var c = COLS[n], rg = sh(n).getRange(row, 1, 1, c.length), cur = rg.getValues()[0];
  c.forEach(function (k, j) { if (o[k] !== undefined) cur[j] = o[k]; });
  rg.setValues([cur]);
}
function setCol(n, key, rowVals) {
  var s = sh(n), c = COLS[n].indexOf(key) + 1, cnt = s.getLastRow() - 1;
  if (cnt < 1) return;
  var rg = s.getRange(2, c, cnt, 1), v = rg.getValues();
  Object.keys(rowVals).forEach(function (r) { v[r - 2][0] = rowVals[r]; });
  rg.setValues(v);
}

/* ---------- ผู้ใช้และรหัสผ่าน ---------- */
function hex(b) { return b.map(function (x) { return ('0' + (x & 0xff).toString(16)).slice(-2); }).join(''); }
function hashPw(pw, salt) {
  var h = salt + pw;
  for (var i = 0; i < 1000; i++) h = hex(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + salt));
  return h;
}
function findUser(un) { return readAll('Users').filter(function (u) { return u.username === un; })[0]; }
function getUser(un) { var u = findUser(un); if (!u) throw new Error('ไม่พบผู้ใช้ ' + un); return u; }
function pub(u) { return { username: u.username, role: u.role, name: u.name, plate: u.plate, pct: u.pct, reserve: u.reserve, wStart: u.wStart, wEnd: u.wEnd, active: u.active }; }
function createUserRow(p) {
  var un = String(p.username || '').trim();
  if (!/^[A-Za-z0-9_.-]{3,30}$/.test(un)) throw new Error('username ใช้ได้เฉพาะ a-z 0-9 _ . - ยาว 3-30 ตัว');
  if (String(p.password || '').length < 6) throw new Error('รหัสผ่านต้องยาวอย่างน้อย 6 ตัว');
  if (findUser(un)) throw new Error('username นี้มีอยู่แล้ว');
  var salt = Utilities.getUuid(), t = today();
  appendRow('Users', {
    username: un, salt: salt, hash: hashPw(String(p.password), salt), role: p.role === 'admin' ? 'admin' : 'user',
    name: String(p.name || un).slice(0, 60), plate: String(p.plate || '').slice(0, 20),
    pct: p.pct === undefined ? 50 : Number(p.pct), reserve: Number(p.reserve) || 0,
    wStart: p.wStart || t, wEnd: p.wEnd || addDays(t, 6), active: 'Y'
  });
}

/* ---------- token (ลงลายเซ็น ไม่ต้องเก็บ session) ---------- */
function secret() {
  var p = PropertiesService.getScriptProperties(), s = p.getProperty('SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('SECRET', s); }
  return s;
}
function sign(s) { return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(s, secret())); }
function makeToken(un) {
  var body = Utilities.base64EncodeWebSafe(un + '|' + (Date.now() + 12 * 3600 * 1000));
  return body + '.' + sign(body);
}
function auth(t) {
  var parts = String(t || '').split('.');
  if (parts.length !== 2 || sign(parts[0]) !== parts[1]) throw new Error('ยังไม่ได้เข้าสู่ระบบ');
  var s = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString().split('|');
  if (Number(s[1]) < Date.now()) throw new Error('เซสชันหมดอายุ');
  var u = findUser(s[0]);
  if (!u || u.active === 'N') throw new Error('บัญชีนี้ถูกปิดใช้งาน');
  return u;
}
function login(p) {
  var un = String(p.username || '').trim(), c = CacheService.getScriptCache(), k = 'f_' + un;
  if (Number(c.get(k) || 0) >= 5) throw new Error('ใส่รหัสผิดหลายครั้ง รอ 10 นาทีแล้วลองใหม่');
  var u = findUser(un);
  if (!u || u.active === 'N' || hashPw(String(p.password || ''), u.salt) !== u.hash) {
    c.put(k, String(Number(c.get(k) || 0) + 1), 600);
    throw new Error('username หรือรหัสผ่านไม่ถูกต้อง');
  }
  c.remove(k);
  return { token: makeToken(un), me: pub(u) };
}

/* ---------- API เดียวสำหรับหน้าเว็บ ---------- */
function api(action, token, p) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return JSON.stringify(route(action, token, p || {})); }
  catch (e) { return JSON.stringify({ error: String(e.message || e) }); }
  finally { lock.releaseLock(); }
}

function checkEntry(p) {
  var a = Number(p.amount);
  if (!(a > 0)) throw new Error('จำนวนเงินต้องมากกว่า 0');
  if (!/^\d{4}-\d\d-\d\d$/.test(String(p.date))) throw new Error('วันที่ไม่ถูกต้อง');
  if (p.type !== 'in' && p.type !== 'out') throw new Error('ประเภทไม่ถูกต้อง');
  return {
    date: p.date, type: p.type, amount: a, note: String(p.note || '').slice(0, 200),
    src: p.type === 'in' ? (SRCK[p.src] ? p.src : 'other') : '',
    cat: p.type === 'out' ? (CATK[p.cat] ? p.cat : 'other') : ''
  };
}
function calc(es, pct, reserve) {
  var i = 0, o = 0, ib = {}, bk = {};
  es.forEach(function (e) {
    if (e.type === 'in') { i += e.amount; ib[e.src || 'other'] = (ib[e.src || 'other'] || 0) + e.amount; }
    else { o += e.amount; bk[e.cat || 'other'] = (bk[e.cat || 'other'] || 0) + e.amount; }
  });
  var base = Math.max(0, i - o), res = Math.min(Math.max(Number(reserve) || 0, 0), base);
  base -= res;
  var pc = Math.min(100, Math.max(0, Number(pct) || 0)), d = Math.round(base * pc) / 100;
  return { i: i, o: o, net: i - o, res: res, pct: pc, d: d, own: base - d, ib: ib, bk: bk };
}
function weekOut(w) {
  ['ib', 'bk', 'items'].forEach(function (k) { try { w[k] = JSON.parse(w[k] || (k === 'items' ? '[]' : '{}')); } catch (e) { w[k] = k === 'items' ? [] : {}; } });
  return w;
}

function route(a, t, p) {
  if (a === 'login') return login(p);
  var me = auth(t), adm = me.role === 'admin';
  if (a !== 'me' && a !== 'addEntry' && a !== 'listEntries' && !adm) throw new Error('ต้องใช้สิทธิ์ Admin');
  switch (a) {
    case 'me': return pub(me);
    case 'addEntry': {
      var un = adm && p.username ? p.username : me.username; getUser(un);
      var e = checkEntry(p); e.id = 'e' + Date.now() + Math.floor(Math.random() * 1000);
      e.username = un; e.ts = Date.now(); e.weekId = '';
      appendRow('Entries', e); return { ok: 1 };
    }
    case 'listEntries': {
      var u2 = adm && p.username ? p.username : me.username;
      return readAll('Entries').filter(function (x) { return x.username === u2 && !x.weekId; });
    }
    case 'updateEntry': {
      var cur = readAll('Entries').filter(function (x) { return x.id === p.id && !x.weekId; })[0];
      if (!cur) throw new Error('ไม่พบรายการ หรือถูกปิดสัปดาห์ไปแล้ว');
      updateRow('Entries', cur._row, checkEntry(p)); return { ok: 1 };
    }
    case 'deleteEntry': {
      var d = readAll('Entries').filter(function (x) { return x.id === p.id && !x.weekId; })[0];
      if (!d) throw new Error('ไม่พบรายการ หรือถูกปิดสัปดาห์ไปแล้ว');
      sh('Entries').deleteRow(d._row); return { ok: 1 };
    }
    case 'closeWeek': {
      var u = getUser(p.username);
      var es = readAll('Entries').filter(function (x) { return x.username === u.username && !x.weekId && x.date >= p.start && x.date <= p.end; });
      if (!es.length) throw new Error('ไม่มีรายการในช่วงวันที่นี้');
      var c = calc(es, u.pct, u.reserve), id = 'w' + Date.now(), rv = {};
      appendRow('Weeks', {
        id: id, username: u.username, start: p.start, end: p.end, i: c.i, o: c.o, net: c.net, res: c.res, pct: c.pct, d: c.d, own: c.own,
        ib: JSON.stringify(c.ib), bk: JSON.stringify(c.bk), closedAt: Date.now(),
        items: JSON.stringify(es.map(function (x) { return { id: x.id, date: x.date, type: x.type, src: x.src, cat: x.cat, note: x.note, amount: x.amount, ts: x.ts }; }))
      });
      es.forEach(function (x) { rv[x._row] = id; });
      setCol('Entries', 'weekId', rv);
      var ns = addDays(p.end, 1);
      updateRow('Users', u._row, { wStart: ns, wEnd: addDays(ns, 6) });
      return { ok: 1, id: id };
    }
    case 'deleteWeek': {
      var w = readAll('Weeks').filter(function (x) { return x.id === p.id; })[0];
      if (!w) throw new Error('ไม่พบรายงาน');
      var rr = {};
      readAll('Entries').forEach(function (x) { if (x.weekId === p.id) rr[x._row] = ''; });
      setCol('Entries', 'weekId', rr);
      sh('Weeks').deleteRow(w._row); return { ok: 1 };
    }
    case 'dashboard':
      return {
        users: readAll('Users').map(pub),
        weeks: readAll('Weeks').map(weekOut),
        withdrawals: readAll('Withdrawals')
      };
    case 'addWithdrawal': {
      getUser(p.username);
      if (!(Number(p.amount) > 0)) throw new Error('ใส่ยอดที่ใช้');
      appendRow('Withdrawals', { id: 'f' + Date.now(), username: p.username, date: p.date || today(), amount: Number(p.amount), note: String(p.note || '').slice(0, 200), ts: Date.now() });
      return { ok: 1 };
    }
    case 'deleteWithdrawal': {
      var f = readAll('Withdrawals').filter(function (x) { return x.id === p.id; })[0];
      if (!f) throw new Error('ไม่พบรายการ');
      sh('Withdrawals').deleteRow(f._row); return { ok: 1 };
    }
    case 'createUser': createUserRow(p); return { ok: 1 };
    case 'updateUser': {
      var tu = getUser(p.username), o = {};
      if (tu.username === me.username && (p.active === 'N' || (p.role && p.role !== 'admin'))) throw new Error('ปิดหรือลดสิทธิ์บัญชีตัวเองไม่ได้');
      if (p.name !== undefined) o.name = String(p.name).slice(0, 60);
      if (p.plate !== undefined) o.plate = String(p.plate).slice(0, 20);
      if (p.pct !== undefined) o.pct = Math.min(100, Math.max(0, Number(p.pct) || 0));
      if (p.reserve !== undefined) o.reserve = Math.max(0, Number(p.reserve) || 0);
      if (p.wStart) o.wStart = p.wStart;
      if (p.wEnd) o.wEnd = p.wEnd;
      if (p.active === 'Y' || p.active === 'N') o.active = p.active;
      if (p.role === 'admin' || p.role === 'user') o.role = p.role;
      if (p.password) {
        if (String(p.password).length < 6) throw new Error('รหัสผ่านต้องยาวอย่างน้อย 6 ตัว');
        var salt = Utilities.getUuid(); o.salt = salt; o.hash = hashPw(String(p.password), salt);
      }
      updateRow('Users', tu._row, o); return { ok: 1 };
    }
  }
  throw new Error('ไม่รู้จักคำสั่ง ' + a);
}
