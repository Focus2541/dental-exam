/**
 * ระบบเชื่อมข้อมูลแอป "กรอกผลตรวจฟันนักเรียน" กับ Google Sheet
 * ติดตั้ง: ดูไฟล์ README.md ในโฟลเดอร์ apps-script
 *
 * Sheet ที่ใช้ (สร้างให้อัตโนมัติ):
 *   รอบตรวจ  : รหัสรอบ | PIN | หน่วยบริการ | โรงเรียน | วันที่ตรวจ | สร้างเมื่อ | สร้างโดย
 *   รายชื่อ   : รหัสรอบ | ห้อง | key | เลขบัตร | คำนำหน้า | ชื่อ | นามสกุล
 *   ผลตรวจ   : รหัสรอบ | key | ห้อง | ชื่อ-สกุล | ฟันน้ำนม(ทั้งหมด ผุ อุด ถอน) | ฟันแท้(…) | เหงือกอักเสบ | หินปูน | Tx อุด | Tx ถอน | Sealant | เวลา(ms) | บันทึกโดย | แก้ไขล่าสุด
 */

var SH = {
  S: ['รอบตรวจ', ['รหัสรอบ', 'PIN', 'หน่วยบริการ', 'โรงเรียน', 'วันที่ตรวจ', 'สร้างเมื่อ', 'สร้างโดย']],
  R: ['รายชื่อ', ['รหัสรอบ', 'ห้อง', 'key', 'เลขบัตร', 'คำนำหน้า', 'ชื่อ', 'นามสกุล']],
  X: ['ผลตรวจ', ['รหัสรอบ', 'key', 'ห้อง', 'ชื่อ-สกุล',
    'น้ำนม-ทั้งหมด', 'น้ำนม-ผุ', 'น้ำนม-อุด', 'น้ำนม-ถอน/หลุด',
    'แท้-ทั้งหมด', 'แท้-ผุ', 'แท้-อุด', 'แท้-ถอน/หลุด',
    'เหงือกอักเสบ', 'หินปูน', 'Tx-อุด', 'Tx-ถอน', 'Tx-Sealant',
    'เวลา(ms)', 'บันทึกโดย', 'แก้ไขล่าสุด']]
};
var NV = 13; // จำนวนช่องผลตรวจ

function doGet() {
  return out_({ ok: true, app: 'dental-exam', msg: 'ระบบเชื่อมข้อมูลทำงานอยู่' });
}

function doPost(e) {
  try {
    var q = JSON.parse(e.postData.contents);
    var fn = { create: create_, join: join_, push: push_, pull: pull_ }[q.action];
    if (!fn) return out_({ ok: false, error: 'ไม่รู้จักคำสั่ง' });
    return out_(fn(q));
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  }
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function sheet_(k) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), d = SH[k], s = ss.getSheetByName(d[0]);
  if (!s) {
    s = ss.insertSheet(d[0]);
    s.getRange(1, 1, 1, d[1].length).setValues([d[1]]).setFontWeight('bold');
    s.setFrozenRows(1);
    if (k === 'R' || k === 'X') s.getRange('B:D').setNumberFormat('@'); // เลขบัตรขึ้นต้น 0 / G ไม่เพี้ยน
  }
  return s;
}

function rows_(s) {
  var n = s.getLastRow();
  return n < 2 ? [] : s.getRange(2, 1, n - 1, s.getLastColumn()).getValues();
}

function withLock_(f) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return f(); } finally { lock.releaseLock(); }
}

function session_(code, pin) {
  code = String(code || '').trim().toUpperCase();
  var r = rows_(sheet_('S')).filter(function (x) { return String(x[0]).toUpperCase() === code; })[0];
  if (!r) throw new Error('ไม่พบรหัสรอบตรวจ ' + code);
  if (String(r[1]) !== String(pin || '').trim()) throw new Error('PIN ไม่ถูกต้อง');
  return { code: code, unit: r[2], school: r[3], date: r[4] };
}

function newCode_(used) {
  var A = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789', c;
  do {
    c = '';
    for (var i = 0; i < 6; i++) c += A.charAt(Math.floor(Math.random() * A.length));
  } while (used[c]);
  return c;
}

/** สร้างรอบตรวจใหม่: { school, date, unit, pin, by, roster:[{room,key,id,pre,fn,ln}] } */
function create_(q) {
  var pin = String(q.pin || '').trim();
  if (!/^\d{4,6}$/.test(pin)) throw new Error('PIN ต้องเป็นตัวเลข 4-6 หลัก');
  if (!q.roster || !q.roster.length) throw new Error('ไม่มีรายชื่อนักเรียน');
  return withLock_(function () {
    var s = sheet_('S'), used = {};
    rows_(s).forEach(function (x) { used[String(x[0]).toUpperCase()] = 1; });
    var code = newCode_(used);
    s.appendRow([code, "'" + pin, q.unit || '', q.school || '', q.date || '', new Date(), q.by || '']);
    var R = sheet_('R'), vals = q.roster.map(function (p) {
      return [code, p.room || '', String(p.key), p.noid ? '' : String(p.id || ''), p.pre || '', p.fn || '', p.ln || ''];
    });
    R.getRange(R.getLastRow() + 1, 1, vals.length, 7).setNumberFormat('@').setValues(vals);
    return { ok: true, code: code };
  });
}

/** เข้าร่วมรอบตรวจ: ได้รายชื่อและผลที่มีอยู่ */
function join_(q) {
  var se = session_(q.code, q.pin);
  var roster = rows_(sheet_('R')).filter(function (x) { return String(x[0]).toUpperCase() === se.code; })
    .map(function (x) { var id = String(x[3]); return { room: String(x[1]), key: String(x[2]), id: id || String(x[2]), noid: !id, pre: x[4], fn: x[5], ln: x[6] }; });
  var p = pull_({ code: q.code, pin: q.pin, since: 0 });
  return { ok: true, session: se, roster: roster, results: p.results, now: p.now };
}

/** ส่งผลขึ้น: items:[{key, room, name, v:[13], t}] — เก็บค่าที่ใหม่กว่า (ตามเวลา t) */
function push_(q) {
  var se = session_(q.code, q.pin), items = q.items || [];
  return withLock_(function () {
    var X = sheet_('X'), data = rows_(X), idx = {}, now = new Date(), saved = 0, appends = [];
    data.forEach(function (r, i) { if (String(r[0]).toUpperCase() === se.code) idx[String(r[1])] = i; });
    items.forEach(function (it) {
      var v = (it.v || []).slice(0, NV); while (v.length < NV) v.push('');
      v = v.map(function (x) { return x === '' || x == null ? '' : Number(x); });
      var t = Number(it.t) || now.getTime();
      var row = [se.code, String(it.key), it.room || '', it.name || ''].concat(v, [t, q.by || '', now]);
      var i = idx[String(it.key)];
      if (i == null) { appends.push(row); idx[String(it.key)] = -1; saved++; }
      else if (i >= 0 && t >= Number(data[i][4 + NV])) { X.getRange(i + 2, 1, 1, row.length).setValues([row]); saved++; }
    });
    if (appends.length) X.getRange(X.getLastRow() + 1, 1, appends.length, appends[0].length).setValues(appends);
    return { ok: true, saved: saved, now: now.getTime() };
  });
}

/** ดึงผลที่แก้ไขหลังเวลา since (ms ตามเวลาเซิร์ฟเวอร์) */
function pull_(q) {
  var se = session_(q.code, q.pin), since = Number(q.since) || 0, now = Date.now();
  var res = rows_(sheet_('X')).filter(function (r) {
    if (String(r[0]).toUpperCase() !== se.code) return false;
    var m = r[6 + NV] instanceof Date ? r[6 + NV].getTime() : Number(new Date(r[6 + NV]));
    return !since || m >= since - 5000;
  }).map(function (r) {
    return { key: String(r[1]), v: r.slice(4, 4 + NV).map(function (x) { return x === '' || x == null ? '' : String(x); }), t: Number(r[4 + NV]) };
  });
  return { ok: true, results: res, now: now };
}
