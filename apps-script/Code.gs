/**
 * ระบบเช็คชื่อผู้เข้าอบรม — Google Apps Script (ผูกกับ Google Sheet)
 *
 * ชีตที่ใช้ (สร้างให้อัตโนมัติ):
 *   รายชื่อ            : รหัส | คำนำหน้า | ชื่อ-นามสกุล | เพศ | หน่วยงาน | ตำแหน่ง | วันที่เพิ่ม | ที่มา
 *   บันทึกเช็คชื่อ      : วันที่ | ช่วง (เช้า/บ่าย) | เวลา | รหัส | คำนำหน้า | ชื่อ-นามสกุล | เพศ | หน่วยงาน | ตำแหน่ง | สถานะ (มา/ขาด) | บันทึกเมื่อ
 *   สรุป               : จำนวน มา / ขาด แยกตามเพศ ของวันที่ (B2) และช่วง (B3) — สูตรคำนวณเอง
 *   ตั้งค่า             : ชื่อการอบรม / สถานที่ / ชื่อชีตคำตอบแบบฟอร์ม
 *
 * ใช้งานได้ 2 แบบ:
 *   ก) เปิดหน้าเว็บจาก Apps Script โดยตรง (ไฟล์ Index.html อยู่ในโปรเจกต์นี้)
 *   ข) วางหน้าเว็บบน GitHub Pages แล้วเรียก API นี้ (ใส่ลิงก์ /exec ใน config.js)
 *      GET  ?action=getData&date=YYYY-MM-DD
 *      POST {action: checkIn|undoCheckIn|markAbsent|addPeople, args: [...], key}
 *      ถ้าตั้ง Script Property ชื่อ API_KEY ไว้ ทุกคำขอต้องส่ง key ให้ตรง
 *
 * รายชื่อเข้ามาได้ 3 ทาง:
 *   1) อัตโนมัติจากชีตคำตอบ Google Form ("การตอบแบบฟอร์ม 1") — แตกผู้สมัครคนที่ 2–5 เป็นรายคนให้
 *   2) ปุ่ม "เพิ่มรายชื่อ" บนเว็บ
 *   3) พิมพ์ต่อท้ายในชีต "รายชื่อ" (เว้นรหัส/เพศว่างได้ ระบบเติมให้)
 */

const TZ = 'Asia/Bangkok';
const SESSIONS = ['เช้า', 'บ่าย'];
const SH = { people: 'รายชื่อ', log: 'บันทึกเช็คชื่อ', sum: 'สรุป', cfg: 'ตั้งค่า' };
const DEFAULT_FORM_SHEET = 'การตอบแบบฟอร์ม 1';
const HEAD = {
  people: ['รหัส', 'คำนำหน้า', 'ชื่อ-นามสกุล', 'เพศ', 'หน่วยงาน', 'ตำแหน่ง', 'วันที่เพิ่ม', 'ที่มา'],
  log: ['วันที่', 'ช่วง', 'เวลา', 'รหัส', 'คำนำหน้า', 'ชื่อ-นามสกุล', 'เพศ', 'หน่วยงาน', 'ตำแหน่ง', 'สถานะ', 'บันทึกเมื่อ'],
  cfg: ['หัวข้อ', 'ค่า']
};
// ตำแหน่งคอลัมน์ (เริ่มที่ 0)
const P = { id: 0, pre: 1, name: 2, sex: 3, dept: 4, pos: 5 };
const L = { date: 0, ses: 1, time: 2, id: 3, pre: 4, name: 5, sex: 6, dept: 7, pos: 8, status: 9 };

/* ---------- หน้าเว็บ / API / เมนู ---------- */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action) {
    if (p.action !== 'getData') return json_({ ok: false, error: 'คำสั่ง ' + p.action + ' ต้องส่งแบบ POST' });
    return api_('getData', [p.date], p.key);
  }
  try {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('เช็คชื่อผู้เข้าอบรม')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return json_({ ok: true, data: 'ระบบเช็คชื่อพร้อมใช้งาน — นำลิงก์นี้ไปใส่เป็น API_URL ใน config.js' });
  }
}

function doPost(e) {
  let body;
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return json_({ ok: false, error: 'ข้อมูลที่ส่งมาไม่ใช่ JSON' }); }
  return api_(body.action, Array.isArray(body.args) ? body.args : [], body.key);
}

function api_(action, args, key) {
  try {
    const need = PropertiesService.getScriptProperties().getProperty('API_KEY');
    if (need && String(key || '') !== need) throw new Error('API key ไม่ถูกต้อง');
    const fns = { getData: getData, checkIn: checkIn, undoCheckIn: undoCheckIn, markAbsent: markAbsent, addPeople: addPeople };
    if (!Object.prototype.hasOwnProperty.call(fns, action)) throw new Error('ไม่รู้จักคำสั่ง ' + action);
    return json_({ ok: true, data: fns[action].apply(null, args) });
  } catch (err) {
    return json_({ ok: false, error: (err && err.message) || String(err) });
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('ระบบเช็คชื่อ')
    .addItem('สร้างชีตเริ่มต้น', 'setup')
    .addItem('ดึงรายชื่อจากแบบฟอร์มทั้งหมดอีกครั้ง', 'reimportForm')
    .addToUi();
}

/* ---------- ตั้งค่าครั้งแรก ---------- */
function setup() {
  sheet_(SH.people, HEAD.people, sh => {
    sh.getRange('A:A').setNumberFormat('@');
    sh.setColumnWidth(2, 80).setColumnWidth(3, 220).setColumnWidth(4, 70).setColumnWidth(5, 240).setColumnWidth(6, 180);
    sh.getRange('D2:D').setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['ชาย', 'หญิง', 'ไม่ระบุ'], true).setAllowInvalid(true).build());
  });
  sheet_(SH.log, HEAD.log, sh => {
    sh.getRange('A:D').setNumberFormat('@'); // เก็บวันที่/เวลาเป็นข้อความ กันปัญหาเขตเวลา
    sh.setColumnWidth(6, 220).setColumnWidth(8, 240);
  });
  sheet_(SH.cfg, HEAD.cfg, sh => {
    sh.getRange(2, 1, 3, 2).setValues([
      ['ชื่อการอบรม', 'ชื่อหลักสูตรอบรม (แก้ไขได้ที่ชีต ตั้งค่า)'],
      ['สถานที่', 'ห้องประชุม'],
      ['ชีตแบบฟอร์ม', DEFAULT_FORM_SHEET]
    ]);
    sh.setColumnWidth(1, 160).setColumnWidth(2, 380);
  });
  summarySheet_();
}

/** ชีต "สรุป": เปลี่ยนวันที่ (B2) หรือช่วง (B3) แล้วตัวเลขคำนวณใหม่เอง */
function summarySheet_() {
  const ss = SpreadsheetApp.getActive();
  if (ss.getSheetByName(SH.sum)) return;
  const sh = ss.insertSheet(SH.sum);
  const R = "'" + SH.people + "'!", G = "'" + SH.log + "'!";
  sh.getRange('A1').setValue('สรุปการเข้าอบรมแยกตามเพศ').setFontWeight('bold').setFontSize(14);
  sh.getRange('A2:A3').setValues([['วันที่ (ปปปป-ดด-วว)'], ['ช่วง']]);
  sh.getRange('B2').setNumberFormat('@').setValue(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'));
  sh.getRange('B3').setValue('เช้า').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(SESSIONS, true).build());
  sh.getRange('B2:B3').setBackground('#fff4c2');
  sh.getRange('A5:F5').setValues([['เพศ', 'ทั้งหมด', 'มา', 'ขาด', 'ยังไม่บันทึก', '% มา']])
    .setFontWeight('bold').setBackground('#2f64a3').setFontColor('#ffffff');
  const rows = ['ชาย', 'หญิง', 'ไม่ระบุ'].map((g, i) => {
    const r = 6 + i, f = st => `=COUNTIFS(${G}A:A,$B$2,${G}B:B,$B$3,${G}G:G,A${r},${G}J:J,"${st}")`;
    return [g, `=COUNTIF(${R}D:D,A${r})`, f('มา'), f('ขาด'), `=MAX(0,B${r}-C${r}-D${r})`, `=IF(B${r}=0,"",C${r}/B${r})`];
  });
  sh.getRange(6, 1, 3, 6).setValues(rows);
  sh.getRange('A9:F9').setValues([['รวม', '=SUM(B6:B8)', '=SUM(C6:C8)', '=SUM(D6:D8)', '=SUM(E6:E8)', '=IF(B9=0,"",C9/B9)']])
    .setFontWeight('bold').setBackground('#e7eef6');
  sh.getRange('F6:F9').setNumberFormat('0%');
  sh.setColumnWidth(1, 170);
}

function sheet_(name, head, init) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, head.length).setValues([head])
      .setFontWeight('bold').setBackground('#2f64a3').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    if (init) init(sh);
  }
  return sh;
}

function ensure_() {
  const ss = SpreadsheetApp.getActive();
  if ([SH.people, SH.log, SH.sum, SH.cfg].some(n => !ss.getSheetByName(n))) setup();
}

/* ---------- API ที่หน้าเว็บเรียกใช้ ---------- */

/** รายชื่อ + สถานะของวันที่ที่เลือก ทั้งช่วงเช้าและบ่าย */
function getData(date) {
  date = checkDate_(date);
  return withLock_(() => {
    ensure_();
    const cfg = cfg_();
    const imported = importForm_(cfg);
    const sessions = {};
    SESSIONS.forEach(s => { sessions[s] = { checks: {}, absent: {} }; });
    rows_(SH.log).forEach(r => {
      const id = String(r[L.id]).trim(), s = sessions[String(r[L.ses]).trim()];
      if (day_(r[L.date]) !== date || !id || !s) return;
      if (isAbsent_(r)) s.absent[id] = true; else s.checks[id] = time_(r[L.time]);
    });
    SESSIONS.forEach(k => Object.keys(sessions[k].checks).forEach(id => delete sessions[k].absent[id]));
    return {
      people: people_(),
      sessions: sessions,
      imported: imported,
      config: { title: cfg['ชื่อการอบรม'] || 'การอบรม', place: cfg['สถานที่'] || '' },
      sheetUrl: SpreadsheetApp.getActive().getUrl()
    };
  });
}

/** เช็คชื่อ 1 คน (ถ้าเคยบันทึกว่า "ขาด" ในช่วงนั้น จะเปลี่ยนเป็น "มา") */
function checkIn(id, date, ses) {
  date = checkDate_(date); ses = checkSes_(ses); id = String(id || '').trim();
  return withLock_(() => {
    ensure_();
    const sh = SpreadsheetApp.getActive().getSheetByName(SH.log);
    const v = rows_(SH.log);
    const mine = r => String(r[L.id]).trim() === id && day_(r[L.date]) === date && String(r[L.ses]).trim() === ses;
    const ex = v.find(r => mine(r) && !isAbsent_(r));
    if (ex) return { time: time_(ex[L.time]), already: true };
    const p = people_().find(x => x.id === id);
    if (!p) throw new Error('ไม่พบรายชื่อรหัส ' + id + ' (อาจถูกลบออกจากชีตแล้ว)');
    for (let i = v.length - 1; i >= 0; i--) if (mine(v[i]) && isAbsent_(v[i])) sh.deleteRow(i + 2);
    const now = new Date();
    const t = Utilities.formatDate(now, TZ, 'HH:mm:ss');
    sh.appendRow([date, ses, t, p.id, p.pre, p.name, p.sex, p.dept, p.pos, 'มา', now]);
    return { time: t };
  });
}

/** ยกเลิกการเช็คชื่อ (เอาติ๊กออก) */
function undoCheckIn(id, date, ses) {
  date = checkDate_(date); ses = checkSes_(ses); id = String(id || '').trim();
  return withLock_(() => {
    ensure_();
    const sh = SpreadsheetApp.getActive().getSheetByName(SH.log);
    const v = rows_(SH.log);
    let removed = 0;
    for (let i = v.length - 1; i >= 0; i--) {
      const r = v[i];
      if (String(r[L.id]).trim() === id && day_(r[L.date]) === date && String(r[L.ses]).trim() === ses && !isAbsent_(r)) {
        sh.deleteRow(i + 2); removed++;
      }
    }
    return { removed: removed };
  });
}

/** บันทึกผู้ไม่มาอบรม (สถานะ "ขาด") ของช่วงนั้น — ข้ามคนที่มาแล้วหรือบันทึกขาดไว้แล้ว */
function markAbsent(ids, date, ses) {
  date = checkDate_(date); ses = checkSes_(ses);
  return withLock_(() => {
    ensure_();
    const sh = SpreadsheetApp.getActive().getSheetByName(SH.log);
    const done = new Set();
    rows_(SH.log).forEach(r => {
      if (day_(r[L.date]) === date && String(r[L.ses]).trim() === ses) done.add(String(r[L.id]).trim());
    });
    const byId = {};
    people_().forEach(p => { byId[p.id] = p; });
    const now = new Date(), rows = [], marked = [];
    (ids || []).forEach(id => {
      id = String(id).trim();
      const p = byId[id];
      if (!p || done.has(id)) return;
      done.add(id);
      rows.push([date, ses, '', p.id, p.pre, p.name, p.sex, p.dept, p.pos, 'ขาด', now]);
      marked.push(id);
    });
    if (rows.length) {
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, HEAD.log.length).setValues(rows).setBackground('#fff4c2');
    }
    return { marked: marked };
  });
}

/** เพิ่มรายชื่อจากหน้าเว็บ — list: [{pre, name, sex, dept, pos}] */
function addPeople(list) {
  return withLock_(() => { ensure_(); return add_(list, 'หน้าเว็บ'); });
}

/** เมนู: ดึงรายชื่อจากแบบฟอร์มทุกแถวอีกครั้ง (ชื่อที่มีอยู่แล้วจะถูกข้าม) */
function reimportForm() {
  withLock_(() => {
    ensure_();
    const cfg = cfg_();
    PropertiesService.getDocumentProperties().deleteProperty(formKey_(cfg));
    const n = importForm_(cfg);
    SpreadsheetApp.getActive().toast('เพิ่มรายชื่อใหม่ ' + n + ' คน', 'ดึงจากแบบฟอร์ม');
  });
}

/* ---------- นำเข้าจาก Google Form ---------- */

/** อ่านแถวคำตอบใหม่ที่ยังไม่เคยนำเข้า แตกผู้สมัครคนที่ 2–5 เป็นรายคน แล้วเพิ่มลงชีต "รายชื่อ" */
function importForm_(cfg) {
  const name = String(cfg['ชีตแบบฟอร์ม'] || DEFAULT_FORM_SHEET).trim();
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return 0;
  const props = PropertiesService.getDocumentProperties(), key = formKey_(cfg);
  const last = Number(props.getProperty(key) || 1); // แถวสุดท้ายที่นำเข้าแล้ว
  const n = sh.getLastRow();
  if (n <= last) return 0;
  const w = sh.getLastColumn();
  const h = sh.getRange(1, 1, 1, w).getValues()[0].map(x => String(x).trim());
  const col = t => h.indexOf(t);
  const groups = [[col('คำนำหน้า'), col('ชื่อ-นามสกุล')]];
  for (let k = 2; k <= 10; k++) groups.push([col('คำนำหน้า (คนที่ ' + k + ')'), col('ชื่อ-นามสกุล ' + k)]);
  const cDept = col('บริษัท/หน่วยงาน'), cPos = col('ตำแหน่ง');
  const list = [];
  sh.getRange(last + 1, 1, n - last, w).getValues().forEach(r => {
    groups.forEach(([pi, ni], gi) => {
      if (ni < 0 || !clean_(r[ni])) return;
      list.push({ pre: pi >= 0 ? r[pi] : '', name: r[ni], dept: cDept >= 0 ? r[cDept] : '',
        pos: gi === 0 && cPos >= 0 ? r[cPos] : '' });
    });
  });
  const res = add_(list, 'แบบฟอร์ม');
  props.setProperty(key, String(n));
  return res.added.length;
}
function formKey_(cfg) { return 'formRows:' + String(cfg['ชีตแบบฟอร์ม'] || DEFAULT_FORM_SHEET).trim(); }

/* ---------- ตัวช่วย ---------- */

function add_(list, src) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SH.people);
  const cur = people_();
  const seen = new Set(cur.map(p => norm_(p.name)));
  let next = cur.reduce((m, p) => { const x = /^P(\d+)$/.exec(p.id); return x ? Math.max(m, +x[1]) : m; }, 0) + 1;
  const added = [], rows = [];
  let skipped = 0;
  (list || []).forEach(p => {
    const sp = splitPre_(p && p.pre, p && p.name);
    if (!sp.name) return;
    if (seen.has(norm_(sp.name))) { skipped++; return; }
    seen.add(norm_(sp.name));
    const o = { id: id_(next++), pre: sp.pre, name: sp.name, sex: sex_(p.sex, sp.pre),
      dept: clean_(p.dept), pos: clean_(p.pos) };
    added.push(o);
    rows.push([o.id, o.pre, o.name, o.sex, o.dept, o.pos, new Date(), src]);
  });
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, HEAD.people.length).setValues(rows);
  return { added: added, skipped: skipped };
}

/** อ่านรายชื่อ และเติมรหัสให้แถวที่พิมพ์ในชีตแต่ยังไม่มีรหัส */
function people_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SH.people);
  const n = sh.getLastRow();
  if (n < 2) return [];
  const v = sh.getRange(2, 1, n - 1, 6).getValues();
  let next = v.reduce((m, r) => { const x = /^P(\d+)$/.exec(String(r[P.id]).trim()); return x ? Math.max(m, +x[1]) : m; }, 0) + 1;
  let changed = false;
  v.forEach(r => {
    if (!clean_(r[P.name])) return;
    if (!String(r[P.id]).trim()) { r[P.id] = id_(next++); changed = true; }
    if (!clean_(r[P.sex])) { r[P.sex] = sex_('', splitPre_(r[P.pre], r[P.name]).pre); changed = true; }
  });
  if (changed) {
    sh.getRange(2, 1, n - 1, 1).setValues(v.map(r => [r[P.id]]));
    sh.getRange(2, P.sex + 1, n - 1, 1).setValues(v.map(r => [r[P.sex]]));
  }
  return v.filter(r => clean_(r[P.name])).map(r => {
    const sp = splitPre_(r[P.pre], r[P.name]);
    return { id: String(r[P.id]).trim(), pre: sp.pre, name: sp.name, sex: sex_(r[P.sex], sp.pre),
      dept: clean_(r[P.dept]), pos: clean_(r[P.pos]) };
  });
}

/** แยกคำนำหน้าที่พิมพ์ติดมากับชื่อ เช่น "นางสาวณัฏชา แสวงหอม" */
function splitPre_(pre, name) {
  pre = clean_(pre); name = clean_(name);
  const m = /^(นางสาว|นาย|นาง|น\.ส\.)\s*/.exec(name);
  if (m) { pre = pre || (m[1] === 'น.ส.' ? 'นางสาว' : m[1]); name = name.slice(m[0].length); }
  return { pre: pre, name: name };
}

/** เพศ: ใช้ค่าที่กรอก ถ้าว่างดูจากคำนำหน้า */
function sex_(v, pre) {
  v = clean_(v);
  if (/^(ชาย|ช|m|male)$/i.test(v)) return 'ชาย';
  if (/^(หญิง|ญ|f|female)$/i.test(v)) return 'หญิง';
  if (/^(นาย|mr\.?)$/i.test(pre)) return 'ชาย';
  if (/^(นาง|นางสาว|mrs\.?|ms\.?|miss)$/i.test(pre)) return 'หญิง';
  return 'ไม่ระบุ';
}

function cfg_() {
  const o = {};
  rows_(SH.cfg).forEach(r => { if (String(r[0]).trim()) o[String(r[0]).trim()] = String(r[1]).trim(); });
  return o;
}

function rows_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  const n = sh.getLastRow();
  if (n < 2) return [];
  const w = { [SH.log]: HEAD.log.length, [SH.people]: HEAD.people.length }[name] || 2;
  return sh.getRange(2, 1, n - 1, w).getValues();
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function checkDate_(d) {
  d = String(d || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) throw new Error('รูปแบบวันที่ไม่ถูกต้อง: ' + d);
  return d;
}
function checkSes_(s) {
  s = String(s || '').trim();
  if (SESSIONS.indexOf(s) < 0) throw new Error('ช่วงเวลาไม่ถูกต้อง: ' + s);
  return s;
}
function clean_(s) { s = String(s == null ? '' : s).replace(/\s+/g, ' ').trim(); return s === '-' ? '' : s; }
function isAbsent_(r) { return String(r[L.status] || '').trim() === 'ขาด'; }
function day_(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v).trim(); }
function time_(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'HH:mm:ss') : String(v).trim(); }
function id_(n) { return 'P' + String(n).padStart(4, '0'); }
function norm_(s) { return clean_(s).toLowerCase(); }
