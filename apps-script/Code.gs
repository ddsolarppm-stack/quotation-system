/**
 * ระบบใบเสนอราคา — ฝั่งเซิร์ฟเวอร์ (Google Apps Script)
 *
 * หน้าเว็บ (index.html) โหลดข้อมูลทั้งหมดครั้งเดียวด้วย getAll() แล้วส่งเฉพาะรายการที่เปลี่ยนมาที่ sync()
 * ข้อมูลเก็บใน Google Sheet ที่ผูกกับสคริปต์นี้ แยกแท็บละประเภทข้อมูล
 */

/** Google Sheet ที่ใช้เก็บข้อมูล (ID จากลิงก์ .../spreadsheets/d/<ID>/edit) */
const SPREADSHEET_ID = '1P-_gglnLqaN9ij9q3K-IEEPq32conXy0Pd_FlhGpvHs';
const book_ = () => SpreadsheetApp.openById(SPREADSHEET_ID);

const SHEETS = {
  Customers: ['id', 'name', 'tax', 'branch', 'addr', 'contact', 'phone'],
  Products: ['id', 'code', 'name', 'cat', 'unit', 'price'],
  Templates: ['id', 'name', 'group', 'data'],
  Documents: ['no', 'type', 'kind', 'status', 'date', 'customerId', 'customerName', 'project', 'ref', 'grand', 'uid', 'updatedAt', 'data'],
  Settings: ['key', 'value']
};

/** เปิดหน้าเว็บแอป */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ระบบใบเสนอราคา')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/** รันครั้งแรกครั้งเดียว: สร้างแท็บและหัวคอลัมน์ (รันซ้ำได้ ไม่ลบข้อมูลเดิม) */
function setup() {
  const ss = book_();
  Object.keys(SHEETS).forEach(name => {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    const cols = SHEETS[name];
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold').setBackground('#f9a825');
    sh.setFrozenRows(1);
    // เก็บทุกคอลัมน์เป็นข้อความ กันเลขผู้เสียภาษี/เบอร์โทรที่ขึ้นต้นด้วย 0 หาย
    sh.getRange(1, 1, sh.getMaxRows(), cols.length).setNumberFormat('@');
  });
  const first = ss.getSheetByName('Sheet1') || ss.getSheetByName('ชีต1');
  if (first && first.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(first);
  ensureAdmin_();
  return 'สร้างแท็บเรียบร้อย';
}

/* ---------- ผู้ใช้และรหัสผ่าน ----------
   รหัสผ่านเก็บใน Script Properties (ไม่อยู่ใน HTML และไม่อยู่ใน Sheet) แบบแฮช SHA-256 + salt
   ดู/ลบได้ที่ Apps Script → ⚙ Project Settings → Script Properties */
const DEFAULT_ADMIN = { id: 'admin', pw: 'PI123456' };
const SESSION_SEC = 12 * 3600;

function users_() { return parse_(PropertiesService.getScriptProperties().getProperty('users'), {}) || {}; }
function saveUsers_(u) { PropertiesService.getScriptProperties().setProperty('users', JSON.stringify(u)); }
function hash_(pw, salt) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + pw, Utilities.Charset.UTF_8));
}
function setPw_(u, pw) { u.salt = Utilities.getUuid(); u.hash = hash_(pw, u.salt); }
function ensureAdmin_() {
  const u = users_();
  if (!Object.keys(u).length) { u[DEFAULT_ADMIN.id] = { name: 'ผู้ดูแลระบบ', role: 'admin' }; setPw_(u[DEFAULT_ADMIN.id], DEFAULT_ADMIN.pw); saveUsers_(u); }
}
/** ลืมรหัส admin: รันฟังก์ชันนี้ใน Apps Script เพื่อตั้งรหัส admin กลับเป็นค่าเริ่มต้น */
function resetAdminPassword() {
  const u = users_();
  u[DEFAULT_ADMIN.id] = Object.assign(u[DEFAULT_ADMIN.id] || { name: 'ผู้ดูแลระบบ' }, { role: 'admin', disabled: false });
  setPw_(u[DEFAULT_ADMIN.id], DEFAULT_ADMIN.pw); saveUsers_(u);
  return 'ตั้งรหัส admin เป็น ' + DEFAULT_ADMIN.pw + ' แล้ว กรุณาเปลี่ยนหลังเข้าใช้';
}
function authErr_() { throw new Error('AUTH: กรุณาเข้าสู่ระบบใหม่'); }
/** ตรวจ token ทุกครั้งก่อนอ่าน/เขียนข้อมูล */
function auth_(token, needAdmin) {
  if (!token) authErr_();
  const s = parse_(CacheService.getScriptCache().get('tok:' + token), null);
  if (!s) authErr_();
  const u = users_()[s.id];
  if (!u || u.disabled) authErr_();
  if (needAdmin && u.role !== 'admin') throw new Error('เฉพาะผู้ดูแลระบบเท่านั้น');
  CacheService.getScriptCache().put('tok:' + token, JSON.stringify(s), SESSION_SEC);
  return { id: s.id, name: u.name, role: u.role };
}
function login(id, pw) {
  ensureAdmin_();
  id = String(id || '').trim().toLowerCase();
  const cache = CacheService.getScriptCache(), fk = 'fail:' + id, fails = Number(cache.get(fk)) || 0;
  if (fails >= 5) throw new Error('ใส่รหัสผิดหลายครั้ง กรุณารอ 15 นาทีแล้วลองใหม่');
  const u = users_()[id];
  if (!u || u.disabled || hash_(String(pw || ''), u.salt) !== u.hash) {
    cache.put(fk, String(fails + 1), 900);
    throw new Error('ID หรือรหัสผ่านไม่ถูกต้อง');
  }
  cache.remove(fk);
  const token = Utilities.getUuid();
  cache.put('tok:' + token, JSON.stringify({ id: id }), SESSION_SEC);
  return { token: token, user: { id: id, name: u.name, role: u.role } };
}
function logout(token) { if (token) CacheService.getScriptCache().remove('tok:' + token); return true; }
function changePassword(token, oldPw, newPw) {
  const me = auth_(token), u = users_();
  if (hash_(String(oldPw || ''), u[me.id].salt) !== u[me.id].hash) throw new Error('รหัสผ่านเดิมไม่ถูกต้อง');
  if (String(newPw || '').length < 8) throw new Error('รหัสผ่านใหม่ต้องมีอย่างน้อย 8 ตัวอักษร');
  setPw_(u[me.id], newPw); saveUsers_(u);
  return true;
}
/* ผู้ดูแลระบบจัดการผู้ใช้ */
function listUsers(token) {
  auth_(token, true);
  const u = users_();
  return Object.keys(u).sort().map(id => ({ id: id, name: u[id].name || '', role: u[id].role, disabled: !!u[id].disabled }));
}
function saveUser(token, data) {
  const me = auth_(token, true), u = users_();
  const id = String(data.id || '').trim().toLowerCase();
  if (!/^[a-z0-9._-]{3,30}$/.test(id)) throw new Error('ID ใช้ได้เฉพาะ a-z 0-9 . _ - ยาว 3–30 ตัว');
  const isNew = !u[id];
  if (isNew && String(data.password || '').length < 8) throw new Error('รหัสผ่านต้องมีอย่างน้อย 8 ตัวอักษร');
  if (data.password && String(data.password).length < 8) throw new Error('รหัสผ่านต้องมีอย่างน้อย 8 ตัวอักษร');
  const role = data.role === 'admin' ? 'admin' : 'user';
  if (id === me.id && role !== 'admin') throw new Error('ไม่สามารถลดสิทธิ์ของตัวเองได้');
  u[id] = Object.assign(u[id] || {}, { name: String(data.name || '').trim(), role: role, disabled: !!data.disabled && id !== me.id });
  if (data.password) setPw_(u[id], String(data.password));
  saveUsers_(u);
  return listUsers(token);
}
function deleteUser(token, id) {
  const me = auth_(token, true), u = users_();
  if (id === me.id) throw new Error('ไม่สามารถลบบัญชีของตัวเองได้');
  delete u[id]; saveUsers_(u);
  return listUsers(token);
}

/** อ่านข้อมูลทั้งหมดให้หน้าเว็บ */
function getAll(token) {
  const me = auth_(token);
  const settings = {};
  rows_('Settings').forEach(r => { if (r.key) settings[r.key] = parse_(r.value, r.value); });
  return {
    settings: Object.keys(settings).length ? settings : null,
    customers: rows_('Customers').filter(r => r.id),
    products: rows_('Products').filter(r => r.id).map(r => Object.assign(r, { price: Number(r.price) || 0 })),
    templates: rows_('Templates').filter(r => r.id).map(r => Object.assign(parse_(r.data, {}), { id: r.id, name: r.name, group: r.group })),
    docs: rows_('Documents').filter(r => r.no).map(r => parse_(r.data, null)).filter(Boolean),
    me: me
  };
}

/**
 * บันทึกเฉพาะรายการที่เปลี่ยน
 * payload = { customers:[], products:[], templates:[], docs:[], deleteTemplates:[id], settings:{} }
 * คืนค่า { renamed: { เลขเดิม: เลขใหม่ } } เมื่อมีคนออกเลขเอกสารซ้ำพร้อมกัน
 */
function sync(token, payload) {
  auth_(token);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const renamed = {};
    (payload.customers || []).forEach(c => upsert_('Customers', 'id', c, c));
    (payload.products || []).forEach(p => upsert_('Products', 'id', p, p));
    (payload.templates || []).forEach(t => {
      const data = Object.assign({}, t); delete data.id; delete data.name; delete data.group;
      upsert_('Templates', 'id', t, { id: t.id, name: t.name, group: t.group, data: JSON.stringify(data) });
    });
    (payload.deleteTemplates || []).forEach(id => remove_('Templates', 'id', id));
    (payload.docs || []).forEach(d => {
      // ถ้าเลขนี้ถูกใช้ไปแล้วโดยเอกสารอื่น (คนละ uid) ให้ออกเลขถัดไปแทน
      const exist = find_('Documents', 'no', d.no);
      if (exist && exist.uid && d.uid && exist.uid !== d.uid) {
        const nn = nextNo_(d.no);
        renamed[d.no] = nn; d.no = nn;
      }
      if (d.ref && renamed[d.ref]) d.ref = renamed[d.ref];
      upsert_('Documents', 'no', d, {
        no: d.no, type: d.type, kind: d.kind || '', status: d.status, date: d.date, customerId: d.custId,
        customerName: d.customerName || '', project: d.project || '', ref: d.ref || '', grand: d.grand || '',
        uid: d.uid || '', updatedAt: new Date().toISOString(), data: JSON.stringify(d)
      });
    });
    if (payload.settings) {
      Object.keys(payload.settings).forEach(k => {
        const v = payload.settings[k];
        upsert_('Settings', 'key', { key: k }, { key: k, value: JSON.stringify(v) });
      });
    }
    return { renamed: renamed };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- ตัวช่วยอ่าน/เขียนชีต ---------- */
function sheet_(name) {
  const sh = book_().getSheetByName(name);
  if (!sh) throw new Error('ไม่พบแท็บ ' + name + ' กรุณารันฟังก์ชัน setup ก่อน');
  return sh;
}
function rows_(name) {
  const sh = sheet_(name), n = sh.getLastRow();
  if (n < 2) return [];
  const cols = SHEETS[name];
  return sh.getRange(2, 1, n - 1, cols.length).getValues().map(v => {
    const o = {}; cols.forEach((c, i) => o[c] = v[i] === '' ? '' : String(v[i])); return o;
  });
}
function rowIndex_(name, keyCol, key) {
  const sh = sheet_(name), n = sh.getLastRow();
  if (n < 2) return -1;
  const ci = SHEETS[name].indexOf(keyCol) + 1;
  const keys = sh.getRange(2, ci, n - 1, 1).getValues();
  for (let i = 0; i < keys.length; i++) if (String(keys[i][0]) === String(key)) return i + 2;
  return -1;
}
function find_(name, keyCol, key) {
  const r = rowIndex_(name, keyCol, key);
  if (r < 0) return null;
  const cols = SHEETS[name], v = sheet_(name).getRange(r, 1, 1, cols.length).getValues()[0], o = {};
  cols.forEach((c, i) => o[c] = String(v[i])); return o;
}
function upsert_(name, keyCol, keyObj, rowObj) {
  const sh = sheet_(name), cols = SHEETS[name];
  const values = [cols.map(c => rowObj[c] === undefined || rowObj[c] === null ? '' : String(rowObj[c]))];
  const r = rowIndex_(name, keyCol, keyObj[keyCol]);
  const target = r > 0 ? r : sh.getLastRow() + 1;
  sh.getRange(target, 1, 1, cols.length).setNumberFormat('@').setValues(values);
}
function remove_(name, keyCol, key) {
  const r = rowIndex_(name, keyCol, key);
  if (r > 0) sheet_(name).deleteRow(r);
}
function parse_(s, fallback) { try { return JSON.parse(s); } catch (e) { return fallback; } }

/** เลขถัดไปของชุดเดียวกัน เช่น QT202610-7 → ตัวนำหน้า QT202610- แล้วหาเลขสูงสุด + 1 */
function nextNo_(no) {
  const prefix = no.replace(/\d+$/, '');
  const max = rows_('Documents').filter(r => r.no.indexOf(prefix) === 0)
    .reduce((m, r) => Math.max(m, Number(r.no.slice(prefix.length)) || 0), 0);
  return prefix + (max + 1);
}
