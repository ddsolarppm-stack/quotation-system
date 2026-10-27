/**
 * ระบบใบเสนอราคา — ฝั่งเซิร์ฟเวอร์ (Google Apps Script)
 *
 * หน้าเว็บ (index.html) โหลดข้อมูลทั้งหมดครั้งเดียวด้วย getAll() แล้วส่งเฉพาะรายการที่เปลี่ยนมาที่ sync()
 * ข้อมูลเก็บใน Google Sheet ที่ผูกกับสคริปต์นี้ แยกแท็บละประเภทข้อมูล
 */

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
  const ss = SpreadsheetApp.getActive();
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
  return 'สร้างแท็บเรียบร้อย';
}

/** อ่านข้อมูลทั้งหมดให้หน้าเว็บ */
function getAll() {
  const settings = {};
  rows_('Settings').forEach(r => { if (r.key) settings[r.key] = parse_(r.value, r.value); });
  return {
    settings: Object.keys(settings).length ? settings : null,
    customers: rows_('Customers').filter(r => r.id),
    products: rows_('Products').filter(r => r.id).map(r => Object.assign(r, { price: Number(r.price) || 0 })),
    templates: rows_('Templates').filter(r => r.id).map(r => Object.assign(parse_(r.data, {}), { id: r.id, name: r.name, group: r.group })),
    docs: rows_('Documents').filter(r => r.no).map(r => parse_(r.data, null)).filter(Boolean)
  };
}

/**
 * บันทึกเฉพาะรายการที่เปลี่ยน
 * payload = { customers:[], products:[], templates:[], docs:[], deleteTemplates:[id], settings:{} }
 * คืนค่า { renamed: { เลขเดิม: เลขใหม่ } } เมื่อมีคนออกเลขเอกสารซ้ำพร้อมกัน
 */
function sync(payload) {
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
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
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
