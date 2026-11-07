/**
 * Dashboard แบบประเมินความพึงพอใจ หน่วยรับบริจาคเลือดเคลื่อนที่ ปีงบประมาณ 2570
 * โปรเจกต์ Apps Script แยกจากแบบสอบถาม — อ่านข้อมูลจากชีตคำตอบอย่างเดียว ไม่แก้ไขข้อมูล
 */

// ===== ตั้งค่า (แก้เฉพาะส่วนนี้) =====

// วางลิงก์ Google Sheet ที่เก็บคำตอบของหน่วยเคลื่อนที่ (คัดลอกจากแถบที่อยู่ของเบราว์เซอร์ได้ทั้งเส้น)
const SPREADSHEET_URL = 'PASTE_GOOGLE_SHEET_URL';

// รหัสผ่านเข้า Dashboard ไม่ได้เขียนไว้ในโค้ด ตั้งที่ ⚙️ การตั้งค่าโปรเจกต์ > คุณสมบัติของสคริปต์
//   ชื่อพร็อพเพอร์ตี: DASHBOARD_PASSWORD   ค่า: รหัสผ่านที่ต้องการ

// (ไม่บังคับ) จำกัดเพิ่มเฉพาะอีเมลเหล่านี้ — ปล่อยว่างไว้ = ใครก็ได้ที่รู้รหัสผ่าน
// ⚠️ ถ้าเปิด Dashboard ผ่าน GitHub Pages ต้องปล่อยว่าง เพราะระบบจะไม่รู้อีเมลผู้ใช้
const ALLOWED_EMAILS = [
  // 'somchai.xxx@mahidol.ac.th',
];

// ===== ค่าคงที่ของระบบ =====
const RESPONSE_SHEET_NAME = 'คำตอบหน่วยเคลื่อนที่ 2570';
const TIME_ZONE = 'Asia/Bangkok';
const PASSWORD_PROPERTY = 'DASHBOARD_PASSWORD';
const SESSION_SECONDS = 6 * 60 * 60;   // จำการเข้าสู่ระบบไว้ 6 ชั่วโมง (สูงสุดที่ CacheService รองรับ)
const MAX_FAILED_LOGINS = 10;         // ใส่รหัสผิดเกินนี้ ล็อก 10 นาที
const LOCK_SECONDS = 10 * 60;

const STATIONS = [
  { key: 'registration', label: 'จุดลงทะเบียนผู้บริจาคเลือด' },
  { key: 'screening', label: 'จุดซักประวัติและคัดกรองผู้บริจาค' },
  { key: 'donation', label: 'จุดบริจาคเลือด' },
  { key: 'recovery', label: 'จุดพักฟื้นและรับเครื่องดื่มหลังบริจาค' },
];
const DIMENSIONS = ['ความเป็นรูปธรรม', 'ความน่าเชื่อถือ', 'การตอบสนอง', 'การสร้างความมั่นใจ', 'การเอาใจใส่'];
const GENERAL_HEADERS = [
  'ภาพรวม - ขั้นตอนการบริจาคเลือดมีลำดับชัดเจน',
  'ภาพรวม - ระยะเวลารอรับบริการเหมาะสม',
  'ภาพรวม - พื้นที่ให้บริการสะอาดและเป็นระเบียบ',
  'ภาพรวม - ความพึงพอใจต่อบริการโดยรวม',
];

function doGet() {
  if (!isAllowed_()) {
    return HtmlService.createHtmlOutput(
      '<div style="font-family:Tahoma,sans-serif;max-width:520px;margin:80px auto;padding:28px;border:1px solid #eedfe3;border-radius:16px;text-align:center;color:#3b2a30">' +
      '<h2 style="color:#8b2f4c;margin-top:0">ไม่มีสิทธิ์เข้าถึง Dashboard</h2>' +
      '<p>บัญชีที่ใช้อยู่: <b>' + escapeHtml_(currentEmail_() || 'ไม่ทราบ') + '</b></p></div>'
    ).setTitle('ไม่มีสิทธิ์เข้าถึง');
  }
  try {
    return HtmlService.createHtmlOutputFromFile('Dashboard')
      .setTitle('Dashboard ความพึงพอใจ หน่วยรับบริจาคเลือดเคลื่อนที่ 2570')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (err) {
    // โปรเจกต์นี้ใช้เป็น API ให้หน้า Dashboard บน GitHub Pages อย่างเดียว
    return HtmlService.createHtmlOutput('<p style="font-family:Tahoma,sans-serif;text-align:center;margin-top:80px">Dashboard API พร้อมใช้งาน — เปิด Dashboard ผ่านลิงก์ GitHub Pages</p>');
  }
}

/** API สำหรับหน้า Dashboard บน GitHub Pages — รับ JSON {action, args} แบบ text/plain */
function doPost(e) {
  let out;
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const args = Array.isArray(body.args) ? body.args : [];
    const actions = { login: login, logout: logout, getDashboardData: getDashboardData };
    if (!actions[body.action]) throw new Error('คำสั่งไม่ถูกต้อง');
    out = { ok: true, result: actions[body.action].apply(null, args) };
  } catch (err) {
    out = { ok: false, error: err && err.message ? err.message : String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/** ตรวจรหัสผ่าน คืน token สำหรับใช้งานต่อ */
function login(password) {
  if (!isAllowed_()) return { ok: false, message: 'บัญชีนี้ไม่มีสิทธิ์เข้าถึง Dashboard' };
  const expected = PropertiesService.getScriptProperties().getProperty(PASSWORD_PROPERTY);
  if (!expected) return { ok: false, message: 'ผู้ดูแลยังไม่ได้ตั้งรหัสผ่าน (DASHBOARD_PASSWORD)' };
  const cache = CacheService.getScriptCache();
  const failKey = 'fail_' + (currentEmail_() || 'anonymous');
  const fails = Number(cache.get(failKey) || 0);
  if (fails >= MAX_FAILED_LOGINS) {
    return { ok: false, message: 'ใส่รหัสผ่านผิดหลายครั้ง กรุณารอ 10 นาทีแล้วลองใหม่' };
  }
  if (!safeEqual_(String(password || ''), expected)) {
    cache.put(failKey, String(fails + 1), LOCK_SECONDS);
    const left = MAX_FAILED_LOGINS - fails - 1;
    return { ok: false, message: left > 0 ? 'รหัสผ่านไม่ถูกต้อง (ลองได้อีก ' + left + ' ครั้ง)' : 'ใส่รหัสผ่านผิดหลายครั้ง กรุณารอ 10 นาทีแล้วลองใหม่' };
  }
  cache.remove(failKey);
  const token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
  cache.put('sess_' + token, currentEmail_() || 'anonymous', SESSION_SECONDS);
  return { ok: true, token: token };
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove('sess_' + token);
  return true;
}

function requireSession_(token) {
  if (!isAllowed_()) throw new Error('ไม่มีสิทธิ์เข้าถึงข้อมูล');
  if (!token || !/^[\w-]{40,100}$/.test(String(token)) || !CacheService.getScriptCache().get('sess_' + token)) {
    throw new Error('SESSION_EXPIRED');
  }
}

function safeEqual_(a, b) {
  const ha = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, a, Utilities.Charset.UTF_8);
  const hb = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, b, Utilities.Charset.UTF_8);
  let diff = 0;
  for (let i = 0; i < ha.length; i++) diff |= ha[i] ^ hb[i];
  return diff === 0;
}

/** เรียกจากหน้า Dashboard — ส่งข้อมูลแบบย่อให้หน้าเว็บคำนวณและกรองเอง */
function getDashboardData(token) {
  requireSession_(token);
  const sheet = openResponseSheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return { rows: [], generatedAt: nowText_(), viewer: currentEmail_() };

  const header = values[0].map(h => String(h).trim());
  const col = name => {
    const i = header.indexOf(name);
    if (i < 0) throw new Error('ไม่พบคอลัมน์ "' + name + '" ในชีต กรุณาตรวจสอบหัวตาราง');
    return i;
  };
  const scoreCols = [];
  STATIONS.forEach(s => DIMENSIONS.forEach(d => scoreCols.push(col(s.label + ' - ' + d))));
  GENERAL_HEADERS.forEach(h => scoreCols.push(col(h)));
  const textCols = [
    col('ความตั้งใจกลับมาใช้บริการ'), col('การแนะนำต่อ'),
    ...STATIONS.map(s => col('ข้อเสนอแนะ - ' + s.label)),
    col('ข้อเสนอแนะภาพรวม'), col('เหตุผลที่ไม่น่ากลับมา'), col('เหตุผลที่ไม่น่าแนะนำ'),
  ];
  const cLoc = col('ชื่อสถานที่ที่บริจาคเลือด');
  const cTime = col('เวลาที่ส่ง'), cDate = col('วันที่บริจาค'), cAge = col('ช่วงอายุ'), cExp = col('ประสบการณ์บริจาค');

  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const v = values[r];
    if (v.every(x => x === '' || x === null)) continue;
    const submitted = v[cTime] instanceof Date ? v[cTime] : null;
    const submittedText = submitted ? Utilities.formatDate(submitted, TIME_ZONE, "yyyy-MM-dd'T'HH:mm") : '';
    const donation = dateText_(v[cDate]) || submittedText.slice(0, 10);
    if (!donation) continue;
    rows.push([
      donation,
      submittedText,
      String(v[cAge] || ''),
      String(v[cExp] || ''),
      scoreCols.map(c => { const n = Number(v[c]); return n >= 1 && n <= 5 ? n : null; }),
      textCols.map(c => stripFormulaGuard_(v[c])),
      stripFormulaGuard_(v[cLoc]).replace(/\s+/g, ' ').trim(),
    ]);
  }
  return { rows: rows, generatedAt: nowText_(), viewer: currentEmail_() };
}

/** รันครั้งเดียวจากหน้าแก้ไขสคริปต์ เพื่ออนุญาตสิทธิ์และตรวจว่าเปิดชีตได้ */
function testConnection() {
  if (!PropertiesService.getScriptProperties().getProperty(PASSWORD_PROPERTY)) {
    Logger.log('⚠️ ยังไม่ได้ตั้งรหัสผ่าน: ไปที่ ⚙️ การตั้งค่าโปรเจกต์ > คุณสมบัติของสคริปต์ > เพิ่ม DASHBOARD_PASSWORD');
  }
  const sheet = openResponseSheet_();
  const message = 'เชื่อมต่อสำเร็จ: ' + sheet.getParent().getName() + ' / ' + sheet.getName() +
    ' มีคำตอบ ' + Math.max(sheet.getLastRow() - 1, 0) + ' แถว';
  Logger.log(message);
  return message;
}

function openResponseSheet_() {
  const match = String(SPREADSHEET_URL).match(/[-\w]{25,}/);
  if (!match || SPREADSHEET_URL.indexOf('PASTE_') === 0) {
    throw new Error('ยังไม่ได้ตั้งค่า SPREADSHEET_URL ใน Code.gs');
  }
  const sheet = SpreadsheetApp.openById(match[0]).getSheetByName(RESPONSE_SHEET_NAME);
  if (!sheet) throw new Error('ไม่พบแท็บ "' + RESPONSE_SHEET_NAME + '" ในชีต');
  return sheet;
}

function currentEmail_() {
  return String(Session.getActiveUser().getEmail() || '').toLowerCase();
}

function isAllowed_() {
  const list = ALLOWED_EMAILS.map(e => String(e).trim().toLowerCase()).filter(Boolean);
  if (!list.length) return true; // ไม่ได้จำกัดอีเมล ใช้รหัสผ่านอย่างเดียว
  const email = currentEmail_();
  if (!email) return false;
  const owner = String(Session.getEffectiveUser().getEmail() || '').toLowerCase();
  if (email === owner) return true;
  return list.includes(email);
}

function dateText_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, TIME_ZONE, 'yyyy-MM-dd');
  const s = String(value || '').trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : '';
}

function stripFormulaGuard_(value) {
  const s = String(value == null ? '' : value);
  return s.charAt(0) === "'" ? s.slice(1) : s;
}

function nowText_() {
  return Utilities.formatDate(new Date(), TIME_ZONE, 'yyyy-MM-dd HH:mm');
}

function escapeHtml_(s) {
  return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}
