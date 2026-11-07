const RESPONSE_SHEET_NAME = 'คำตอบหน่วยเคลื่อนที่ 2570';
const SPREADSHEET_ID_PROPERTY = 'MOBILE_SURVEY_2570_SPREADSHEET_ID';
const ADMIN_NOTES_SHEET_NAME = 'หมายเหตุหน่วยเคลื่อนที่ 2570';
const TIME_ZONE = 'Asia/Bangkok';
const MAX_COMMENT_LENGTH = 1500;
const FISCAL_YEAR_START = '2026-10-01'; // ปีงบประมาณ 2570 เริ่ม 1 ต.ค. 2569

function onOpen() {
  SpreadsheetApp.getUi().createMenu('หน่วยเคลื่อนที่ 2570')
    .addItem('เปิด Dashboard ผู้ดูแล', 'showAdminDashboard')
    .addToUi();
}

function showAdminDashboard() {
  requireAdminSpreadsheet_();
  SpreadsheetApp.getUi().showModelessDialog(
    HtmlService.createHtmlOutputFromFile('Dashboard').setWidth(1100).setHeight(760),
    'Dashboard หน่วยรับบริจาคเลือดเคลื่อนที่ 2570'
  );
}

function requireAdminSpreadsheet_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  const configured = PropertiesService.getScriptProperties().getProperty(SPREADSHEET_ID_PROPERTY);
  if (!active || !configured || active.getId() !== configured) {
    throw new Error('กรุณาเปิด Dashboard จากเมนูใน Google Sheet ที่ตั้งค่ารับคำตอบไว้');
  }
  return active;
}

function getAdminDashboardData() {
  const spreadsheet = requireAdminSpreadsheet_();
  const responseSheet = getOrCreateResponseSheet_(spreadsheet);
  const timezone = spreadsheet.getSpreadsheetTimeZone() || 'Asia/Bangkok';
  const values = responseSheet.getDataRange().getValues();
  const headers = values[0] || [];
  const idx = name => headers.indexOf(name);
  const scoreColumns = GENERAL_QUESTIONS.map(q => idx(q.label)).filter(i => i >= 0);
  const daily = {};
  const comments = [];
  values.slice(1).forEach(row => {
    if (!(row[0] instanceof Date) || isNaN(row[0].getTime())) return;
    const date = Utilities.formatDate(row[0], timezone, 'yyyy-MM-dd');
    if (!daily[date]) daily[date] = { date: date, count: 0, scoreSum: 0, scoreCount: 0, unlikelyReturn: 0, unlikelyRecommend: 0, locations: {} };
    const day = daily[date];
    day.count++;
    const location = String(row[idx('ชื่อสถานที่ที่บริจาคเลือด')] || '').trim();
    if (location) day.locations[location] = (day.locations[location] || 0) + 1;
    scoreColumns.forEach(i => {
      const score = Number(row[i]);
      if (score >= 1 && score <= 5) { day.scoreSum += score; day.scoreCount++; }
    });
    if (row[idx('ความตั้งใจกลับมาใช้บริการ')] === 'ไม่น่ากลับมา') day.unlikelyReturn++;
    if (row[idx('การแนะนำต่อ')] === 'ไม่น่าแนะนำ') day.unlikelyRecommend++;
    const comment = [
      row[idx('ข้อเสนอแนะภาพรวม')],
      row[idx('เหตุผลที่ไม่น่ากลับมา')],
      row[idx('เหตุผลที่ไม่น่าแนะนำ')],
    ].filter(Boolean).map(String).join(' · ');
    if (comment) comments.push({ date: date, location: location, text: comment.slice(0, 1200) });
  });
  const notesSheet = getOrCreateAdminNotesSheet_(spreadsheet);
  const notes = {};
  notesSheet.getDataRange().getValues().slice(1).forEach(row => {
    const date = row[0] instanceof Date ? Utilities.formatDate(row[0], timezone, 'yyyy-MM-dd') : String(row[0] || '').slice(0, 10);
    if (date) notes[date] = String(row[1] || '');
  });
  return {
    today: Utilities.formatDate(new Date(), timezone, 'yyyy-MM-dd'),
    days: Object.keys(daily).sort().reverse().map(date => ({
      date: date,
      count: daily[date].count,
      average: daily[date].scoreCount ? Math.round(daily[date].scoreSum / daily[date].scoreCount * 100) / 100 : null,
      unlikelyReturn: daily[date].unlikelyReturn,
      unlikelyRecommend: daily[date].unlikelyRecommend,
      locations: Object.keys(daily[date].locations).map(name => ({ name: name, count: daily[date].locations[name] })),
    })),
    notes: notes,
    comments: comments.reverse().slice(0, 30),
  };
}

function saveDailyAdminNote(date, note) {
  const spreadsheet = requireAdminSpreadsheet_();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(date || ''))) throw new Error('กรุณาเลือกวันที่ให้ถูกต้อง');
  note = String(note || '').trim();
  if (note.length > 3000) throw new Error('หมายเหตุยาวได้ไม่เกิน 3,000 ตัวอักษร');
  const sheet = getOrCreateAdminNotesSheet_(spreadsheet);
  const timezone = spreadsheet.getSpreadsheetTimeZone() || 'Asia/Bangkok';
  const rows = sheet.getDataRange().getValues();
  let rowNumber = -1;
  for (let i = 1; i < rows.length; i++) {
    const savedDate = rows[i][0] instanceof Date ? Utilities.formatDate(rows[i][0], timezone, 'yyyy-MM-dd') : String(rows[i][0] || '').slice(0, 10);
    if (savedDate === date) { rowNumber = i + 1; break; }
  }
  const entry = [date, safeText_(note), new Date(), Session.getActiveUser().getEmail() || 'ผู้แก้ไขชีต'];
  if (rowNumber < 0) sheet.appendRow(entry);
  else sheet.getRange(rowNumber, 1, 1, entry.length).setValues([entry]);
  return { ok: true, date: date, note: note };
}

function getOrCreateAdminNotesSheet_(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(ADMIN_NOTES_SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(ADMIN_NOTES_SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['วันที่', 'สรุปและหมายเหตุประจำวัน', 'แก้ไขล่าสุด', 'ผู้แก้ไข']);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(2, 520);
    sheet.getRange(1, 1, 1, 4).setBackground('#803f53').setFontColor('#ffffff').setFontWeight('bold');
  }
  return sheet;
}

const STATIONS = [
  { key: 'registration', label: 'จุดลงทะเบียนผู้บริจาคเลือด' },
  { key: 'screening', label: 'จุดซักประวัติและคัดกรองผู้บริจาค' },
  { key: 'donation', label: 'จุดบริจาคเลือด' },
  { key: 'recovery', label: 'จุดพักฟื้นและรับเครื่องดื่มหลังบริจาค' },
];

const DIMENSIONS = [
  { key: 'tangibles', label: 'ความเป็นรูปธรรม' },
  { key: 'reliability', label: 'ความน่าเชื่อถือ' },
  { key: 'responsiveness', label: 'การตอบสนอง' },
  { key: 'assurance', label: 'การสร้างความมั่นใจ' },
  { key: 'empathy', label: 'การเอาใจใส่' },
];

const GENERAL_QUESTIONS = [
  { key: 'process_clear', label: 'ภาพรวม - ขั้นตอนการบริจาคเลือดมีลำดับชัดเจน' },
  { key: 'wait_time', label: 'ภาพรวม - ระยะเวลารอรับบริการเหมาะสม' },
  { key: 'facility_cleanliness', label: 'ภาพรวม - พื้นที่ให้บริการสะอาดและเป็นระเบียบ' },
  { key: 'overall_satisfaction', label: 'ภาพรวม - ความพึงพอใจต่อบริการโดยรวม' },
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('แบบประเมินความพึงพอใจ หน่วยรับบริจาคเลือดเคลื่อนที่ 2570')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** รับคำตอบจากหน้าแบบสอบถามบน GitHub Pages (ส่งเป็น JSON แบบ text/plain) */
function doPost(e) {
  let result;
  try {
    if (!e || !e.postData || !e.postData.contents) throw new Error('ไม่พบข้อมูลคำตอบ กรุณาลองส่งอีกครั้ง');
    let payload;
    try { payload = JSON.parse(e.postData.contents); }
    catch (err) { throw new Error('รูปแบบข้อมูลไม่ถูกต้อง กรุณาลองส่งอีกครั้ง'); }
    result = submitResponse(payload);
  } catch (err) {
    result = { ok: false, message: err && err.message ? err.message : 'บันทึกไม่สำเร็จ กรุณาลองใหม่อีกครั้ง' };
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

/** Run once from the Apps Script editor while the target Google Sheet is open. */
function setupSurvey() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('กรุณาเปิด Apps Script จาก Google Sheet ที่ต้องการเก็บคำตอบ');
  }
  PropertiesService.getScriptProperties().setProperty(SPREADSHEET_ID_PROPERTY, spreadsheet.getId());
  const sheet = getOrCreateResponseSheet_(spreadsheet);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, sheet.getLastColumn())
    .setBackground('#17324d').setFontColor('#ffffff').setFontWeight('bold').setWrap(true);
  sheet.setRowHeight(1, 48);
  sheet.getDataRange().setVerticalAlignment('middle');
  sheet.getDataRange().setWrap(true);
  sheet.autoResizeColumns(1, sheet.getLastColumn());
  sheet.setColumnWidth(1, 170);
  sheet.setColumnWidth(2, 115);
  sheet.setColumnWidth(3, 130);
  sheet.setColumnWidth(4, 150);
  return 'ตั้งค่าชีตสำหรับรับคำตอบเรียบร้อย';
}

function submitResponse(payload) {
  // honeypot: บอทมักกรอกช่องที่ซ่อนไว้ ตอบว่าสำเร็จแต่ไม่บันทึก
  if (payload && String(payload.website || '').trim()) {
    return { ok: true, message: 'บันทึกคำตอบเรียบร้อยแล้ว ขอบคุณที่ร่วมประเมินบริการ' };
  }
  validatePayload_(payload);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = getConfiguredResponseSheet_();
    const row = [
      new Date(),
      safeText_(payload.donationDate),
      safeText_(payload.donationLocation),
      safeText_(payload.ageGroup),
      safeText_(payload.donorExperience),
      'ยินยอม',
    ];

    STATIONS.forEach(station => {
      DIMENSIONS.forEach(dimension => {
        row.push(Number(payload.ratings[station.key][dimension.key]));
      });
    });
    GENERAL_QUESTIONS.forEach(question => row.push(Number(payload.ratings.general[question.key])));
    row.push(
      safeText_(payload.returnIntention),
      safeText_(payload.recommendation),
      ...STATIONS.map(station => safeText_(payload.comments[station.key] || '')),
      safeText_(payload.generalComment || ''),
      safeText_(payload.returnReason || ''),
      safeText_(payload.recommendationReason || ''),
    );

    sheet.appendRow(row);
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    return { ok: true, message: 'บันทึกคำตอบเรียบร้อยแล้ว ขอบคุณที่ร่วมประเมินบริการ' };
  } finally {
    lock.releaseLock();
  }
}

function getConfiguredResponseSheet_() {
  const id = PropertiesService.getScriptProperties().getProperty(SPREADSHEET_ID_PROPERTY);
  if (!id) throw new Error('ยังไม่ได้ตั้งค่าชีต กรุณาทำขั้นตอน setupSurvey ก่อนเผยแพร่แบบประเมิน');
  const spreadsheet = SpreadsheetApp.openById(id);
  return getOrCreateResponseSheet_(spreadsheet);
}

function getOrCreateResponseSheet_(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(RESPONSE_SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(RESPONSE_SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(buildHeaders_());
  } else {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const expected = buildHeaders_();
    if (headers.length < expected.length) {
      sheet.getRange(1, headers.length + 1, 1, expected.length - headers.length)
        .setValues([expected.slice(headers.length)]);
    }
  }
  return sheet;
}

function buildHeaders_() {
  const headers = ['เวลาที่ส่ง', 'วันที่บริจาค', 'ชื่อสถานที่ที่บริจาคเลือด', 'ช่วงอายุ', 'ประสบการณ์บริจาค', 'ความยินยอม'];
  STATIONS.forEach(station => DIMENSIONS.forEach(dimension => headers.push(station.label + ' - ' + dimension.label)));
  GENERAL_QUESTIONS.forEach(question => headers.push(question.label));
  headers.push(
    'ความตั้งใจกลับมาใช้บริการ',
    'การแนะนำต่อ',
    ...STATIONS.map(station => 'ข้อเสนอแนะ - ' + station.label),
    'ข้อเสนอแนะภาพรวม',
    'เหตุผลที่ไม่น่ากลับมา',
    'เหตุผลที่ไม่น่าแนะนำ',
  );
  return headers;
}

function validatePayload_(payload) {
  if (!payload || typeof payload !== 'object') throw new Error('ไม่พบข้อมูลคำตอบ กรุณาลองส่งอีกครั้ง');
  if (payload.consent !== true) throw new Error('ต้องให้ความยินยอมก่อนส่งแบบประเมิน');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(payload.donationDate || ''))) {
    throw new Error('กรุณาระบุวันที่บริจาคให้ถูกต้อง');
  }
  const today = Utilities.formatDate(new Date(), TIME_ZONE, 'yyyy-MM-dd');
  if (String(payload.donationDate) > today || String(payload.donationDate) < FISCAL_YEAR_START) {
    throw new Error('วันที่บริจาคต้องอยู่ในปีงบประมาณ 2570 และไม่เป็นวันในอนาคต');
  }
  if (!String(payload.donationLocation || '').trim() || String(payload.donationLocation).length > 200) {
    throw new Error('กรุณาระบุชื่อสถานที่ที่บริจาคเลือด');
  }
  if (!['17–30 ปี', '31–45 ปี', '46–60 ปี', 'มากกว่า 60 ปี'].includes(payload.ageGroup)) {
    throw new Error('กรุณาเลือกช่วงอายุ');
  }
  if (!['ครั้งแรก', '2–10 ครั้ง', '11–30 ครั้ง', 'มากกว่า 30 ครั้ง'].includes(payload.donorExperience)) {
    throw new Error('กรุณาเลือกประสบการณ์บริจาคเลือด');
  }
  if (!payload.ratings || !payload.comments) throw new Error('ข้อมูลแบบประเมินไม่ครบถ้วน');
  STATIONS.forEach(station => {
    if (!payload.ratings[station.key]) throw new Error('กรุณาตอบคะแนนให้ครบทุกจุดบริการ');
    DIMENSIONS.forEach(dimension => assertScore_(payload.ratings[station.key][dimension.key]));
    if (typeof (payload.comments[station.key] || '') !== 'string') throw new Error('รูปแบบข้อเสนอแนะไม่ถูกต้อง');
  });
  if (!payload.ratings.general) throw new Error('กรุณาตอบคำถามภาพรวมให้ครบ');
  GENERAL_QUESTIONS.forEach(question => assertScore_(payload.ratings.general[question.key]));
  if (!['กลับมาแน่นอน', 'น่าจะกลับมา', 'ยังไม่แน่ใจ', 'ไม่น่ากลับมา'].includes(payload.returnIntention)) {
    throw new Error('กรุณาเลือกความตั้งใจกลับมาใช้บริการ');
  }
  if (!['แนะนำแน่นอน', 'น่าจะแนะนำ', 'ยังไม่แน่ใจ', 'ไม่น่าแนะนำ'].includes(payload.recommendation)) {
    throw new Error('กรุณาเลือกการแนะนำต่อ');
  }
  if (payload.returnIntention === 'ไม่น่ากลับมา' && !String(payload.returnReason || '').trim()) {
    throw new Error('กรุณาระบุเหตุผลที่ไม่น่ากลับมา');
  }
  if (payload.recommendation === 'ไม่น่าแนะนำ' && !String(payload.recommendationReason || '').trim()) {
    throw new Error('กรุณาระบุเหตุผลที่ไม่น่าแนะนำ');
  }
}

function assertScore_(value) {
  const score = Number(value);
  if (!Number.isInteger(score) || score < 1 || score > 5) {
    throw new Error('กรุณาให้คะแนนทุกข้อด้วยคะแนน 1–5');
  }
}

function safeText_(value) {
  const text = String(value == null ? '' : value).trim().slice(0, MAX_COMMENT_LENGTH);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}
