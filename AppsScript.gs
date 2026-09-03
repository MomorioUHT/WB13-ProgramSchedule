/**
 * CombinedScraper - Google Apps Script Web App
 *
 * Deploy:
 *   1. เปิด Google Sheet ที่จะใช้เก็บข้อมูล -> Extensions -> Apps Script
 *   2. วางโค้ดนี้ทั้งหมด แล้วบันทึก
 *   3. Deploy -> New deployment -> type: Web app
 *        - Execute as: Me
 *        - Who has access: Anyone
 *   4. คัดลอก Web app URL (.../exec) ไปใส่ APPS_SCRIPT_URL ใน program.py
 *
 * รับ request แบบ POST body = JSON:
 *   { "action": "create", "sheet": "พีพีทีวี (PPTV)" }
 *   { "action": "get",    "sheet": "พีพีทีวี (PPTV)", "range": "A2:C" }
 *   { "action": "put",    "sheet": "พีพีทีวี (PPTV)", "data": [[...],[...]], "corner": "A2" }
 *
 * ตอบกลับเป็น JSON: { ok: true, action: ..., result: ... }
 *                   { ok: false, error: "..." }
 */

var HEADERS = ['วัน', 'เวลา', 'รายการ', 'Facebook Link', 'Youtube Link', 'X Link', 'TikTok Link'];

// จำนวนคอลัมน์ที่โปรแกรมเขียน/เขียนทับ (วัน, เวลา, รายการ)
// คอลัมน์ที่เหลือ (Facebook/Youtube/X/TikTok) เป็นของผู้ใช้กรอกเอง โปรแกรมจะไม่แตะ
var PROGRAM_COLS = 3;

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var req = JSON.parse(e.postData.contents);
    var action = req.action;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var result;

    if (action === 'create') {
      result = createSheet(ss, req.sheet);
    } else if (action === 'get') {
      result = getData(ss, req.sheet, req.range, req.start, req.end);
    } else if (action === 'put') {
      result = putData(ss, req.sheet, req.data, req.corner);
    } else {
      throw new Error('unknown action: ' + action);
    }

    return jsonOut({ ok: true, action: action, result: result });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err && err.stack ? err.stack : err) });
  } finally {
    lock.releaseLock();
  }
}

/**
 * สร้างชีทใหม่พร้อมหัวคอลัมน์ ; ถ้ามีชีทอยู่แล้วจะไม่สร้างซ้ำ
 * แต่จะเขียน/อัปเดตหัวคอลัมน์ให้ตรงกับ HEADERS เสมอ (idempotent)
 */
function createSheet(ss, name) {
  requireName(name);
  var sh = ss.getSheetByName(name);
  var created = false;
  if (!sh) {
    sh = ss.insertSheet(name);
    created = true;
  }
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sh.setFrozenRows(1);
  // บังคับคอลัมน์ วัน/เวลา เป็น text กัน Sheets แปลงเป็นวันที่/เวลาอัตโนมัติ
  sh.getRange('A:B').setNumberFormat('@');
  return { created: created, sheet: name };
}

/** อ่านค่าในช่วงที่ระบุ (a1 notation) ; ถ้าไม่ส่ง range จะรวม start:end */
function getData(ss, name, range, start, end) {
  requireName(name);
  var sh = ss.getSheetByName(name);
  if (!sh) throw new Error('sheet not found: ' + name);

  var a1 = range;
  if (!a1 && start && end) {
    a1 = String(start).split(':')[0] + ':' + String(end).split(':').pop();
  }
  if (!a1) a1 = 'A2:C';

  var values = sh.getRange(a1).getValues();
  return { sheet: name, range: a1, values: values };
}

/** เขียนทับข้อมูลใต้หัวคอลัมน์ทั้งหมด แล้วใส่ data ใหม่ที่ corner (ดีฟอลต์ A2) */
function putData(ss, name, data, corner) {
  requireName(name);
  var sh = ss.getSheetByName(name);
  if (!sh) throw new Error('sheet not found: ' + name);

  data = data || [];
  corner = corner || 'A2';
  var rc = a1ToRowCol(corner);

  // ล้างเฉพาะคอลัมน์ที่โปรแกรมดูแล (วัน/เวลา/รายการ) ไม่แตะคอลัมน์ลิงก์ที่ผู้ใช้กรอกเอง
  var width = data.length > 0 ? data[0].length : PROGRAM_COLS;
  var lastRow = sh.getLastRow();
  if (lastRow >= rc.row) {
    sh.getRange(rc.row, rc.col, lastRow - rc.row + 1, width).clearContent();
  }

  if (data.length > 0) {
    sh.getRange(rc.row, rc.col, data.length, width)
      .setNumberFormat('@')
      .setValues(data);
  }
  return { sheet: name, rows: data.length };
}

/* ----------------------------- helpers ----------------------------- */

function requireName(name) {
  if (!name || typeof name !== 'string') {
    throw new Error('missing "sheet" name');
  }
}

function a1ToRowCol(a1) {
  var m = String(a1).match(/^([A-Za-z]+)(\d+)$/);
  if (!m) throw new Error('bad corner: ' + a1);
  var letters = m[1].toUpperCase();
  var col = 0;
  for (var i = 0; i < letters.length; i++) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }
  return { row: parseInt(m[2], 10), col: col };
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
