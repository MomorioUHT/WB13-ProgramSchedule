/**
 * CombinedScraper & ViewStatsScraper - Google Apps Script Web App
 *
 * Deploy:
 *   1. เปิด Google Sheet ที่จะใช้เก็บข้อมูล -> Extensions -> Apps Script
 *   2. วางโค้ดนี้ทั้งหมดแทนที่โค้ดเดิม แล้วกด Save
 *   3. Deploy -> Manage deployments -> กดไอคอนดินสอ (Edit) -> เลือก Version: New version -> กด Deploy
 *   4. ใช้ Web app URL (.../exec) ใน .env: POST_SCRIPT_API="..."
 *
 * รองรับ:
 *   1. create / get / put (ระบบตารางรายการเดิมของ CombinedScraper)
 *   2. get_all / list_sheets (ดึงข้อมูลผังรายการของทุกช่องพร้อมกัน)
 *   3. write_row / update_urls (เขียนลิงก์สด FB/YT/X/TikTok ที่ Crawl เจอกลับลงตารางผัง)
 *   4. append_view_stats (บันทึกยอดวิวย้อนหลังลงชีท 'View Stats' คอลัมน์ A-L)
 *   5. doGet(e) (ดึงข้อมูลแบบรวดเร็วผ่าน HTTP GET)
 */

var HEADERS = ['วัน', 'เวลา', 'รายการ', 'Facebook Link', 'Youtube Link', 'X Link', 'TikTok Link'];

var VIEW_STATS_HEADERS = [
  'วันที่',
  'ช่อง',
  'ชื่อรายการ',
  'เวลาเริ่มในผัง',
  'Facebook',
  'YouTube',
  'TikTok',
  'X (Twitter)',
  'Facebook Peak Time',
  'Facebook Peak View',
  'YouTube Peak Time',
  'YouTube Peak View',
  'TikTok Peak Time',
  'TikTok Peak View',
  'X Peak Time',
  'X Peak View'
];

// จำนวนคอลัมน์ที่โปรแกรมเขียน/เขียนทับ (วัน, เวลา, รายการ)
var PROGRAM_COLS = 3;
var START_ROW = 2;

/* ----------------------------- HTTP GET ----------------------------- */

function doGet(e) {
  try {
    // [ADDED] Live View Stats dashboard endpoint — อ่านชีท 'View Stats' ทั้ง 16 คอลัมน์ (อ่านอย่างเดียว)
    // รองรับ JSONP: ส่ง ?callback=fnName เพื่อเลี่ยงปัญหา CORS เมื่อเปิดไฟล์แบบ file://
    if (e && e.parameter && (e.parameter.action === 'view_stats' || e.parameter.action === 'get_view_stats')) {
      var vsPayload = buildViewStatsPayload_();
      var vsCb = e.parameter.callback;
      if (vsCb) {
        return ContentService
          .createTextOutput(vsCb + '(' + JSON.stringify(vsPayload) + ');')
          .setMimeType(ContentService.MimeType.JAVASCRIPT);
      }
      return jsonOut(vsPayload);
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetParam = (e && e.parameter && e.parameter.sheet) ? e.parameter.sheet.trim() : '';

    // หากระบุชื่อชีท ให้ดึงเฉพาะชีทนั้น
    if (sheetParam) {
      var sh = ss.getSheetByName(sheetParam);
      if (!sh) {
        return jsonOut({ ok: false, status: 'error', error: 'sheet not found: ' + sheetParam });
      }
      var rows = extractSheetRows_(sh);
      return jsonOut({ ok: true, status: 'ok', sheet: sheetParam, total: rows.length, data: rows });
    }

    // หากไม่ระบุ ให้ดึงข้อมูลผังรายการของทุกช่อง (ยกเว้นชีท 'View Stats')
    var allSheets = ss.getSheets();
    var channelData = {};
    var totalCount = 0;

    allSheets.forEach(function (sh) {
      var sName = sh.getName();
      if (sName.trim().toLowerCase().indexOf('view stats') === 0) return;
      var rows = extractSheetRows_(sh);
      channelData[sName] = rows;
      totalCount += rows.length;
    });

    return jsonOut({
      ok: true,
      status: 'ok',
      total_count: totalCount,
      sheets: Object.keys(channelData),
      data: channelData
    });
  } catch (err) {
    return jsonOut({ ok: false, status: 'error', error: String(err && err.stack ? err.stack : err) });
  }
}

/* ----------------------------- HTTP POST ----------------------------- */

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var req = JSON.parse(e.postData.contents);
    var action = req.action || '';
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var result;

    // [ADDED] Live View Stats dashboard endpoint (อ่านอย่างเดียว)
    if (action === 'view_stats' || action === 'get_view_stats') {
      return jsonOut(buildViewStatsPayload_());
    }

    // 1. บันทึกยอดวิว Peak View (One row per broadcast per day) ลงชีท "View Stats"
    if (action === 'append_view_stats' || action === 'upsert_view_stats' || req.target_sheet === 'View Stats') {
      result = upsertViewStats_(ss, req);
      return jsonOut({ ok: true, action: action, result: result });
    }

    // 2. ดึงผังรายการของทุกช่อง
    if (action === 'get_all' || action === 'list_sheets') {
      var allSheets = ss.getSheets();
      var channelData = {};
      var totalCount = 0;
      allSheets.forEach(function (sh) {
        var sName = sh.getName();
        if (sName.trim().toLowerCase().indexOf('view stats') === 0) return;
        var rows = extractSheetRows_(sh);
        channelData[sName] = rows;
        totalCount += rows.length;
      });
      return jsonOut({
        ok: true,
        action: action,
        total_count: totalCount,
        sheets: Object.keys(channelData),
        data: channelData
      });
    }

    // 3. เขียนผลลัพธ์ลิงก์สดกลับลงตารางผังรายการ (D=FB, E=YT, F=X, G=TikTok)
    if (action === 'write_row' || action === 'update_urls') {
      result = updateCrawledUrls_(ss, req);
      return jsonOut({ ok: true, action: action, result: result });
    }

    // 4. คำสั่งเดิม: create
    if (action === 'create') {
      result = createSheet(ss, req.sheet);
      return jsonOut({ ok: true, action: action, result: result });
    }

    // 5. คำสั่งเดิม: get
    if (action === 'get') {
      result = getData(ss, req.sheet, req.range, req.start, req.end);
      return jsonOut({ ok: true, action: action, result: result });
    }

    // 6. คำสั่งเดิม: put
    if (action === 'put') {
      result = putData(ss, req.sheet, req.data, req.corner);
      return jsonOut({ ok: true, action: action, result: result });
    }

    throw new Error('unknown action: ' + action);
  } catch (err) {
    return jsonOut({ ok: false, error: String(err && err.stack ? err.stack : err) });
  } finally {
    lock.releaseLock();
  }
}

/* ----------------------------- ฟังก์ชันเดิม (CombinedScraper) ----------------------------- */

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

/** อ่านค่าในช่วงที่ระบุ (a1 notation) ; ถ้าไม่ส่ง range จะอ่าน A2:G */
function getData(ss, name, range, start, end) {
  requireName(name);
  var sh = ss.getSheetByName(name);
  if (!sh) throw new Error('sheet not found: ' + name);

  var a1 = range;
  if (!a1 && start && end) {
    a1 = String(start).split(':')[0] + ':' + String(end).split(':').pop();
  }
  if (!a1) a1 = 'A2:G';

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

/* ----------------------------- ฟังก์ชันเพิ่มเติมสำหรับ Live Monitoring ----------------------------- */

/**
 * ดึงแถวข้อมูลตารางผังรายการจากชีทที่ระบุ โดยแปลงเป็น Array of Objects
 */
function extractSheetRows_(sh) {
  var lastRow = sh.getLastRow();
  if (lastRow < START_ROW) return [];

  var numRows = lastRow - START_ROW + 1;
  var rangeValues = sh.getRange(START_ROW, 1, numRows, 7).getValues();
  var rows = [];

  for (var i = 0; i < rangeValues.length; i++) {
    var r = rangeValues[i];
    var title = String(r[2] || '').trim();
    if (!title) continue;

    rows.append ? null : rows.push({
      row: START_ROW + i,
      date: formatCellValue_(r[0]),
      time: formatCellValue_(r[1]),
      title: title,
      facebook_url: formatCellValue_(r[3]),
      youtube_url: formatCellValue_(r[4]),
      x_url: formatCellValue_(r[5]),
      tiktok_url: formatCellValue_(r[6])
    });
  }
  return rows;
}

/**
 * เขียนลิงก์ที่ Crawl เจอ (D: Facebook, E: YouTube, F: X, G: TikTok) กลับลงตารางผังรายการ
 */
function updateCrawledUrls_(ss, req) {
  var targetSheet = req.sheet ? ss.getSheetByName(req.sheet) : ss.getActiveSheet();
  if (!targetSheet) {
    targetSheet = ss.getSheets()[0];
  }

  var updates = req.updates || (req.row ? [req] : []);
  var updatedResults = [];

  for (var i = 0; i < updates.length; i++) {
    var u = updates[i];
    var targetRow = parseInt(u.row, 10);
    if (!targetRow || targetRow < START_ROW) continue;

    // ตรวจสอบชื่อรายการก่อนเขียนเพื่อป้องกันการเลื่อนแถว
    if (u.title) {
      var currentTitle = String(targetSheet.getRange(targetRow, 3).getValue() || '').trim();
      if (currentTitle && currentTitle !== String(u.title).trim()) {
        updatedResults.push({ row: targetRow, status: 'skipped', reason: 'title mismatch' });
        continue;
      }
    }

    var rowVals = [
      u.facebook_url || '-',
      u.youtube_url || '-',
      u.x_url || '-',
      u.tiktok_url || '-'
    ];

    targetSheet.getRange(targetRow, 4, 1, 4).setValues([rowVals]);
    updatedResults.push({ row: targetRow, status: 'updated' });
  }

  return { sheet: targetSheet.getName(), updated: updatedResults };
}

/**
 * บันทึกยอดวิวแบบ Peak View (One row per broadcast per day) ลงชีท 'View Stats' (Columns 1-15)
 * บันทึก/เขียนทับเฉพาะเมื่อยอดวิวใหม่มากกว่ายอดวิวเดิมในชีท
 * แยกการบันทึก Peak Time และ Peak View ของแต่ละแพลตฟอร์มอิสระต่อกัน (FB, YT, TikTok, X)
 */
function upsertViewStats_(ss, req) {
  var sh = ss.getSheetByName('View Stats');
  if (sh) {
    var headerVal = String(sh.getRange(1, 1).getValue() || '').trim();
    if (headerVal && headerVal !== 'วันที่') {
      // ตรวจพบโครงสร้างคอลัมน์แบบเดิม สำรองชีทเก่าเก็บไว้เพื่อความปลอดภัยของข้อมูล
      var archiveName = 'View Stats (Archive ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Bangkok', 'yyyyMMdd_HHmm') + ')';
      try {
        sh.setName(archiveName);
        sh = null;
      } catch (e) {
        try {
          sh.setName('View Stats Archive ' + new Date().getTime());
          sh = null;
        } catch (e2) {}
      }
    }
  }

  if (!sh) {
    sh = ss.insertSheet('View Stats');
    sh.getRange(1, 1, 1, VIEW_STATS_HEADERS.length).setValues([VIEW_STATS_HEADERS]);
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('@');
    sh.getRange('D:D').setNumberFormat('@');
    sh.getRange('I:I').setNumberFormat('@');
    sh.getRange('K:K').setNumberFormat('@');
    sh.getRange('M:M').setNumberFormat('@');
    sh.getRange('O:O').setNumberFormat('@');
  }

  var rowsData = req.rows || (req.row ? [req.row] : []);
  if (!rowsData || rowsData.length === 0) {
    return { updated: 0, inserted: 0, preserved: 0, total_rows: sh.getLastRow() > 1 ? sh.getLastRow() - 1 : 0 };
  }

  var lastRow = sh.getLastRow();
  var existingRows = [];
  if (lastRow >= 2) {
    existingRows = sh.getRange(2, 1, lastRow - 1, VIEW_STATS_HEADERS.length).getValues();
  }

  // สร้างดัชนีระบุแถวตาม: "date___channel___title"
  var rowIndexMap = {};
  for (var i = 0; i < existingRows.length; i++) {
    var dStr = normalizeDateStr_(existingRows[i][0]);
    var chStr = String(existingRows[i][1] || '').trim().toLowerCase();
    var tStr = String(existingRows[i][2] || '').trim().toLowerCase();
    var key = dStr + '___' + chStr + '___' + tStr;
    if (key !== '______' && rowIndexMap[key] === undefined) {
      rowIndexMap[key] = i;
    }
  }

  var updatedCount = 0;
  var insertedCount = 0;
  var preservedCount = 0;

  for (var j = 0; j < rowsData.length; j++) {
    var r = rowsData[j];
    var rDate = normalizeDateStr_(r.date || '');
    var rTime = String(r.time || r.capture_time || r.capture_dt || '').trim();
    var rScheduledTime = String(r.scheduled_time || r.schedule_time || r.program_time || r.time_in_schedule || '').trim();
    var rChannel = String(r.channel_name || r.channel || '').trim();
    var rTitle = String(r.broadcast_name || r.title || '').trim();

    if (!rChannel || !rTitle) continue;

    var key = rDate + '___' + rChannel.toLowerCase() + '___' + rTitle.toLowerCase();

    var fbLink = r.facebook_live_link || r.facebook_url || '-';
    var ytLink = r.youtube_live_link || r.youtube_url || '-';
    var ttLink = r.tiktok_live_link || r.tiktok_url || '-';
    var xLink = r.x_live_link || r.x_url || '-';

    var fbViews = parseViewCount_(r.facebook_views);
    var ytViews = parseViewCount_(r.youtube_views);
    var ttViews = parseViewCount_(r.tiktok_views);
    var xViews = parseViewCount_(r.x_views);

    if (rowIndexMap[key] !== undefined) {
      var rowIdx = rowIndexMap[key];
      var targetRow = existingRows[rowIdx];
      var modified = false;

      // อัปเดตเวลาเริ่มในผังหากมีข้อมูลและในชีทยังว่างหรือเป็น '-'
      if (rScheduledTime && (!targetRow[3] || targetRow[3] === '-')) {
        targetRow[3] = rScheduledTime;
        modified = true;
      }

      // อัปเดตลิงก์หากได้ลิงก์สดที่ถูกต้องมาใหม่
      if (isValidUrl_(fbLink) && (!isValidUrl_(targetRow[4]) || targetRow[4] === '-')) {
        targetRow[4] = fbLink;
        modified = true;
      }
      if (isValidUrl_(ytLink) && (!isValidUrl_(targetRow[5]) || targetRow[5] === '-')) {
        targetRow[5] = ytLink;
        modified = true;
      }
      if (isValidUrl_(ttLink) && (!isValidUrl_(targetRow[6]) || targetRow[6] === '-')) {
        targetRow[6] = ttLink;
        modified = true;
      }
      if (isValidUrl_(xLink) && (!isValidUrl_(targetRow[7]) || targetRow[7] === '-')) {
        targetRow[7] = xLink;
        modified = true;
      }

      // 1. Facebook: เปรียบเทียบยอดวิวพีค (Col I: Peak Time, Col J: Peak View)
      var curFbPeak = parseViewCount_(targetRow[9]);
      if (fbViews >= 0 && (curFbPeak < 0 || fbViews > curFbPeak)) {
        targetRow[8] = rTime || targetRow[8] || '-';
        targetRow[9] = fbViews;
        modified = true;
      }

      // 2. YouTube: เปรียบเทียบยอดวิวพีค (Col K: Peak Time, Col L: Peak View)
      var curYtPeak = parseViewCount_(targetRow[11]);
      if (ytViews >= 0 && (curYtPeak < 0 || ytViews > curYtPeak)) {
        targetRow[10] = rTime || targetRow[10] || '-';
        targetRow[11] = ytViews;
        modified = true;
      }

      // 3. TikTok: เปรียบเทียบยอดวิวพีค (Col M: Peak Time, Col N: Peak View)
      var curTtPeak = parseViewCount_(targetRow[13]);
      if (ttViews >= 0 && (curTtPeak < 0 || ttViews > curTtPeak)) {
        targetRow[12] = rTime || targetRow[12] || '-';
        targetRow[13] = ttViews;
        modified = true;
      }

      // 4. X (Twitter): เปรียบเทียบยอดวิวพีค (Col O: Peak Time, Col P: Peak View)
      var curXPeak = parseViewCount_(targetRow[15]);
      if (xViews >= 0 && (curXPeak < 0 || xViews > curXPeak)) {
        targetRow[14] = rTime || targetRow[14] || '-';
        targetRow[15] = xViews;
        modified = true;
      }

      if (modified) {
        updatedCount++;
      } else {
        preservedCount++;
      }
    } else {
      // แถวใหม่ประจำวันนี้สำหรับรายการนี้
      var newRow = [
        rDate,
        rChannel,
        rTitle,
        rScheduledTime || '-',
        isValidUrl_(fbLink) ? fbLink : '-',
        isValidUrl_(ytLink) ? ytLink : '-',
        isValidUrl_(ttLink) ? ttLink : '-',
        isValidUrl_(xLink) ? xLink : '-',
        fbViews >= 0 ? (rTime || '-') : '-',
        fbViews >= 0 ? fbViews : '-',
        ytViews >= 0 ? (rTime || '-') : '-',
        ytViews >= 0 ? ytViews : '-',
        ttViews >= 0 ? (rTime || '-') : '-',
        ttViews >= 0 ? ttViews : '-',
        xViews >= 0 ? (rTime || '-') : '-',
        xViews >= 0 ? xViews : '-'
      ];
      existingRows.push(newRow);
      rowIndexMap[key] = existingRows.length - 1;
      insertedCount++;
    }
  }

  // บันทึกกลับลง Google Sheet
  if (existingRows.length > 0) {
    sh.getRange(2, 1, existingRows.length, VIEW_STATS_HEADERS.length).setValues(existingRows);
    // บังคับรูปแบบข้อความ (Plain Text) สำหรับคอลัมน์ วันที่, เวลาเริ่มในผัง และ เวลาพีค
    sh.getRange(2, 1, existingRows.length, 1).setNumberFormat('@');
    sh.getRange(2, 4, existingRows.length, 1).setNumberFormat('@');
    sh.getRange(2, 9, existingRows.length, 1).setNumberFormat('@');
    sh.getRange(2, 11, existingRows.length, 1).setNumberFormat('@');
    sh.getRange(2, 13, existingRows.length, 1).setNumberFormat('@');
    sh.getRange(2, 15, existingRows.length, 1).setNumberFormat('@');
  }

  return {
    updated: updatedCount,
    inserted: insertedCount,
    preserved: preservedCount,
    total_rows: existingRows.length
  };
}

/* ----------------------------- helpers ----------------------------- */

function formatCellValue_(val) {
  if (val === null || val === undefined) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone() || 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
  }
  return String(val).trim();
}

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

function parseViewCount_(val) {
  if (val === null || val === undefined) return -1;
  if (typeof val === 'number') return isNaN(val) ? -1 : Math.round(val);
  var s = String(val).replace(/,/g, '').trim();
  if (s === '' || s === '-' || s.toUpperCase() === 'N/A' || s === 'NOT FOUND') return -1;
  var n = parseInt(s, 10);
  return isNaN(n) ? -1 : n;
}

function normalizeDateStr_(d) {
  if (!d) return '';
  if (d instanceof Date) {
    return Utilities.formatDate(d, Session.getScriptTimeZone() || 'Asia/Bangkok', 'yyyy-MM-dd');
  }
  var s = String(d).trim();
  if (s.indexOf('T') > -1) {
    s = s.split('T')[0];
  }
  return s;
}

function isValidUrl_(url) {
  if (!url || typeof url !== 'string') return false;
  var s = url.trim();
  return s.indexOf('http://') === 0 || s.indexOf('https://') === 0;
}

/* ============================================================================
 * [ADDED] Live View Stats Dashboard — อ่านชีท 'View Stats' (หน้า Dashboard ที่ 2)
 * เป็นการอ่านอย่างเดียว ไม่แก้ไข/ไม่แตะโค้ดหรือชีทเดิม
 *
 * โครงสร้างชีท 'View Stats':
 *   A วันที่ | B ช่อง | C ชื่อรายการ | D เวลาเริ่มในผัง
 *   E Facebook Link | F YouTube Link | G TikTok Link | H X Link
 *   I Facebook Peak time | J Facebook Views
 *   K YouTube Peak time  | L YouTube Views
 *   M TikTok Peak time   | N TikTok Views
 *   O X Peak time        | P X Views
 *
 * เรียกใช้:  GET  <exec>?action=view_stats
 *           POST <exec>  body: {"action":"view_stats"}
 * ========================================================================== */

function buildViewStatsPayload_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('View Stats');
    if (!sh) return { ok: false, error: 'sheet not found: View Stats' };

    var lastRow = sh.getLastRow();
    if (lastRow < 2) {
      return { ok: true, sheet: 'View Stats', total: 0, generated_at: new Date().toISOString(), data: [] };
    }

    var values = sh.getRange(1, 1, lastRow, 16).getValues();
    var out = [];

    for (var i = 0; i < values.length; i++) {
      var rowNum = i + 1;
      if (rowNum < 2) continue; // ข้ามแถวหัวตาราง

      var r = values[i];
      var date = normalizeDateStr_(r[0]);
      var channel = String(r[1] || '').trim();
      var title = String(r[2] || '').trim();
      var scheduledTime = formatTimeCell_(r[3]);

      if (!title || !channel) continue;
      if (channel === 'ช่อง' || title === 'ชื่อรายการ' || date === 'วันที่') continue;

      out.push({
        row: rowNum,
        date: date,
        channel: channel,
        title: title,
        scheduled_time: scheduledTime,
        platforms: {
          facebook: platformStat_(r[4], r[8], r[9]),
          youtube: platformStat_(r[5], r[10], r[11]),
          tiktok: platformStat_(r[6], r[12], r[13]),
          x: platformStat_(r[7], r[14], r[15])
        }
      });
    }

    return {
      ok: true,
      sheet: 'View Stats',
      total: out.length,
      generated_at: new Date().toISOString(),
      data: out
    };
  } catch (err) {
    return { ok: false, error: String(err && err.stack ? err.stack : err) };
  }
}

function platformStat_(link, peakTime, peakView) {
  var v = parseViewCount_(peakView);
  return {
    link: isValidUrl_(link) ? String(link).trim() : '-',
    peak_time: formatTimeCell_(peakTime),
    peak_view: v >= 0 ? v : null
  };
}

function formatTimeCell_(val) {
  if (val === null || val === undefined || val === '') return '-';
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone() || 'Asia/Bangkok', 'HH:mm:ss');
  }
  var s = String(val).trim();
  return s === '' ? '-' : s;
}