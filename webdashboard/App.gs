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
 *   1.5 append_programs (sync แบบ append-only: เพิ่มเฉพาะรายการใหม่ ต่อท้าย ไม่แตะของเดิม)
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
  'หมวดหมู่',
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
    // [ADDED] Live View Stats dashboard endpoint — อ่านชีท 'View Stats' ทั้ง 15 คอลัมน์ (อ่านอย่างเดียว)
    // รองรับ JSONP: ส่ง ?callback=fnName เพื่อเลี่ยงปัญหา CORS เมื่อเปิดไฟล์แบบ file://
    if (e && e.parameter && (e.parameter.action === 'view_stats' || e.parameter.action === 'get_view_stats')) {
      var vsPayload = buildViewStatsPayload_({ offset: e.parameter.offset, limit: e.parameter.limit });
      var vsCb = e.parameter.callback;
      if (vsCb) {
        return ContentService
          .createTextOutput(vsCb + '(' + JSON.stringify(vsPayload) + ');')
          .setMimeType(ContentService.MimeType.JAVASCRIPT);
      }
      return jsonOut(vsPayload);
    }

    // [ADDED] Link Config endpoint — อ่าน Channel_Links และ Broadcast_Overrides
    if (e && e.parameter && (e.parameter.action === 'get_link_config' || e.parameter.action === 'link_config')) {
      var includeBc = (e.parameter.quick !== '1' && e.parameter.include_broadcasts !== '0');
      var cfgPayload = getLinkConfig_(SpreadsheetApp.getActiveSpreadsheet(), { includeBroadcasts: includeBc });
      var cfgCb = e.parameter.callback;
      if (cfgCb) {
        return ContentService
          .createTextOutput(cfgCb + '(' + JSON.stringify(cfgPayload) + ');')
          .setMimeType(ContentService.MimeType.JAVASCRIPT);
      }
      return jsonOut(cfgPayload);
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

    // หากไม่ระบุ ให้ดึงข้อมูลผังรายการของทุกช่อง (ยกเว้นชีท 'View Stats', Config, Overrides)
    var allSheets = ss.getSheets();
    var channelData = {};
    var totalCount = 0;

    allSheets.forEach(function (sh) {
      var sName = sh.getName();
      if (!isScheduleSheetName_(sName)) return;
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
      return jsonOut(buildViewStatsPayload_({ offset: req.offset, limit: req.limit }));
    }

    // [ADDED] Link Config endpoints (อ่านและบันทึก Channel_Links และ Broadcast_Overrides)
    if (action === 'get_link_config' || action === 'link_config') {
      var includeBc = (req.quick !== true && req.quick !== '1' && req.include_broadcasts !== false && req.include_broadcasts !== '0');
      return jsonOut(getLinkConfig_(ss, { includeBroadcasts: includeBc }));
    }
    if (action === 'save_link_config') {
      result = saveLinkConfig_(ss, req);
      return jsonOut({ ok: true, action: action, result: result });
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
        if (!isScheduleSheetName_(sName)) return;
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

    // 6.5 append-only sync: เพิ่มเฉพาะรายการที่ยังไม่มีในชีท ต่อท้ายตาราง (record เดิมไม่ถูกแตะ)
    if (action === 'append_programs' || action === 'append_new' || action === 'sync_programs') {
      result = appendPrograms_(ss, req.sheet, req.data);
      return jsonOut({ ok: true, action: action, result: result });
    }

    // 6.6 purge: ลบแถวผังรายการที่ "วันเก่ากว่า" cutoff แล้วเลื่อนข้อมูลที่เหลือขึ้นแทนช่องว่าง
    //     (หัวตารางแถว 1 ไม่ถูกแตะ ; คอลัมน์ลิงก์ D:G เลื่อนตามแถวไปด้วย)
    if (action === 'purge_before' || action === 'purge_old' || action === 'delete_before') {
      result = purgeProgramsBefore_(ss, req.sheet, req.cutoff || req.date || req.before, req);
      return jsonOut({ ok: true, action: action, result: result });
    }

    // 7. เขียนข้อมูลลงช่วงที่กำหนดโดยตรง (เช่น Backfill คอลัมน์หมวดหมู่)
    if (action === 'set_range') {
      var targetSh = req.sheet ? ss.getSheetByName(req.sheet) : ss.getActiveSheet();
      if (!targetSh) throw new Error('sheet not found: ' + req.sheet);
      targetSh.getRange(req.range).setValues(req.values);
      return jsonOut({ ok: true, action: action, range: req.range, rows: req.values.length });
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

/**
 * เพิ่มเฉพาะรายการที่ยังไม่มีในชีท ต่อท้ายตาราง (append-only, forward-only)
 *
 * - เทียบด้วยคีย์ normalize "date|time|title" กับรายการเดิมในคอลัมน์ A:C
 *   (รองรับวันที่ทั้งรูปแบบ DD-MM-YY, DD-MM-YYYY, YYYY-MM-DD และ Date object)
 * - รายการที่มีอยู่แล้วจะถูกข้าม, ของเดิมในชีท "ไม่ถูกแตะ" เลยสักเซลล์
 * - "forward-only": หาวันที่ล่าสุดที่มีในชีท (maxDate) แล้วเพิ่มเฉพาะรายการที่วัน >= maxDate
 *   -> รายการที่วัน "เก่ากว่า" วันล่าสุดในชีทจะถูกข้ามเสมอ แม้จะยังไม่มีในชีท (ไม่ backfill ย้อนหลัง)
 *   -> วันเดียวกับ maxDate ที่ยังไม่มีในชีทยังเพิ่มได้ (เติมช่องว่างของวันล่าสุด)
 * - รายการใหม่ต่อท้ายใต้แถวสุดท้าย เขียนเฉพาะ A:C (ไม่ยุ่งกับคอลัมน์ลิงก์ D:G)
 *
 * req.data = [[วัน, เวลา, รายการ], ...]
 */
function appendPrograms_(ss, name, data) {
  requireName(name);
  var sh = ss.getSheetByName(name);
  var createdSheet = false;
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sh.setFrozenRows(1);
    sh.getRange('A:B').setNumberFormat('@');
    createdSheet = true;
  }

  data = (data && data.length) ? data : [];

  // 1. เก็บคีย์ของรายการเดิมทั้งหมด + หาวันที่ล่าสุดที่มีในชีท
  var seen = {};
  var lastRow = sh.getLastRow();
  var existingCount = 0;
  var maxDate = '';  // 'YYYY-MM-DD' ของ record ที่ใหม่ที่สุดในชีท ('' = ชีทยังว่าง)
  if (lastRow >= START_ROW) {
    var existing = sh.getRange(START_ROW, 1, lastRow - START_ROW + 1, PROGRAM_COLS).getValues();
    for (var i = 0; i < existing.length; i++) {
      var ek = progKey_(existing[i][0], existing[i][1], existing[i][2]);
      if (ek !== '||') { seen[ek] = true; existingCount++; }
      var ed = normProgDate_(existing[i][0]);
      if (/^\d{4}-\d{2}-\d{2}$/.test(ed) && ed > maxDate) maxDate = ed;
    }
  }

  // 2. กรอง: เพิ่มเฉพาะแถวที่ (ก) ยังไม่มีในชีท และ (ข) วันไม่เก่ากว่าวันล่าสุดในชีท
  var toAppend = [];
  var skipped = 0;
  var skippedOld = 0;
  for (var j = 0; j < data.length; j++) {
    var row = data[j] || [];
    var d = row[0] == null ? '' : String(row[0]).trim();
    var t = row[1] == null ? '' : String(row[1]).trim();
    var title = row[2] == null ? '' : String(row[2]).trim();
    if (!d && !t && !title) continue;

    var k = progKey_(d, t, title);
    if (seen[k]) { skipped++; continue; }

    if (maxDate) {
      var dn = normProgDate_(d);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(dn) || dn < maxDate) { skippedOld++; continue; }
    }

    seen[k] = true;
    toAppend.push([d, t, title]);
  }

  // 3. ต่อท้ายใต้แถวสุดท้าย — ไม่ทับของเดิม
  if (toAppend.length > 0) {
    var startRow = Math.max(sh.getLastRow() + 1, START_ROW);
    sh.getRange(startRow, 1, toAppend.length, PROGRAM_COLS)
      .setNumberFormat('@')
      .setValues(toAppend);
  }

  return {
    sheet: name,
    created: createdSheet,
    latest_date_in_sheet: maxDate,
    skipped_old: skippedOld,
    existing_before: existingCount,
    appended: toAppend.length,
    skipped: skipped,
    total_rows: existingCount + toAppend.length
  };
}

/**
 * ลบแถวผังรายการที่ "วันเก่ากว่า" cutoff (ไม่รวมวันเท่ากับ cutoff)
 * แล้วเลื่อนข้อมูลที่เหลือขึ้นไปแทนช่องว่างจนชิดหัวตาราง
 *
 * - cutoff: string วันที่รูปแบบใดก็ได้ที่ normProgDate_ เข้าใจ
 *   (DD-MM-YY, DD-MM-YYYY, DD/MM/YYYY, YYYY-MM-DD) -> normalize เป็น 'YYYY-MM-DD'
 * - เทียบแบบพจนานุกรมบน 'YYYY-MM-DD' : แถวที่ dNorm < cutoffNorm จะถูกลบ
 * - แถวที่ parse วันไม่ได้ (ว่าง/ผิดรูปแบบ) จะถูก "เก็บไว้" เพื่อความปลอดภัยของข้อมูล
 * - เลื่อนทั้งแถว (คอลัมน์ A จนถึงคอลัมน์สุดท้ายที่มีข้อมูล) ขึ้น ทำให้ลิงก์ D:G ติดไปกับแถวเดิม
 * - หัวตารางแถวที่ 1 ไม่ถูกแตะ ; แถวส่วนเกินท้ายตารางถูก deleteRows ทิ้งจริง (ชีทหดตาม)
 * - req.dry_run === true : รายงานจำนวนที่จะลบโดยไม่แก้ไขชีท
 */
function purgeProgramsBefore_(ss, name, cutoff, req) {
  requireName(name);
  var sh = ss.getSheetByName(name);
  if (!sh) throw new Error('sheet not found: ' + name);

  var cutoffNorm = normProgDate_(cutoff);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(cutoffNorm)) {
    throw new Error('bad cutoff date: ' + cutoff + ' (ต้องเป็น DD-MM-YYYY หรือ YYYY-MM-DD)');
  }

  var dryRun = !!(req && (req.dry_run === true || req.dry_run === '1' || req.dryRun === true));

  var lastRow = sh.getLastRow();
  var lastCol = Math.max(sh.getLastColumn(), HEADERS.length);
  if (lastRow < START_ROW) {
    return { sheet: name, cutoff: cutoffNorm, removed: 0, kept: 0, total_rows: 0, removed_sample: [], dry_run: dryRun };
  }

  var numRows = lastRow - START_ROW + 1;
  var values = sh.getRange(START_ROW, 1, numRows, lastCol).getValues();

  var keep = [];
  var removed = 0;
  var removedSample = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var dNorm = normProgDate_(row[0]);
    var isDate = /^\d{4}-\d{2}-\d{2}$/.test(dNorm);
    if (isDate && dNorm < cutoffNorm) {
      removed++;
      if (removedSample.length < 5) {
        removedSample.push([String(row[0]), String(row[1]), String(row[2])]);
      }
      continue;
    }
    keep.push(row);
  }

  if (dryRun || removed === 0) {
    return {
      sheet: name, cutoff: cutoffNorm,
      removed: removed, kept: keep.length, total_rows: keep.length,
      removed_sample: removedSample, dry_run: dryRun
    };
  }

  // เขียนแถวที่เก็บไว้กลับ เริ่มที่แถว 2 (ชิดหัวตาราง) แล้วลบแถวส่วนเกินท้ายตารางทิ้ง
  if (keep.length > 0) {
    sh.getRange(START_ROW, 1, keep.length, lastCol).setValues(keep);
  }
  var firstEmptyRow = START_ROW + keep.length;
  var trailing = lastRow - firstEmptyRow + 1;
  if (trailing > 0) {
    sh.deleteRows(firstEmptyRow, trailing);
  }
  // คงรูปแบบข้อความคอลัมน์ วัน/เวลา กัน Sheets แปลงเป็นวันที่/เวลาอัตโนมัติ
  sh.getRange('A:B').setNumberFormat('@');

  return {
    sheet: name, cutoff: cutoffNorm,
    removed: removed, kept: keep.length, total_rows: keep.length,
    removed_sample: removedSample, dry_run: false
  };
}

/** คีย์เทียบรายการ: "YYYY-MM-DD|HH:mm|ชื่อรายการ(lowercase, ยุบช่องว่าง)" */
function progKey_(d, t, title) {
  return normProgDate_(d) + '|' + normProgTime_(t) + '|' + normProgTitle_(title);
}

function normProgDate_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone() || 'Asia/Bangkok', 'yyyy-MM-dd');
  }
  var s = String(v == null ? '' : v).trim();
  if (!s) return '';
  var m = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{2,4})$/);   // DD-MM-YY / DD-MM-YYYY / DD/MM/YYYY
  if (m) {
    var dd = ('0' + m[1]).slice(-2);
    var mm = ('0' + m[2]).slice(-2);
    var yy = m[3].length === 2 ? '20' + m[3] : m[3];
    return yy + '-' + mm + '-' + dd;
  }
  var m2 = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/);      // YYYY-MM-DD
  if (m2) {
    return m2[1] + '-' + ('0' + m2[2]).slice(-2) + '-' + ('0' + m2[3]).slice(-2);
  }
  return s.toLowerCase();
}

function normProgTime_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone() || 'Asia/Bangkok', 'HH:mm');
  }
  var s = String(v == null ? '' : v).trim();
  var m = s.match(/^(\d{1,2}):(\d{2})/);
  if (m) return ('0' + m[1]).slice(-2) + ':' + m[2];
  return s;
}

function normProgTitle_(v) {
  return String(v == null ? '' : v).trim().replace(/\s+/g, ' ').toLowerCase();
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
    sh.getRange('E:E').setNumberFormat('@');
    sh.getRange('J:J').setNumberFormat('@');
    sh.getRange('L:L').setNumberFormat('@');
    sh.getRange('N:N').setNumberFormat('@');
    sh.getRange('P:P').setNumberFormat('@');
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
    var rGenre = String(r.genre || r.category || r.หมวดหมู่ || '').trim();
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

      // อัปเดตหมวดหมู่หากมีข้อมูลและในชีทยังว่างหรือเป็น '-'
      if (rGenre && (!targetRow[3] || targetRow[3] === '-')) {
        targetRow[3] = rGenre;
        modified = true;
      }

      // อัปเดตเวลาเริ่มในผังหากมีข้อมูลและในชีทยังว่างหรือเป็น '-'
      if (rScheduledTime && (!targetRow[4] || targetRow[4] === '-')) {
        targetRow[4] = rScheduledTime;
        modified = true;
      }

      // อัปเดตลิงก์หากได้ลิงก์สดที่ถูกต้องมาใหม่
      // Facebook: เฉพาะลิงก์วิดีโอ/ไลฟ์จริงเท่านั้น (ป้องกัน URL หน้าช่อง เช่น /watch/ThaiPBS/)
      if (isFacebookVideoUrl_(fbLink)) {
        if (!isFacebookVideoUrl_(targetRow[5])) {
          targetRow[5] = fbLink;
          modified = true;
        }
      } else if (!isFacebookVideoUrl_(targetRow[5]) && targetRow[5] !== '-') {
        // หากในชีทเคยบันทึกเป็น URL ช่องที่ไม่ใช่วิดีโอ ให้ล้างเป็น '-'
        targetRow[5] = '-';
        modified = true;
      }

      // YouTube: เฉพาะลิงก์วิดีโอ/สตรีมจริงเท่านั้น (ป้องกัน URL หน้าช่อง เช่น /streams)
      if (isYouTubeVideoUrl_(ytLink)) {
        if (!isYouTubeVideoUrl_(targetRow[6])) {
          targetRow[6] = ytLink;
          modified = true;
        }
      } else if (!isYouTubeVideoUrl_(targetRow[6]) && targetRow[6] !== '-') {
        // หากในชีทเคยบันทึกเป็น URL หน้าช่องที่ไม่ใช่วิดีโอ ให้ล้างเป็น '-'
        targetRow[6] = '-';
        modified = true;
      }

      if (isValidUrl_(ttLink) && (!isValidUrl_(targetRow[7]) || targetRow[7] === '-')) {
        targetRow[7] = ttLink;
        modified = true;
      }
      if (isValidUrl_(xLink) && (!isValidUrl_(targetRow[8]) || targetRow[8] === '-')) {
        targetRow[8] = xLink;
        modified = true;
      }

      // 1. Facebook: เปรียบเทียบยอดวิวพีค (Col J: Peak Time, Col K: Peak View)
      var curFbPeak = parseViewCount_(targetRow[10]);
      if (fbViews >= 0 && (curFbPeak < 0 || fbViews > curFbPeak)) {
        targetRow[9] = rTime || targetRow[9] || '-';
        targetRow[10] = fbViews;
        modified = true;
      }

      // 2. YouTube: เปรียบเทียบยอดวิวพีค (Col L: Peak Time, Col M: Peak View)
      var curYtPeak = parseViewCount_(targetRow[12]);
      if (ytViews >= 0 && (curYtPeak < 0 || ytViews > curYtPeak)) {
        targetRow[11] = rTime || targetRow[11] || '-';
        targetRow[12] = ytViews;
        modified = true;
      }

      // 3. TikTok: เปรียบเทียบยอดวิวพีค (Col N: Peak Time, Col O: Peak View)
      var curTtPeak = parseViewCount_(targetRow[14]);
      if (ttViews >= 0 && (curTtPeak < 0 || ttViews > curTtPeak)) {
        targetRow[13] = rTime || targetRow[13] || '-';
        targetRow[14] = ttViews;
        modified = true;
      }

      // 4. X (Twitter): เปรียบเทียบยอดวิวพีค (Col P: Peak Time, Col Q: Peak View)
      var curXPeak = parseViewCount_(targetRow[16]);
      if (xViews >= 0 && (curXPeak < 0 || xViews > curXPeak)) {
        targetRow[15] = rTime || targetRow[15] || '-';
        targetRow[16] = xViews;
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
        rGenre || '-',
        rScheduledTime || '-',
        isFacebookVideoUrl_(fbLink) ? fbLink : '-',
        isYouTubeVideoUrl_(ytLink) ? ytLink : '-',
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
    sh.getRange(2, 5, existingRows.length, 1).setNumberFormat('@');
    sh.getRange(2, 10, existingRows.length, 1).setNumberFormat('@');
    sh.getRange(2, 12, existingRows.length, 1).setNumberFormat('@');
    sh.getRange(2, 14, existingRows.length, 1).setNumberFormat('@');
    sh.getRange(2, 16, existingRows.length, 1).setNumberFormat('@');
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

function isFacebookVideoUrl_(url) {
  if (!isValidUrl_(url)) return false;
  var s = url.trim();
  if (/^https?:\/\/(?:www\.|m\.)?fb\.watch\/[A-Za-z0-9_-]+/i.test(s)) return true;
  if (/[?&]v=\d+/.test(s)) return true;
  if (/\/videos\/(?:[^\/?#]+\/)?\d+/.test(s)) return true;
  if (/\/live\/(?:videos\/)?\d+/.test(s)) return true;
  if (/video\.php.*[?&]v=\d+/.test(s)) return true;
  return false;
}

function isYouTubeVideoUrl_(url) {
  if (!isValidUrl_(url)) return false;
  var s = url.trim();
  if (s.indexOf('watch?v=') > -1) return true;
  if (/^https?:\/\/youtu\.be\/[A-Za-z0-9_-]+/i.test(s)) return true;
  if (/youtube\.com\/(?:live|embed|v)\/[A-Za-z0-9_-]+/i.test(s)) return true;
  return false;
}

/* ============================================================================
 * [ADDED] Live View Stats Dashboard — อ่านชีท 'View Stats' (หน้า Dashboard ที่ 2)
 * เป็นการอ่านอย่างเดียว ไม่แก้ไข/ไม่แตะโค้ดหรือชีทเดิม
 *
 * โครงสร้างชีท 'View Stats':
 *   A วันที่ | B ช่อง | C ชื่อรายการ | D หมวดหมู่ | E เวลาเริ่มในผัง
 *   F Facebook Link | G YouTube Link | H TikTok Link | I X Link
 *   J Facebook Peak time | K Facebook Views
 *   L YouTube Peak time  | M YouTube Views
 *   N TikTok Peak time   | O TikTok Views
 *   P X Peak time        | Q X Views
 *
 * เรียกใช้:  GET  <exec>?action=view_stats
 *           POST <exec>  body: {"action":"view_stats"}
 *
 * รองรับแบ่งหน้า (กันเกินลิมิตขนาด response เมื่อชีทมีหลายแสนแถว):
 *   ?action=view_stats&offset=0&limit=25000   -> คืน 25000 แถวถัดจาก offset + has_more
 *   ไม่ใส่ offset/limit = คืนทั้งหมด (พฤติกรรมเดิม, has_more=false)
 * ========================================================================== */

function buildViewStatsPayload_(opts) {
  opts = opts || {};
  var offset = parseInt(opts.offset, 10);
  if (!(offset > 0)) offset = 0;
  var limit = parseInt(opts.limit, 10);
  if (!(limit > 0)) limit = 0;   // 0 = ไม่จำกัด (คืนทั้งหมด)

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('View Stats');
    if (!sh) return { ok: false, error: 'sheet not found: View Stats' };

    var lastRow = sh.getLastRow();
    var dataRows = Math.max(0, lastRow - 1);   // ไม่รวมหัวตารางแถว 1
    var base = {
      ok: true, sheet: 'View Stats', total: dataRows,
      offset: offset, limit: limit, generated_at: new Date().toISOString()
    };
    if (dataRows === 0) {
      base.has_more = false; base.data = []; return base;
    }

    var startRow = 2 + offset;                 // แถวจริงในชีท
    if (startRow > lastRow) {
      base.has_more = false; base.data = []; return base;
    }
    var avail = lastRow - startRow + 1;
    var count = (limit > 0) ? Math.min(limit, avail) : avail;

    var values = sh.getRange(startRow, 1, count, 17).getValues();
    var out = [];

    for (var i = 0; i < values.length; i++) {
      var r = values[i];
      var date = normalizeDateStr_(r[0]);
      var channel = String(r[1] || '').trim();
      var title = String(r[2] || '').trim();
      var genre = String(r[3] || '').trim();
      var scheduledTime = formatTimeCell_(r[4]);

      if (!title || !channel) continue;
      if (channel === 'ช่อง' || title === 'ชื่อรายการ' || date === 'วันที่') continue;

      out.push({
        row: startRow + i,
        date: date,
        channel: channel,
        title: title,
        genre: genre,
        scheduled_time: scheduledTime,
        platforms: {
          facebook: platformStat_(isFacebookVideoUrl_(r[5]) ? r[5] : '-', r[9], r[10]),
          youtube: platformStat_(isYouTubeVideoUrl_(r[6]) ? r[6] : '-', r[11], r[12]),
          tiktok: platformStat_(r[7], r[13], r[14]),
          x: platformStat_(r[8], r[15], r[16])
        }
      });
    }

    base.has_more = (limit > 0) && (startRow + count - 1 < lastRow);
    base.count = out.length;
    base.data = out;
    return base;
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

/* ============================================================================
 * [ADDED] Link Config Management — Channel_Links & Broadcast_Overrides
 * จัดเก็บลิงก์หลักประจำช่อง (รองรับหลายลิงก์) และลิงก์เฉพาะรายรายการ (Overrides)
 * ========================================================================== */

var CHANNEL_LINKS_SHEET = 'Channel_Links';
var BROADCAST_OVERRIDES_SHEET = 'Broadcast_Overrides';

var CHANNEL_LINKS_HEADERS = ['Channel', 'Facebook URLs', 'YouTube URLs', 'TikTok URLs', 'X URLs', 'Last Updated'];
var BROADCAST_OVERRIDES_HEADERS = ['Channel', 'Program Title', 'Facebook URLs', 'YouTube URLs', 'TikTok URLs', 'X URLs', 'Last Updated'];

function isScheduleSheetName_(sName) {
  var n = String(sName || '').trim().toLowerCase();
  if (!n) return false;
  if (n.indexOf('view stats') === 0) return false;
  if (n === 'channel_links' || n === 'broadcast_overrides') return false;
  if (n.indexOf('config') >= 0 || n.indexOf('override') >= 0) return false;
  return true;
}

function getChannelBroadcastsCached_(ss) {
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get('channel_broadcasts_v1');
    if (cached) {
      return JSON.parse(cached);
    }
  } catch (e) {}

  var channelBroadcasts = {};
  var allSheets = ss.getSheets();
  allSheets.forEach(function (sh) {
    var sName = sh.getName();
    if (!isScheduleSheetName_(sName)) return;
    var lastRow = sh.getLastRow();
    if (lastRow < START_ROW) return;
    var titlesCol = sh.getRange(START_ROW, 3, lastRow - START_ROW + 1, 1).getValues();
    var titles = [];
    var seenTitles = {};
    for (var i = 0; i < titlesCol.length; i++) {
      var t = String(titlesCol[i][0] || '').trim();
      if (!t || t === '-' || t.toLowerCase() === 'n/a' || t === 'รายการ' || t.indexOf('http') === 0) continue;
      if (!seenTitles[t]) {
        seenTitles[t] = true;
        titles.push(t);
      }
    }
    if (titles.length > 0) {
      channelBroadcasts[sName] = titles;
    }
  });

  try {
    var cache = CacheService.getScriptCache();
    var str = JSON.stringify(channelBroadcasts);
    if (str.length < 95000) {
      cache.put('channel_broadcasts_v1', str, 1800); // Cache for 30 mins
    }
  } catch (e) {}

  return channelBroadcasts;
}

function getLinkConfig_(ss, options) {
  options = options || {};
  var includeBroadcasts = options.includeBroadcasts !== false;

  try {
    var channelLinks = {};
    var chSheet = ss.getSheetByName(CHANNEL_LINKS_SHEET);
    if (!chSheet) {
      chSheet = ss.insertSheet(CHANNEL_LINKS_SHEET);
      chSheet.getRange(1, 1, 1, CHANNEL_LINKS_HEADERS.length).setValues([CHANNEL_LINKS_HEADERS]);
      chSheet.setFrozenRows(1);
    } else if (chSheet.getLastRow() >= 2) {
      var numRows = chSheet.getLastRow() - 1;
      var vals = chSheet.getRange(2, 1, numRows, 6).getValues();
      vals.forEach(function(r) {
        var ch = String(r[0] || '').trim();
        if (!ch) return;
        channelLinks[ch] = {
          facebook: parseUrlList_(r[1]),
          youtube: parseUrlList_(r[2]),
          tiktok: parseUrlList_(r[3]),
          x: parseUrlList_(r[4]),
          last_updated: r[5] ? String(r[5]) : ''
        };
      });
    }

    var broadcastOverrides = [];
    var boSheet = ss.getSheetByName(BROADCAST_OVERRIDES_SHEET);
    if (!boSheet) {
      boSheet = ss.insertSheet(BROADCAST_OVERRIDES_SHEET);
      boSheet.getRange(1, 1, 1, BROADCAST_OVERRIDES_HEADERS.length).setValues([BROADCAST_OVERRIDES_HEADERS]);
      boSheet.setFrozenRows(1);
    } else if (boSheet.getLastRow() >= 2) {
      var numBoRows = boSheet.getLastRow() - 1;
      var boVals = boSheet.getRange(2, 1, numBoRows, 7).getValues();
      boVals.forEach(function(r) {
        var ch = String(r[0] || '').trim();
        var title = String(r[1] || '').trim();
        if (!ch || !title) return;
        broadcastOverrides.push({
          channel: ch,
          title: title,
          facebook: parseUrlList_(r[2]),
          youtube: parseUrlList_(r[3]),
          tiktok: parseUrlList_(r[4]),
          x: parseUrlList_(r[5]),
          last_updated: r[6] ? String(r[6]) : ''
        });
      });
    }

    var res = {
      ok: true,
      channel_links: channelLinks,
      broadcast_overrides: broadcastOverrides
    };

    if (includeBroadcasts) {
      res.channel_broadcasts = getChannelBroadcastsCached_(ss);
    }

    return res;
  } catch (err) {
    return { ok: false, error: String(err && err.stack ? err.stack : err) };
  }
}

function normalizeUrlForDedupe_(url) {
  if (!url) return '';
  var s = String(url).trim();
  return s.toLowerCase().replace(/^https?:\/\//, '').replace(/\/+$/, '');
}

function dedupeUrls_(urlList) {
  if (!urlList) return [];
  var arr = Array.isArray(urlList) ? urlList : String(urlList).split(/[\r\n,]+/);
  var out = [];
  var seen = {};
  for (var i = 0; i < arr.length; i++) {
    var u = String(arr[i] || '').trim();
    if (!u || u === '-' || u.toUpperCase() === 'N/A') continue;
    var norm = normalizeUrlForDedupe_(u);
    if (!seen[norm]) {
      seen[norm] = true;
      out.push(u);
    }
  }
  return out;
}

function parseUrlList_(val) {
  return dedupeUrls_(val);
}

function mergeUrlLists3WayServer_(baseList, sheetList, incomingList) {
  var base = dedupeUrls_(baseList);
  var sheet = dedupeUrls_(sheetList);
  var incoming = dedupeUrls_(incomingList);

  var baseMap = {};
  for (var i = 0; i < base.length; i++) baseMap[normalizeUrlForDedupe_(base[i])] = true;

  var incomingMap = {};
  for (var i = 0; i < incoming.length; i++) incomingMap[normalizeUrlForDedupe_(incoming[i])] = true;

  var userAdded = [];
  for (var i = 0; i < incoming.length; i++) {
    if (!baseMap[normalizeUrlForDedupe_(incoming[i])]) {
      userAdded.push(incoming[i]);
    }
  }

  var userDeletedMap = {};
  for (var i = 0; i < base.length; i++) {
    if (!incomingMap[normalizeUrlForDedupe_(base[i])]) {
      userDeletedMap[normalizeUrlForDedupe_(base[i])] = true;
    }
  }

  var result = [];
  var seen = {};

  // 1. Keep URLs currently in sheet (unless explicitly deleted by this request)
  for (var i = 0; i < sheet.length; i++) {
    var norm = normalizeUrlForDedupe_(sheet[i]);
    if (!userDeletedMap[norm]) {
      if (!seen[norm]) {
        seen[norm] = true;
        result.push(sheet[i]);
      }
    }
  }

  // 2. Add URLs added by this request
  for (var i = 0; i < userAdded.length; i++) {
    var norm = normalizeUrlForDedupe_(userAdded[i]);
    if (!seen[norm]) {
      seen[norm] = true;
      result.push(userAdded[i]);
    }
  }

  return result;
}

function areOverridesEqualServer_(bo1, bo2) {
  if (!bo1 && !bo2) return true;
  if (!bo1 || !bo2) return false;
  var plats = ['facebook', 'youtube', 'tiktok', 'x'];
  for (var i = 0; i < plats.length; i++) {
    var p = plats[i];
    var l1 = dedupeUrls_(bo1[p]).map(normalizeUrlForDedupe_).sort().join('\n');
    var l2 = dedupeUrls_(bo2[p]).map(normalizeUrlForDedupe_).sort().join('\n');
    if (l1 !== l2) return false;
  }
  return true;
}

function saveLinkConfig_(ss, req) {
  var lock = LockService.getScriptLock();
  var hasLock = lock.tryLock(30000);
  if (!hasLock) {
    return {
      success: false,
      ok: false,
      error: 'ระบบกำลังประมวลผลคำขอบันทึกอื่น กรุณาลองใหม่อีกครั้งในครู่เดียว'
    };
  }

  try {
    var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
    var updatedChannels = 0;
    var updatedOverrides = 0;

    // Read current sheet state while holding the lock (captures changes saved mere milliseconds ago)
    var currentSheetConfig = getLinkConfig_(ss, { includeBroadcasts: false });
    var sheetChMap = currentSheetConfig.channel_links || {};
    var sheetBoList = currentSheetConfig.broadcast_overrides || [];

    // 1. Channel Links
    if (req.channel_links && typeof req.channel_links === 'object') {
      var chSheet = ss.getSheetByName(CHANNEL_LINKS_SHEET);
      if (!chSheet) {
        chSheet = ss.insertSheet(CHANNEL_LINKS_SHEET);
        chSheet.getRange(1, 1, 1, CHANNEL_LINKS_HEADERS.length).setValues([CHANNEL_LINKS_HEADERS]);
        chSheet.setFrozenRows(1);
      }

      var baseChMap = (req.base_channel_links && typeof req.base_channel_links === 'object') ? req.base_channel_links : {};
      var allChannelNames = {};
      Object.keys(req.channel_links).forEach(function(k) { allChannelNames[k] = true; });
      Object.keys(sheetChMap).forEach(function(k) { allChannelNames[k] = true; });

      var chRows = [];
      var chKeys = Object.keys(allChannelNames).sort();
      chKeys.forEach(function(ch) {
        var baseItem = baseChMap[ch] || {};
        var sheetItem = sheetChMap[ch] || {};
        var incItem = req.channel_links[ch] || sheetItem;

        var fb = mergeUrlLists3WayServer_(baseItem.facebook, sheetItem.facebook, incItem.facebook).join('\n');
        var yt = mergeUrlLists3WayServer_(baseItem.youtube, sheetItem.youtube, incItem.youtube).join('\n');
        var tt = mergeUrlLists3WayServer_(baseItem.tiktok, sheetItem.tiktok, incItem.tiktok).join('\n');
        var x = mergeUrlLists3WayServer_(baseItem.x, sheetItem.x, incItem.x).join('\n');
        chRows.push([ch, fb, yt, tt, x, nowStr]);
      });

      if (chSheet.getLastRow() > 1) {
        chSheet.getRange(2, 1, chSheet.getLastRow() - 1, chSheet.getLastColumn()).clearContent();
      }
      if (chRows.length > 0) {
        chSheet.getRange(2, 1, chRows.length, 6).setNumberFormat('@').setValues(chRows);
      }
      updatedChannels = chRows.length;
    }

    // 2. Broadcast Overrides
    if (req.broadcast_overrides && Array.isArray(req.broadcast_overrides)) {
      var boSheet = ss.getSheetByName(BROADCAST_OVERRIDES_SHEET);
      if (!boSheet) {
        boSheet = ss.insertSheet(BROADCAST_OVERRIDES_SHEET);
        boSheet.getRange(1, 1, 1, BROADCAST_OVERRIDES_HEADERS.length).setValues([BROADCAST_OVERRIDES_HEADERS]);
        boSheet.setFrozenRows(1);
      }

      var getBoKey = function(b) {
        return String(b.channel || '').trim() + ':::' + String(b.title || '').trim().toLowerCase();
      };

      var baseBoMap = {};
      if (Array.isArray(req.base_broadcast_overrides)) {
        req.base_broadcast_overrides.forEach(function(b) {
          var k = getBoKey(b);
          if (k !== ':::') baseBoMap[k] = b;
        });
      }

      var sheetBoMap = {};
      sheetBoList.forEach(function(b) {
        var k = getBoKey(b);
        if (k !== ':::') sheetBoMap[k] = b;
      });

      var incBoMap = {};
      req.broadcast_overrides.forEach(function(b) {
        var k = getBoKey(b);
        if (k !== ':::') incBoMap[k] = b;
      });

      // Overrides explicitly deleted by this incoming request
      var userDeletedOverrides = {};
      Object.keys(baseBoMap).forEach(function(k) {
        if (!incBoMap[k]) userDeletedOverrides[k] = true;
      });

      var boRows = [];
      var handledKeys = {};

      // Preserve existing sheet overrides (unless explicitly deleted by this request)
      Object.keys(sheetBoMap).forEach(function(k) {
        if (userDeletedOverrides[k]) return;
        var sBo = sheetBoMap[k];
        if (!incBoMap[k]) {
          boRows.push([
            sBo.channel,
            sBo.title,
            dedupeUrls_(sBo.facebook).join('\n'),
            dedupeUrls_(sBo.youtube).join('\n'),
            dedupeUrls_(sBo.tiktok).join('\n'),
            dedupeUrls_(sBo.x).join('\n'),
            nowStr
          ]);
        } else {
          var bBo = baseBoMap[k] || {};
          var iBo = incBoMap[k];
          var fb = mergeUrlLists3WayServer_(bBo.facebook, sBo.facebook, iBo.facebook).join('\n');
          var yt = mergeUrlLists3WayServer_(bBo.youtube, sBo.youtube, iBo.youtube).join('\n');
          var tt = mergeUrlLists3WayServer_(bBo.tiktok, sBo.tiktok, iBo.tiktok).join('\n');
          var x = mergeUrlLists3WayServer_(bBo.x, sBo.x, iBo.x).join('\n');
          if (fb || yt || tt || x) {
            boRows.push([iBo.channel || sBo.channel, iBo.title || sBo.title, fb, yt, tt, x, nowStr]);
          }
        }
        handledKeys[k] = true;
      });

      // Newly added overrides by this request
      Object.keys(incBoMap).forEach(function(k) {
        if (!handledKeys[k]) {
          var iBo = incBoMap[k];
          var bBo = baseBoMap[k];
          // If this override was in base, but missing from sheet, another user deleted it from the sheet!
          // If incoming request made no new edits compared to base, respect the server deletion and DO NOT resurrect!
          if (bBo && areOverridesEqualServer_(bBo, iBo)) {
            return;
          }
          var fb = dedupeUrls_(iBo.facebook).join('\n');
          var yt = dedupeUrls_(iBo.youtube).join('\n');
          var tt = dedupeUrls_(iBo.tiktok).join('\n');
          var x = dedupeUrls_(iBo.x).join('\n');
          if (fb || yt || tt || x) {
            boRows.push([
              iBo.channel,
              iBo.title,
              fb,
              yt,
              tt,
              x,
              nowStr
            ]);
          }
          handledKeys[k] = true;
        }
      });

      if (boSheet.getLastRow() > 1) {
        boSheet.getRange(2, 1, boSheet.getLastRow() - 1, boSheet.getLastColumn()).clearContent();
      }
      if (boRows.length > 0) {
        boSheet.getRange(2, 1, boRows.length, 7).setNumberFormat('@').setValues(boRows);
      }
      updatedOverrides = boRows.length;
    }

    var finalConfig = getLinkConfig_(ss, { includeBroadcasts: false });
    return {
      success: true,
      ok: true,
      channel_links: finalConfig.channel_links || {},
      broadcast_overrides: finalConfig.broadcast_overrides || [],
      channel_links_count: updatedChannels,
      broadcast_overrides_count: updatedOverrides,
      timestamp: nowStr
    };
  } finally {
    lock.releaseLock();
  }
}