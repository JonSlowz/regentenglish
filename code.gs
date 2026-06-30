// ==========================================
// 1. CORE FUNCTIONS
// ==========================================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Regent Account')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==========================================
// ⚡ CACHE CONFIG
// ScriptCache = แชร์ทุก user, TTL 5 นาที
// เมื่อ save/update/cancel → invalidate ทันที
// ==========================================
var _CACHE_KEY   = 'receipts_v1';
var _CACHE_TTL   = 300; // วินาที (5 นาที)
var _CHUNK_SIZE  = 90000; // bytes (ต่ำกว่า limit 100KB)

function _loadCache() {
  try {
    var c = CacheService.getScriptCache();
    // ลองอ่าน single-key ก่อน (dataset เล็ก)
    var v = c.get(_CACHE_KEY);
    if (v) return JSON.parse(v);
    // ถ้าไม่มี ลอง chunked (dataset ใหญ่)
    var n = parseInt(c.get(_CACHE_KEY + '_n') || '0');
    if (!n) return null;
    var parts = [];
    for (var i = 0; i < n; i++) {
      var part = c.get(_CACHE_KEY + '_' + i);
      if (!part) return null; // chunk หายระหว่าง TTL → miss
      parts.push(part);
    }
    return JSON.parse(parts.join(''));
  } catch (e) { return null; }
}

function _saveCache(data) {
  try {
    var c    = CacheService.getScriptCache();
    var json = JSON.stringify(data);
    if (json.length <= _CHUNK_SIZE) {
      c.put(_CACHE_KEY, json, _CACHE_TTL);
    } else {
      // แบ่ง chunk แล้ว putAll ครั้งเดียว (เร็วกว่า put ทีละอัน)
      var n   = Math.ceil(json.length / _CHUNK_SIZE);
      var obj = {};
      obj[_CACHE_KEY + '_n'] = String(n);
      for (var i = 0; i < n; i++) {
        obj[_CACHE_KEY + '_' + i] = json.substr(i * _CHUNK_SIZE, _CHUNK_SIZE);
      }
      c.putAll(obj, _CACHE_TTL);
    }
  } catch (e) { /* silently skip — ข้อมูลจะ fetch จาก sheet แทน */ }
}

function _invalidateCache() {
  try {
    var c = CacheService.getScriptCache();
    c.remove(_CACHE_KEY);
    c.remove(_CACHE_KEY + '_n');
    // chunk keys (_0, _1, ...) จะหมดอายุเองตาม TTL
  } catch (e) {}
}

// ==========================================
// 2. INIT — รวม 3 calls เป็น 1 ใช้ ss เดียว
// ==========================================
function getInitData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ⚡ ดึง RG Config sheet แค่ครั้งเดียว แล้วส่งต่อทั้ง 2 ฟังก์ชัน
  var cfgSheet = ss.getSheetByName('RG Config');

  return {
    courses:  _getCourseConfig(cfgSheet),
    payments: _getPaymentConfig(cfgSheet),
    receipts: _getReceiptList(ss),   // ใช้ cache ถ้ามี
    nextId:   _computeNextId(ss, null),
    ptData:   _getPTData(ss)   // ⚡ โหลด PT ทั้งหมดครั้งเดียว → client lookup 0ms
  };
}

// ==========================================
// 3. SAVE & UPDATE FUNCTIONS
// ==========================================
function saveNewReceipt(data, username) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    // ⚡ gen เลขจริงหลัง lock ได้ เพื่อ race-condition-safe
    data.receiptNo = _computeNextId(ss, data.date);

    var sheet = ss.getSheetByName('ข้อมูลลูกค้า');
    if (!sheet) {
      sheet = ss.insertSheet('ข้อมูลลูกค้า');
      sheet.appendRow([
        'เลขที่เอกสาร','วันที่','PT No.','ชื่อ-นามสกุล','ชื่อเล่น','เลขที่ผู้เสียภาษี','ที่อยู่','เบอร์โทร',
        'คอร์ส1','คอร์ส2','คอร์ส3','คอร์ส4',
        'ชั่วโมง1','ชั่วโมง2','ชั่วโมง3','ชั่วโมง4',
        'ราคา1','ราคา2','ราคา3','ราคา4',
        'รวมราคา',
        'ส่วนลด1','ส่วนลด2','ส่วนลด3','ส่วนลด4',
        'รวมส่วนลด',
        'จำนวนเงิน1','จำนวนเงิน2','จำนวนเงิน3','จำนวนเงิน4',
        'รวมเงินที่ชำระ','ลดพิเศษ',
        'ช่องทางชำระ','ธนาคาร','ชื่อบัญชี','หมายเหตุ','สถานะนักเรียน','สถานะ'
      ]);
    }

    var c = data.courses;
    var rowData = [
      data.receiptNo, data.date, data.placementNo, data.customerName, data.nickname,
      data.taxId, data.address, data.tel,
      c.names[0], c.names[1], c.names[2], c.names[3],
      c.hours[0], c.hours[1], c.hours[2], c.hours[3],
      c.prices[0], c.prices[1], c.prices[2], c.prices[3],
      data.sumPrice,
      c.discounts[0], c.discounts[1], c.discounts[2], c.discounts[3],
      data.sumDiscount,
      c.nets[0], c.nets[1], c.nets[2], c.nets[3],
      data.grandTotal, data.specialDiscount,
      data.paymentMethod, data.bank, data.account, data.remark, data.studentStatus,
      'Active'
    ];
    // ✅ Fix 2: หาแถวสุดท้ายจาก Col A เท่านั้น ไม่ให้ Col AP มาทำให้เพี้ยน
    var nextRow = _getLastRowByColA(sheet) + 1;
    sheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);

    _invalidateCache();
    recordLog(username, 'บันทึก', data.receiptNo, '-',
      'ใหม่: ' + data.customerName + ' (' + data.grandTotal + ' บาท)');

    // ⚡ อ่าน + cache รายการใหม่ทันทีใน call เดียวกัน → client ไม่ต้อง refresh แยก
    var freshList = _getReceiptList(ss);
    var nextId    = _computeNextId(ss, data.date);
    return { msg: 'บันทึกข้อมูลเรียบร้อยแล้ว!\nเลขที่เอกสารของคุณคือ: ' + data.receiptNo,
             receipts: freshList, nextId: nextId };

  } catch (e) {
    return { msg: '⚠️ ระบบกำลังยุ่ง กรุณากดบันทึกอีกครั้ง (' + e.message + ')', receipts: null, nextId: null };
  } finally {
    lock.releaseLock();
  }
}

function updateReceiptData(data, username) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('ข้อมูลลูกค้า');
  var lastRow = sheet.getLastRow();

  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var rowIndex = -1;
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0] == data.receiptNo) { rowIndex = i + 2; break; }
  }
  if (rowIndex === -1) return 'Error: ไม่พบใบเสร็จต้นฉบับ (' + data.receiptNo + ')';

  var oldDataRow = sheet.getRange(rowIndex, 1, 1, 38).getDisplayValues()[0];

  var c = data.courses;
  var rowData = [
    data.receiptNo, data.date, data.placementNo, data.customerName, data.nickname,
    data.taxId, data.address, data.tel,
    c.names[0], c.names[1], c.names[2], c.names[3],
    c.hours[0], c.hours[1], c.hours[2], c.hours[3],
    c.prices[0], c.prices[1], c.prices[2], c.prices[3],
    data.sumPrice,
    c.discounts[0], c.discounts[1], c.discounts[2], c.discounts[3],
    data.sumDiscount,
    c.nets[0], c.nets[1], c.nets[2], c.nets[3],
    data.grandTotal, data.specialDiscount,
    data.paymentMethod, data.bank, data.account, data.remark, data.studentStatus,
    'Active'
  ];

  var diff = getDiff(oldDataRow, rowData);
  sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);

  _invalidateCache();
  recordLog(username, 'แก้ไข', data.receiptNo,
    diff.oldStr !== '-' ? diff.oldStr : '-',
    diff.newStr !== '-' ? diff.newStr : 'กดบันทึกโดยไม่เปลี่ยนข้อมูล');

  // ⚡ return fresh list ใน call เดียวกัน
  var freshList = _getReceiptList(ss);
  return { msg: 'อัพเดตข้อมูลเรียบร้อยแล้ว (' + data.receiptNo + ')', receipts: freshList, nextId: null };
}

function cancelReceipt(receiptNo, username) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('ข้อมูลลูกค้า');
  var lastRow = sheet.getLastRow();

  // ⚡ อ่าน col A + col 38 พร้อมกัน 1 range call (เดิม 3 calls แยก)
  var cols = sheet.getRange(2, 1, lastRow - 1, 38).getValues();
  for (var i = 0; i < cols.length; i++) {
    if (cols[i][0] == receiptNo) {
      var targetRow = i + 2;
      var oldStatus = cols[i][37];
      sheet.getRange(targetRow, 38).setValue('Cancel');
      _invalidateCache();
      recordLog(username, 'ยกเลิก', receiptNo, 'Status: ' + oldStatus, 'Status: Cancel');
      // ⚡ return fresh list ใน call เดียวกัน
      var freshList = _getReceiptList(ss);
      return { msg: 'ยกเลิกใบเสร็จ ' + receiptNo + ' เรียบร้อยแล้ว', receipts: freshList, nextId: null };
    }
  }
  return { msg: 'ไม่พบใบเสร็จ ' + receiptNo, receipts: null, nextId: null };
}

// ==========================================
// 4. READ FUNCTIONS
// ==========================================
function getReceiptList() {
  return _getReceiptList(SpreadsheetApp.getActiveSpreadsheet());
}

// ✅ Fix 2+3: หาแถวสุดท้ายจาก Column A เท่านั้น (ไม่สนใจคอลัมน์อื่น เช่น AP)
function _getLastRowByColA(sheet) {
  // ⚡ ใช้ getLastRow() เป็น upper bound — ไม่อ่าน getMaxRows() (อาจ 1000+ แถวว่าง)
  var last = sheet.getLastRow();
  if (last < 1) return 1;
  var colA = sheet.getRange(1, 1, last, 1).getValues();
  for (var i = colA.length - 1; i >= 0; i--) {
    if (colA[i][0] !== '') return i + 1;
  }
  return 1;
}

// ⚡ Cache-aware receipt list
function _getReceiptList(ss) {
  var cached = _loadCache();
  if (cached) return cached;

  var sheet = ss.getSheetByName('ข้อมูลลูกค้า');
  if (!sheet) return [];

  // ✅ Fix 3: หาแถวสุดท้ายจาก Col A เท่านั้น ไม่สนใจ Col AP หรือคอลัมน์อื่น
  var lastRow = _getLastRowByColA(sheet);
  if (lastRow < 2) return [];

  var tz = ss.getSpreadsheetTimeZone();

  // ✅ Fix 3: อ่านแค่ 38 คอลัมน์ (A ถึง AL) เท่านั้น
  var displayData = sheet.getRange(2, 1, lastRow - 1, 38).getDisplayValues();
  var rawDates    = sheet.getRange(2, 2, lastRow - 1, 1).getValues();

  var result = displayData
    .map(function(row, idx) {
      if (row[0] === '') return null; // ✅ Fix 3: ข้ามแถวที่ col A ว่าง
      var year = '', month = '', dateISO = '';
      var rd   = rawDates[idx][0];
      if (rd instanceof Date && !isNaN(rd.getTime())) {
        year    = rd.getFullYear().toString();
        month   = ('0' + (rd.getMonth() + 1)).slice(-2);
        dateISO = Utilities.formatDate(rd, tz, 'yyyy-MM-dd');
      } else {
        var p = row[1].split('/');
        if (p.length === 3) { year = p[2]; month = ('0' + p[1]).slice(-2); }
        dateISO = p.length === 3 ? p[2] + '-' + ('0'+p[1]).slice(-2) + '-' + ('0'+p[0]).slice(-2) : '';
      }
      return {
        id: row[0], date: row[1], dateISO: dateISO,
        placementNo: row[2], fullName: row[3], nickname: row[4],
        taxId: row[5], address: row[6], tel: row[7],
        courses:   [row[8],  row[9],  row[10], row[11]],
        hours:     [row[12], row[13], row[14], row[15]],
        prices:    [row[16], row[17], row[18], row[19]],
        sumPrice:  row[20],
        discounts: [row[21], row[22], row[23], row[24]],
        sumDiscount: row[25],
        nets:      [row[26], row[27], row[28], row[29]],
        grandTotal: row[30], specialDiscount: row[31],
        paymentMethod: row[32], bank: row[33], account: row[34],
        remark: row[35], studentStatus: row[36], status: row[37],
        year: year, month: month
      };
    }).filter(function(r) { return r !== null; }).reverse();

  _saveCache(result); // ⚡ เก็บ cache ไว้ใช้ครั้งต่อไป
  return result;
}

// ⚡ ใช้เฉพาะ getPrintHtml (ดึง raw values สำหรับ Receipt template)
function getReceiptDetail(receiptId) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('ข้อมูลลูกค้า');
  if (!sheet) return null;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0] == receiptId) {
      var row = sheet.getRange(i + 2, 1, 1, 38).getValues()[0];
      return {
        receiptNo: row[0],
        date: Utilities.formatDate(new Date(row[1]), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd'),
        placementNo: row[2], customerName: row[3], nickname: row[4],
        taxId: row[5], address: row[6], tel: row[7],
        courses:   [row[8],  row[9],  row[10], row[11]],
        hours:     [row[12], row[13], row[14], row[15]],
        prices:    [row[16], row[17], row[18], row[19]],
        sumPrice:  row[20],
        discounts: [row[21], row[22], row[23], row[24]],
        sumDiscount: row[25],
        nets:      [row[26], row[27], row[28], row[29]],
        grandTotal: row[30], specialDiscount: row[31],
        paymentMethod: row[32], bank: row[33], account: row[34],
        remark: row[35], studentStatus: row[36]
      };
    }
  }
  return null;
}

// ==========================================
// 5. UTILITY FUNCTIONS
// ==========================================

// ⚡ Compute next receipt ID — ใช้ทั้ง server (ตอน save) และ client (preview)
function _computeNextId(ss, dateStr) {
  var sheet = (ss || SpreadsheetApp.getActiveSpreadsheet()).getSheetByName('ข้อมูลลูกค้า');
  var d     = dateStr ? new Date(dateStr) : new Date();
  var yy    = d.getFullYear().toString().substr(-2);
  var mm    = ('0' + (d.getMonth() + 1)).slice(-2);
  var prefix = 'RE' + mm + yy + '-';
  if (!sheet || sheet.getLastRow() < 2) return prefix + '01';

  var data   = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  var maxNum = 0;
  for (var i = data.length - 1; i >= 0; i--) {
    var id = data[i][0].toString();
    if (id.startsWith(prefix)) {
      var n = parseInt(id.split('-')[1]);
      if (!isNaN(n) && n > maxNum) maxNum = n;
      break; // ⚡ loop ย้อนหลัง เจออันล่าสุดแล้วหยุดเลย
    }
  }
  return prefix + ('0' + (maxNum + 1)).slice(-2);
}

// Public wrapper (ใช้เรียกจาก HTML ตอน manual refresh เลขใบเสร็จ)
function getNextReceiptNumber(dateStr) {
  return _computeNextId(null, dateStr);
}

// Public wrappers (fallback กรณีเรียกแบบ standalone)
function getCourseConfigData() {
  return _getCourseConfig(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('RG Config'));
}
function getPaymentConfigData() {
  return _getPaymentConfig(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('RG Config'));
}

// ⚡ รับ sheet มาตรงๆ — ไม่ต้องเรียก getSheetByName ซ้ำ
function _getCourseConfig(cfgSheet) {
  if (!cfgSheet || cfgSheet.getLastRow() < 2) return [];
  return cfgSheet.getRange(2, 1, cfgSheet.getLastRow() - 1, 3).getValues()
    .filter(function(r) { return r[0]; })
    .map(function(r)    { return { name: r[0], hours: r[1], price: r[2] }; });
}

function _getPaymentConfig(cfgSheet) {
  if (!cfgSheet || cfgSheet.getLastRow() < 2) return [];
  return cfgSheet.getRange(2, 5, cfgSheet.getLastRow() - 1, 3).getValues()
    .filter(function(r) { return r[0]; })
    .map(function(r)    { return { method: r[0].toString(), bank: r[1].toString(), account: r[2].toString() }; });
}

function getPTInfo(rowNumber) {
  if (!rowNumber) return null;
  var ss    = SpreadsheetApp.openById('1iX6s4Kbsglc7uEuJtFoJTkZn7C8vA-80IfeGSlDGJg4');
  var sheet = ss.getSheets()[0];
  var rowIndex = parseInt(rowNumber) + 1;
  if (rowIndex > sheet.getLastRow()) return null;
  var data = sheet.getRange(rowIndex, 1, 1, 10).getValues()[0];
  return { fullName: data[3] + ' ' + data[4], nickname: data[5], tel: data[6], address: data[9] };
}


// ⚡ โหลด PT ทั้งหมดครั้งเดียวตอน init → ส่งมาเก็บ client-side ใช้ lookup แบบ real-time 0ms
function _getPTData(ss) {
  var sheet = ss.getSheetByName('RG_PT');
  if (!sheet || sheet.getLastRow() < 2) return [];
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  var result = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    result.push({
      ptNo:     i + 1,
      fullName: (String(r[0]) + ' ' + String(r[1])).trim(),
      nickname: r[2] || '',
      tel:      r[3] || '',
      address:  r[4] || ''
    });
  }
  return result;
}

function getPrintHtml(id, lang) {
  var template = HtmlService.createTemplateFromFile('Receipt');
  var receipt  = getReceiptDetail(id);
  if (receipt) {
    template.data      = receipt;
    template.data.lang = lang || 'TH';
    return template.evaluate().setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).getContent();
  }
  return "<h3 style='text-align:center;padding:20px;'>ไม่พบข้อมูลใบเสร็จ</h3>";
}

function checkLogin(username, password) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('RG Config');
  if (!sheet || sheet.getLastRow() < 2) return false;
  var data = sheet.getRange(2, 9, sheet.getLastRow() - 1, 2).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === username && String(data[i][1]).trim() === password) return true;
  }
  return false;
}

function recordLog(user, action, refId, oldVal, newVal) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('RG_log');
  if (!sheet) {
    sheet = ss.insertSheet('RG_log');
    sheet.appendRow(['User', 'Date/Time', 'Action', 'Receipt ID', 'Old Data', 'New Data']);
  }
  sheet.appendRow([user, new Date(), action, refId, oldVal, newVal]);
}

var FIELD_NAMES = [
  'เลขที่','วันที่','PT No.','ชื่อ','ชื่อเล่น','Tax ID','ที่อยู่','เบอร์โทร',
  'คอร์ส 1','คอร์ส 2','คอร์ส 3','คอร์ส 4',
  'ชั่วโมง 1','ชั่วโมง 2','ชั่วโมง 3','ชั่วโมง 4',
  'ราคา 1','ราคา 2','ราคา 3','ราคา 4', 'รวมราคา',
  'ส่วนลด 1','ส่วนลด 2','ส่วนลด 3','ส่วนลด 4', 'รวมส่วนลด',
  'สุทธิ 1','สุทธิ 2','สุทธิ 3','สุทธิ 4',
  'ยอดสุทธิรวม','ลดพิเศษ',
  'ช่องทาง','ธนาคาร','บัญชี','หมายเหตุ','สถานะ นร.','สถานะ'
];

function getDiff(oldData, newData) {
  var oldLog = [], newLog = [];
  for (var i = 0; i < FIELD_NAMES.length; i++) {
    var vOld = String(oldData[i] || '').trim().replace(/^'/, '');
    var vNew = String(newData[i] || '').trim().replace(/^'/, '');
    if (vOld !== vNew) { oldLog.push(FIELD_NAMES[i] + ': ' + vOld); newLog.push(FIELD_NAMES[i] + ': ' + vNew); }
  }
  if (oldLog.length === 0) return { oldStr: '-', newStr: '-' };
  return { oldStr: oldLog.join('\n'), newStr: newLog.join('\n') };
}
