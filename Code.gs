/**
 * Credit Card Expense Tracker - Google Apps Script
 * ระบบบันทึกและจัดการค่าใช้จ่ายบัตรเครดิต ตัดรอบทุกวันที่ 5
 */

const SHEET_NAME = 'Transactions';
const HEADERS = ['ID', 'Date', 'Memo', 'Amount', 'Type', 'CreatedAt'];

/**
 * ฟังก์ชันเริ่มต้นสร้าง/ดึง Sheet
 */
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    // จัด format หัวตาราง
    const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setBackground('#2563eb');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * ให้บริการ Web App และ API
 */
function doGet(e) {
  // หากเรียกเป็น API (เช่น ?action=getData)
  if (e && e.parameter && e.parameter.action === 'getData') {
    const result = getInitialData();
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // หากเปิดเป็นหน้าเว็บตรงๆ ผ่าน Apps Script
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Credit Card Tracker')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * รองรับ API POST จาก GitHub Pages (CORS Friendly)
 */
function doPost(e) {
  try {
    let payload = {};
    if (e && e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    }
    const action = payload.action;
    let result = {};

    if (action === 'getData') {
      result = getInitialData();
    } else if (action === 'add') {
      result = addTransaction(payload.data);
    } else if (action === 'update') {
      result = updateTransaction(payload.data);
    } else if (action === 'delete') {
      result = deleteTransaction(payload.id);
    } else {
      result = getInitialData();
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * คำนวณช่วงวันที่รอบบิล (ตัดรอบวันที่ 5 ถึง 4 ของเดือนถัดไป)
 * @param {Date} dateObj
 * @returns {object} { cycleName, startDate, endDate }
 */
function getBillingCycle(dateObj) {
  const d = new Date(dateObj);
  const year = d.getFullYear();
  const month = d.getMonth(); // 0-11
  const day = d.getDate();

  let startYear = year;
  let startMonth = month;
  let endYear = year;
  let endMonth = month + 1;

  if (day < 5) {
    // ถ้าวันที่ < 5 แสดงว่ายังอยู่ในรอบของเดือนก่อน
    startMonth = month - 1;
    endMonth = month;
    if (startMonth < 0) {
      startMonth = 11;
      startYear = year - 1;
    }
  } else {
    // วันที่ >= 5 อยู่ในรอบเดือนนี้
    if (endMonth > 11) {
      endMonth = 0;
      endYear = year + 1;
    }
  }

  const startDate = new Date(startYear, startMonth, 5, 0, 0, 0, 0);
  const endDate = new Date(endYear, endMonth, 4, 23, 59, 59, 999);

  return {
    start: startDate,
    end: endDate,
    key: Utilities.formatDate(startDate, Session.getScriptTimeZone(), 'yyyy-MM')
  };
}

/**
 * ดึงข้อมูลสรุป (รอบก่อนหน้า + รอบปัจจุบัน) และรายการทั้งหมด
 */
function getInitialData() {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const tz = Session.getScriptTimeZone();

  const now = new Date();
  const currentCycle = getBillingCycle(now);

  // คำนวณรอบก่อนหน้า โดยลบจากวันเริ่มต้นของรอบปัจจุบันไป 1 วัน
  const prevDate = new Date(currentCycle.start.getTime() - (24 * 60 * 60 * 1000));
  const previousCycle = getBillingCycle(prevDate);

  const transactions = [];

  let curTotal = 0, curPersonal = 0, curSharing = 0;
  let prevTotal = 0, prevPersonal = 0, prevSharing = 0;

  // วนลูปอ่านข้อมูล (ข้ามแถวหัวตาราง i=0)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] && !row[1]) continue;

    const id = String(row[0]);
    const rawDate = row[1];
    const memo = String(row[2] || '');
    const amount = Number(row[3]) || 0;
    const type = String(row[4] || 'Personal');
    const createdAt = row[5] ? Utilities.formatDate(new Date(row[5]), tz, 'yyyy-MM-dd HH:mm') : '';

    let dateObj;
    if (rawDate instanceof Date) {
      dateObj = rawDate;
    } else {
      dateObj = new Date(rawDate);
    }

    const dateStr = Utilities.formatDate(dateObj, tz, 'yyyy-MM-dd');
    const txTime = dateObj.getTime();

    // สะสมยอดรอบปัจจุบัน
    if (txTime >= currentCycle.start.getTime() && txTime <= currentCycle.end.getTime()) {
      curTotal += amount;
      if (type.toLowerCase() === 'sharing') {
        curSharing += amount;
      } else {
        curPersonal += amount;
      }
    }

    // สะสมยอดรอบก่อนหน้า
    if (txTime >= previousCycle.start.getTime() && txTime <= previousCycle.end.getTime()) {
      prevTotal += amount;
      if (type.toLowerCase() === 'sharing') {
        prevSharing += amount;
      } else {
        prevPersonal += amount;
      }
    }

    transactions.push({
      id: id,
      date: dateStr,
      memo: memo,
      amount: amount,
      type: type,
      createdAt: createdAt,
      timestamp: txTime
    });
  }

  // เรียงรายการจากล่าสุดไปเก่าสุด
  transactions.sort((a, b) => b.timestamp - a.timestamp);

  const thaiMonths = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
  const formatCycleRange = (cycle) => {
    const s = cycle.start;
    const e = cycle.end;
    return `5 ${thaiMonths[s.getMonth()]} - 4 ${thaiMonths[e.getMonth()]}`;
  };

  return {
    success: true,
    previousCycle: {
      label: 'รอบก่อนหน้า',
      dateRange: formatCycleRange(previousCycle),
      total: prevTotal,
      personal: prevPersonal,
      sharing: prevSharing
    },
    currentCycle: {
      label: 'รอบปัจจุบัน',
      dateRange: formatCycleRange(currentCycle),
      total: curTotal,
      personal: curPersonal,
      sharing: curSharing
    },
    transactions: transactions
  };
}

/**
 * บันทึกรายการใหม่
 */
function addTransaction(formData) {
  const sheet = getSheet();
  const tz = Session.getScriptTimeZone();
  const id = 'TXN-' + new Date().getTime();
  const date = formData.date; // YYYY-MM-DD
  const memo = formData.memo.trim();
  const amount = parseFloat(formData.amount);
  const type = formData.type; // Personal | Sharing
  const createdAt = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');

  sheet.appendRow([id, date, memo, amount, type, createdAt]);

  return getInitialData();
}

/**
 * แก้ไขรายการ
 */
function updateTransaction(formData) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const id = formData.id;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      const rowIdx = i + 1;
      sheet.getRange(rowIdx, 2).setValue(formData.date);
      sheet.getRange(rowIdx, 3).setValue(formData.memo.trim());
      sheet.getRange(rowIdx, 4).setValue(parseFloat(formData.amount));
      sheet.getRange(rowIdx, 5).setValue(formData.type);
      break;
    }
  }

  return getInitialData();
}

/**
 * ลบรายการ
 */
function deleteTransaction(id) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      break;
    }
  }

  return getInitialData();
}
