// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 達文西排刀預約系統 — Google Apps Script 後端
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
//
// 【部署步驟】
// 1. 開啟 Google Sheets，建立新試算表（名稱隨意）
// 2. 點選上方選單「擴充功能」→「Apps Script」
// 3. 將此檔案內容完整貼入編輯器，取代預設的 function myFunction()
// 4. 點左上「儲存」（磁碟圖示）
// 5. 在編輯器上方選擇函式「setupSheet」，點「執行」→ 授權並允許
//    （這會自動建立表頭與格式，只需執行一次）
// 6. 點「部署」→「新增部署作業」
//    - 類型：網頁應用程式
//    - 說明：達文西排刀系統 v1
//    - 執行身份：我（您自己的 Google 帳號）
//    - 誰可以存取：任何人
//    → 點「部署」→ 複製「網頁應用程式網址」
// 7. 將複製的網址貼入 index.html 最底部的 API_URL 變數
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

const SHEET_NAME = 'Bookings';
const HEADERS = [
  'id', 'date', 'machine', 'startH', 'endH',
  'dept', 'customDept', 'doctor', 'assist',
  'procedure', 'duration', 'mrn', 'note', 'createdAt'
];

// ── 初始化表頭（只需執行一次）──
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  sheet.clearContents();
  sheet.appendRow(HEADERS);

  // 表頭樣式（北醫紅）
  const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
  headerRange.setFontWeight('bold')
             .setBackground('#C41230')
             .setFontColor('#ffffff')
             .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);

  // 欄寬
  const widths = [120,110,70,70,70,80,100,100,100,200,70,100,200,160];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  SpreadsheetApp.flush();
  Logger.log('✅ setupSheet 完成');
}

function getSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
}

// ── GET：讀取全部預約 ──
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'getAll';
  if (action === 'getAll') return jsonOk(getAllBookings());
  return jsonErr('Unknown action');
}

// ── POST：新增 / 更新 / 刪除 ──
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    switch (body.action) {
      case 'create': return jsonOk(createBooking(body.booking));
      case 'update': return jsonOk(updateBooking(body.booking));
      case 'delete': return jsonOk(deleteBooking(body.id));
      default:       return jsonErr('Unknown action: ' + body.action);
    }
  } catch (err) {
    return jsonErr(err.message);
  }
}

// ── 讀取全部 ──
function getAllBookings() {
  const sheet = getSheet();
  const rows  = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { bookings: [] };

  const headers  = rows[0];
  const bookings = rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      const v = row[i];
      obj[h] = (v instanceof Date)
        ? Utilities.formatDate(v, 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss')
        : String(v);
    });
    return obj;
  });
  return { bookings };
}

// ── 新增 ──
function createBooking(b) {
  const sheet = getSheet();
  const now   = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');
  // 使用前端傳來的 id（timestamp），保持一致
  sheet.appendRow([
    b.id, b.date, b.machine, b.startH, b.endH,
    b.dept, b.customDept || '', b.doctor, b.assist || '',
    b.procedure, b.duration || '', b.mrn || '', b.note || '', now
  ]);
  return { id: b.id };
}

// ── 更新 ──
function updateBooking(b) {
  const sheet = getSheet();
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(b.id)) {
      sheet.getRange(i + 1, 1, 1, HEADERS.length).setValues([[
        b.id, b.date, b.machine, b.startH, b.endH,
        b.dept, b.customDept || '', b.doctor, b.assist || '',
        b.procedure, b.duration || '', b.mrn || '', b.note || '',
        rows[i][13] // 保留原 createdAt
      ]]);
      return { updated: true };
    }
  }
  return { updated: false, note: 'ID not found' };
}

// ── 刪除 ──
function deleteBooking(id) {
  const sheet = getSheet();
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { deleted: true };
    }
  }
  return { deleted: false, note: 'ID not found' };
}

// ── 輔助：回應格式 ──
function jsonOk(data) {
  return ContentService
    .createTextOutput(JSON.stringify({ success: true, ...data }))
    .setMimeType(ContentService.MimeType.JSON);
}
function jsonErr(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ success: false, error: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}
