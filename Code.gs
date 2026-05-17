// ============================================================
//  Papa Jones — Interview CRM  |  Google Apps Script Backend
//  Sheet ID  : 1zuq4RIpAkPDshtq4wphNMjYXDr632mKpo9eyV2GcnLg
//  Deploy as : Web App → Execute as Me → Anyone can access
// ============================================================

const SHEET_ID        = '1zuq4RIpAkPDshtq4wphNMjYXDr632mKpo9eyV2GcnLg';
const DATA_SHEET_NAME = 'Sheet1';
const LOGIN_SHEET     = 'Login';
const PASSWORD        = 'Vantage2026';

// ── Entry point ──────────────────────────────────────────────
function doGet(e) {
  const p      = e.parameter;
  const action = p.action || '';
  let   result;

  try {
    switch (action) {
      case 'login':     result = login(p.email, p.password);                           break;
      case 'getData':   result = getData();                                              break;
      case 'addRow':    result = addRow(JSON.parse(p.data));                            break;
      case 'updateRow': result = updateRow(+p.rowIndex, JSON.parse(p.data));            break;
      case 'deleteRow': result = deleteRow(+p.rowIndex);                               break;
      default:          result = { success: false, error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { success: false, error: err.toString() };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Login ─────────────────────────────────────────────────────
// Login sheet columns (row 1 = headers):  Email | Role | Name
// Roles: Admin | Editor | Viewer
function login(email, password) {
  if (!email || !password)
    return { success: false, message: 'البريد وكلمة المرور مطلوبان' };

  if (password !== PASSWORD)
    return { success: false, message: 'كلمة المرور غير صحيحة' };

  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(LOGIN_SHEET);
  if (!sheet) return { success: false, message: 'Login sheet not found' };

  const rows = sheet.getDataRange().getValues();          // row 0 = headers
  for (let i = 1; i < rows.length; i++) {
    const rowEmail = String(rows[i][0] || '').trim().toLowerCase();
    if (rowEmail === email.trim().toLowerCase()) {
      return {
        success : true,
        role    : String(rows[i][1] || 'Viewer').trim(),
        name    : String(rows[i][2] || email).trim()
      };
    }
  }

  return { success: false, message: 'البريد الإلكتروني غير مسجّل في النظام' };
}

// ── Read data ─────────────────────────────────────────────────
function getData() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DATA_SHEET_NAME);
  if (!sheet) return { success: false, error: 'Sheet1 not found' };

  const all = sheet.getDataRange().getValues();
  if (all.length === 0) return { success: true, headers: [], rows: [] };

  return {
    success : true,
    headers : all[0].map(h => String(h)),
    rows    : all.slice(1).map(row => row.map(cell => {
      // Convert Date objects to readable strings
      if (cell instanceof Date) {
        return Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      return String(cell);
    }))
  };
}

// ── Add row ───────────────────────────────────────────────────
function addRow(rowData) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DATA_SHEET_NAME);
  sheet.appendRow(rowData);
  return { success: true };
}

// ── Update row ────────────────────────────────────────────────
// rowIndex is 0-based (relative to data rows, excluding header)
function updateRow(rowIndex, rowData) {
  const sheet   = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DATA_SHEET_NAME);
  const sheetRow = rowIndex + 2;          // +1 header, +1 for 1-based indexing
  sheet.getRange(sheetRow, 1, 1, rowData.length).setValues([rowData]);
  return { success: true };
}

// ── Delete row ────────────────────────────────────────────────
function deleteRow(rowIndex) {
  const sheet   = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DATA_SHEET_NAME);
  const sheetRow = rowIndex + 2;
  sheet.deleteRow(sheetRow);
  return { success: true };
}


// البوابة الرسمية لاستقبال الطلبات من الـ index.html الجديد
function executeAction(params) {
  const action = params.action || '';
  let result;

  try {
    switch (action) {
      case 'login':     result = login(params.email, params.password); break;
      case 'getData':   result = getData(); break;
      case 'addRow':    result = addRow(JSON.parse(params.data)); break;
      case 'updateRow': result = updateRow(+params.rowIndex, JSON.parse(params.data)); break;
      case 'deleteRow': result = deleteRow(+params.rowIndex); break;
      default:          result = { success: false, error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { success: false, error: err.toString() };
  }

  // نرجع النتيجة كـ Object وجوجل هتوصلها فوراً للـ Frontend
  return result;
}
