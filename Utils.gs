/**
 * Utils.gs
 * 共通ユーティリティ関数
 */

// ── ID生成 ────────────────────────────────────────────────

function generateId(prefix) {
  const now = new Date();
  const date = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMdd');
  const rand = Math.random().toString(36).substring(2, 6).toUpperCase();
  return `${prefix}-${date}-${rand}`;
}

// ── 日付ユーティリティ ────────────────────────────────────

function formatDate(date) {
  if (!date) return '';
  if (typeof date === 'string') return date;
  return Utilities.formatDate(new Date(date), 'Asia/Tokyo', 'yyyy/MM/dd');
}

function formatDateTime(date) {
  if (!date) return '';
  return Utilities.formatDate(new Date(date), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
}

function now() {
  return new Date();
}

function daysFromNow(days) {
  const d = new Date();
  d.setDate(d.getDate() + days);
  return d;
}

// ── スプレッドシートユーティリティ ───────────────────────

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error(`シート「${name}」が見つかりません。`);
  return sheet;
}

function getSheetData(name) {
  const sheet = getSheet(name);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

function appendRow(sheetName, rowData) {
  const sheet = getSheet(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => rowData[h] !== undefined ? rowData[h] : '');
  sheet.appendRow(row);
}

function updateRow(sheetName, idColumn, idValue, rowData) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf(idColumn);
  if (idIdx === -1) throw new Error(`列「${idColumn}」が見つかりません。`);

  for (let i = 1; i < data.length; i++) {
    if (data[i][idIdx] === idValue) {
      headers.forEach((h, j) => {
        if (rowData[h] !== undefined) {
          sheet.getRange(i + 1, j + 1).setValue(rowData[h]);
        }
      });
      return true;
    }
  }
  return false;
}

function findRow(sheetName, idColumn, idValue) {
  const data = getSheetData(sheetName);
  return data.find(row => row[idColumn] === idValue) || null;
}

// ── 文字列ユーティリティ ──────────────────────────────────

function normalizePhone(phone) {
  if (!phone) return '';
  return String(phone).replace(/[-\s\(\)]/g, '');
}

function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

// ── ログユーティリティ ────────────────────────────────────

function logInfo(msg) {
  console.log(`[INFO] ${formatDateTime(new Date())} - ${msg}`);
}

function logError(msg, err) {
  console.error(`[ERROR] ${formatDateTime(new Date())} - ${msg}`, err ? err.message : '');
}

// ── レスポンスユーティリティ ──────────────────────────────

function successResponse(data) {
  return JSON.stringify({ success: true, data });
}

function errorResponse(message) {
  return JSON.stringify({ success: false, error: message });
}
