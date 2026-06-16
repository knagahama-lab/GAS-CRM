/**
 * FeatureManager.gs
 * 機能ON/OFF制御 - Settingsシートから設定を読み込みキャッシュ管理
 */

const SETTINGS_SHEET = 'Settings';
const CACHE_KEY = 'CRM_SETTINGS';
const CACHE_EXPIRY = 300; // 5分

// 許可された機能名ホワイトリスト
const ALLOWED_FEATURE_KEYS = [
  'FEATURE_CUSTOMERS', 'FEATURE_DEALS', 'FEATURE_ACTIVITIES', 'FEATURE_CONTACTS',
  'FEATURE_EMAIL_TRACKING', 'FEATURE_EMAIL_SEND', 'FEATURE_CAL_CREATE', 'FEATURE_CAL_REMINDER',
  'FEATURE_DASHBOARD', 'FEATURE_PIPELINE', 'FEATURE_REPORT_AUTO', 'FEATURE_TARGETS',
  'FEATURE_KANBAN', 'FEATURE_SEARCH', 'FEATURE_DEDUP',
  // 設定系キー
  'REPORT_RECIPIENTS', 'REPORT_SCHEDULE', 'REMINDER_DAYS_BEFORE', 'EMAIL_SCAN_INTERVAL',
  'ADMIN_EMAILS', 'SYSTEM_VERSION',
];

/**
 * 機能が有効かどうかを返す
 */
function isEnabled(key) {
  const settings = _getSettingsCache();
  const val = settings[key];
  if (val === undefined) return false;
  return String(val).toUpperCase() === 'TRUE';
}

/**
 * 設定値を文字列/数値で取得
 */
function getSetting(key, defaultVal = '') {
  const settings = _getSettingsCache();
  const val = settings[key];
  return val !== undefined && val !== '' ? val : defaultVal;
}

/**
 * 設定を更新（管理者のみ）
 */
function setFeature(key, value) {
  return wrapAction(() => {
    _requireAdmin();
    if (!ALLOWED_FEATURE_KEYS.includes(key)) {
      return errorResponse(`機能キー「${key}」は許可されていません。`);
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SETTINGS_SHEET);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(value);
        _clearCache();
        return successResponse({ key });
      }
    }
    // 新規追加
    sheet.appendRow([key, value, new Date(), Session.getActiveUser().getEmail()]);
    _clearCache();
    return successResponse({ key });
  });
}

/**
 * 全設定をオブジェクトで返す（管理画面用）
 */
function getAllSettings() {
  return _getSettingsCache();
}

/**
 * 機能がOFFの場合は例外をスロー
 */
function requireFeature(key) {
  if (!isEnabled(key)) {
    throw new Error(`機能「${key}」は現在無効です。管理者にお問い合わせください。`);
  }
}

// ── 内部関数 ──────────────────────────────────────────────

function _getSettingsCache() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CACHE_KEY);
  if (cached) return JSON.parse(cached);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SETTINGS_SHEET);
  if (!sheet) return {};

  const data = sheet.getDataRange().getValues();
  const settings = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) settings[data[i][0]] = data[i][1];
  }
  cache.put(CACHE_KEY, JSON.stringify(settings), CACHE_EXPIRY);
  return settings;
}

function _clearCache() {
  CacheService.getScriptCache().remove(CACHE_KEY);
}

function _requireAdmin() {
  const email = Session.getActiveUser().getEmail();
  const admins = getSetting('ADMIN_EMAILS', '');
  if (!admins.includes(email)) {
    throw new Error('この操作は管理者のみ実行できます。');
  }
}
