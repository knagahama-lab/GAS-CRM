/**
 * ImportModule.gs
 * CSVデータ移行 - 他システムからエクスポートしたCSVを
 * 👥 Customers / 💼 Deals / 📝 Activities の各シートへ直接取り込む（管理者のみ）
 *
 * - ヘッダー名でシート列にマッピングするため、列順が異なっていても取り込める
 * - 日付・数値列は自動でDate/Number型に変換してから書き込む（文字列のまま入れない）
 * - 主キー列（customer_id/deal_id/activity_id）が既存シートと重複する行は自動スキップ
 *   → 同じCSVを誤って2回実行しても二重登録されない
 */

const IMPORT_DATE_COLUMNS = ['last_contact_date', 'created_at', 'updated_at', 'close_date', 'activity_date', 'next_action_date', 'meeting_date'];
const IMPORT_NUMBER_COLUMNS = ['amount', 'probability', 'weighted_amount', 'duration_sec'];

function _coerceImportValue(colName, raw) {
  const v = (raw === undefined || raw === null) ? '' : String(raw).trim();
  if (!v) return '';

  if (IMPORT_DATE_COLUMNS.indexOf(colName) !== -1) {
    const d = new Date(v);
    return isNaN(d.getTime()) ? v : d;
  }
  if (IMPORT_NUMBER_COLUMNS.indexOf(colName) !== -1) {
    const n = Number(v);
    return isNaN(n) ? 0 : n;
  }
  return v;
}

function _importCsvToSheet(sheetName, csvText, requiredKeyColumn) {
  if (typeof csvText !== 'string' || !csvText.trim()) {
    throw new Error('CSVの内容が空です。');
  }
  // UTF-8 BOMを除去（utf-8-sigで保存されたCSV対策）
  if (csvText.charCodeAt(0) === 0xFEFF) csvText = csvText.slice(1);

  const sheet = getSheet(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const keyIdx = requiredKeyColumn ? headers.indexOf(requiredKeyColumn) : -1;

  // 既存の主キー一覧を取得（重複インポート防止）
  const existingIds = new Set();
  if (keyIdx !== -1 && sheet.getLastRow() > 1) {
    const existing = sheet.getRange(2, keyIdx + 1, sheet.getLastRow() - 1, 1).getValues();
    existing.forEach(r => { if (r[0]) existingIds.add(String(r[0])); });
  }

  let rows;
  try {
    rows = Utilities.parseCsv(csvText);
  } catch (e) {
    throw new Error('CSVの解析に失敗しました: ' + e.message);
  }
  if (!rows.length) return { imported: 0, skipped: 0, duplicate: 0, total: 0 };

  const csvHeaders = rows[0].map(h => String(h).trim());
  const unknownColumns = csvHeaders.filter(h => headers.indexOf(h) === -1);
  const dataRows = rows.slice(1);

  const outRows = [];
  let skipped = 0, duplicate = 0;

  dataRows.forEach(r => {
    if (!r.some(v => String(v).trim() !== '')) return; // 完全な空行はスキップ

    const rowArr = headers.map(() => '');
    csvHeaders.forEach((h, i) => {
      const idx = headers.indexOf(h);
      if (idx !== -1) rowArr[idx] = _coerceImportValue(h, r[i]);
    });

    if (keyIdx !== -1) {
      const keyVal = String(rowArr[keyIdx] || '').trim();
      if (!keyVal) { skipped++; return; }
      if (existingIds.has(keyVal)) { duplicate++; return; }
      existingIds.add(keyVal);
    }
    outRows.push(rowArr);
  });

  if (outRows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, outRows.length, headers.length).setValues(outRows);
  }

  return {
    imported: outRows.length,
    skipped,
    duplicate,
    total: dataRows.length,
    unknownColumns,
  };
}

function importCustomersCsv(csvText) {
  return wrapAction(() => {
    _requireAdmin();
    requireFeature('FEATURE_CUSTOMERS');
    const result = _importCsvToSheet(CUSTOMERS_SHEET, csvText, 'customer_id');
    logInfo(`顧客CSVインポート: ${result.imported}件（重複${result.duplicate}件 / 不正${result.skipped}件 / 全${result.total}件中）`);
    return successResponse(result);
  });
}

function importDealsCsv(csvText) {
  return wrapAction(() => {
    _requireAdmin();
    requireFeature('FEATURE_DEALS');
    const result = _importCsvToSheet(DEALS_SHEET, csvText, 'deal_id');
    logInfo(`商談CSVインポート: ${result.imported}件（重複${result.duplicate}件 / 不正${result.skipped}件 / 全${result.total}件中）`);
    return successResponse(result);
  });
}

function importActivitiesCsv(csvText) {
  return wrapAction(() => {
    _requireAdmin();
    requireFeature('FEATURE_ACTIVITIES');
    const result = _importCsvToSheet(ACTIVITIES_SHEET, csvText, 'activity_id');
    logInfo(`活動CSVインポート: ${result.imported}件（重複${result.duplicate}件 / 不正${result.skipped}件 / 全${result.total}件中）`);
    return successResponse(result);
  });
}
