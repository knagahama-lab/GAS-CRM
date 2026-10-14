/**
 * PriceNegotiationModule.gs
 * 価格改定（値上げ）交渉支援AI
 *
 * サプライチェーン調達向けの「価格交渉自動化AI」（交渉戦略の学習→価格交渉→交渉合意）を、
 * 営業が顧客へ価格改定を申し入れる場面向けに組み替えたもの。
 *   ①交渉シナリオ作成（対象品目・現行/希望価格・値上げ理由を入力）
 *   → ②AIが顧客との関係性情報（活動履歴・顧客分析）を踏まえ交渉戦略・想定問答・
 *      譲歩案・案内文ドラフトを生成（創作禁止・関係維持重視）
 *   → ③営業が交渉を進め、ステータス・合意価格を記録（活動履歴にも自動記録）
 */

const PRICE_NEGOTIATIONS_SHEET = '💹 PriceNegotiations';

// ── 既存デプロイへの自動マイグレーション ─────────────────────

function ensurePriceNegotiationSetup() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let settingsChanged = false;

    if (!ss.getSheetByName(PRICE_NEGOTIATIONS_SHEET)) {
      const sheet = ss.insertSheet(PRICE_NEGOTIATIONS_SHEET);
      const headers = ['negotiation_id', 'customer_id', 'deal_id', 'item_name',
        'current_price', 'target_price', 'increase_rate', 'reason_category', 'reason_detail',
        'status', 'strategy_json', 'letter_draft', 'agreed_price', 'agreed_date',
        'assigned_user', 'created_at', 'updated_at'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers])
        .setBackground('#1A56DB').setFontColor('#FFFFFF').setFontWeight('bold');
      sheet.autoResizeColumns(1, headers.length);
      logInfo('PriceNegotiations シートを自動補完作成しました');
    }

    // MasterDataに値上げ理由区分（PRICE_REASON）が無ければ追記する
    const masterSheet = ss.getSheetByName('📋 MasterData');
    if (masterSheet) {
      const data = masterSheet.getDataRange().getValues();
      const hasPriceReason = data.slice(1).some(r => r[0] === 'PRICE_REASON');
      if (!hasPriceReason) {
        const rows = [
          ['PRICE_REASON', 'MATERIAL', '原材料費高騰', 1, 'TRUE'],
          ['PRICE_REASON', 'FX', '為替変動', 2, 'TRUE'],
          ['PRICE_REASON', 'ENERGY', 'エネルギー価格高騰', 3, 'TRUE'],
          ['PRICE_REASON', 'LABOR', '労務費上昇', 4, 'TRUE'],
          ['PRICE_REASON', 'LOGISTICS', '物流費上昇', 5, 'TRUE'],
          ['PRICE_REASON', 'OTHER', 'その他', 6, 'TRUE'],
        ];
        masterSheet.getRange(masterSheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
        logInfo('MasterData に値上げ理由区分（PRICE_REASON）を自動補完しました');
      }
    }

    const settingsSheet = ss.getSheetByName(SETTINGS_SHEET);
    if (settingsSheet) {
      const data = settingsSheet.getDataRange().getValues();
      const existingKeys = data.slice(1).map(r => r[0]);
      const defaults = [
        ['FEATURE_PRICE_NEGOTIATION', 'TRUE', now(), '', '価格改定（値上げ）交渉支援AI'],
      ];
      const missing = defaults.filter(d => existingKeys.indexOf(d[0]) === -1);
      if (missing.length) {
        settingsSheet.getRange(settingsSheet.getLastRow() + 1, 1, missing.length, missing[0].length).setValues(missing);
        settingsChanged = true;
        logInfo(`Settings に価格交渉AI設定キーを自動補完しました: ${missing.map(m => m[0]).join(', ')}`);
      }
    }

    if (settingsChanged) _clearCache();
  } catch (e) {
    logError('価格交渉AIの自動セットアップ補完に失敗しました', e);
  }
}

// ── 交渉シナリオ生成 ──────────────────────────────────────────

function generatePriceNegotiation(data) {
  return wrapAction(() => {
    requireFeature('FEATURE_PRICE_NEGOTIATION');

    const valErr = validateRequired(data, ['customer_id', 'item_name']);
    if (valErr) return errorResponse(valErr.error);

    const customer = findRow(CUSTOMERS_SHEET, 'customer_id', data.customer_id);
    if (!customer) return errorResponse(`顧客ID「${data.customer_id}」が見つかりません。`);

    const currentPrice = Number(data.current_price) || 0;
    const targetPrice = Number(data.target_price) || 0;
    const increaseRate = (currentPrice > 0 && targetPrice > 0)
      ? Math.round((targetPrice - currentPrice) / currentPrice * 1000) / 10
      : '';

    // 顧客との関係性情報を集約（創作防止のため、実データのみを渡す）
    const activities = getSheetData(ACTIVITIES_SHEET)
      .filter(a => a.customer_id === data.customer_id)
      .sort((a, b) => new Date(b.activity_date) - new Date(a.activity_date));
    const recentSummary = activities.slice(0, 5)
      .map(a => `${formatDate(a.activity_date)} [${a.type}] ${a.subject}: ${(a.content || '').slice(0, 150)}`)
      .join('\n') || '';

    let problemAwareness = '';
    const insights = getSheetData(INSIGHTS_SHEET).filter(r => r.customer_id === data.customer_id);
    if (insights.length) {
      insights.sort((a, b) => new Date(b.created_at) - new Date(a.created_at));
      problemAwareness = insights[0].problem_awareness || '';
    }

    const pastNegotiations = getSheetData(PRICE_NEGOTIATIONS_SHEET)
      .filter(r => r.customer_id === data.customer_id)
      .map(r => `${r.item_name}: ${r.current_price}→${r.target_price}（${r.status}${r.agreed_price ? ', 合意価格' + r.agreed_price : ''}）`)
      .join('\n');

    const strategy = generatePriceNegotiationStrategyAI({
      companyName: customer.company_name,
      itemName: data.item_name,
      currentPrice: data.current_price,
      targetPrice: data.target_price,
      increaseRate,
      reasonCategory: data.reason_category_label || data.reason_category || '',
      reasonDetail: data.reason_detail || '',
      customerStatus: customer.status,
      recentSummary,
      problemAwareness,
      pastNegotiations,
    });

    const negotiationId = generateId('NEG');
    const row = {
      negotiation_id: negotiationId,
      customer_id: data.customer_id,
      deal_id: data.deal_id || '',
      item_name: sanitizeString(data.item_name),
      current_price: currentPrice,
      target_price: targetPrice,
      increase_rate: increaseRate,
      reason_category: data.reason_category || '',
      reason_detail: sanitizeString(data.reason_detail),
      status: 'DRAFT',
      strategy_json: JSON.stringify(strategy),
      letter_draft: strategy.letter_draft || '',
      agreed_price: '',
      agreed_date: '',
      assigned_user: Session.getActiveUser().getEmail(),
      created_at: now(),
      updated_at: now(),
    };
    appendRow(PRICE_NEGOTIATIONS_SHEET, row);

    logInfo(`価格交渉シナリオ生成: ${negotiationId}（顧客${data.customer_id} / ${data.item_name}）`);
    return successResponse(row);
  });
}

function regenerateNegotiationStrategy(negotiationId) {
  return wrapAction(() => {
    requireFeature('FEATURE_PRICE_NEGOTIATION');
    const neg = findRow(PRICE_NEGOTIATIONS_SHEET, 'negotiation_id', negotiationId);
    if (!neg) return errorResponse(`交渉ID「${negotiationId}」が見つかりません。`);

    const customer = findRow(CUSTOMERS_SHEET, 'customer_id', neg.customer_id);
    const activities = getSheetData(ACTIVITIES_SHEET)
      .filter(a => a.customer_id === neg.customer_id)
      .sort((a, b) => new Date(b.activity_date) - new Date(a.activity_date));
    const recentSummary = activities.slice(0, 5)
      .map(a => `${formatDate(a.activity_date)} [${a.type}] ${a.subject}: ${(a.content || '').slice(0, 150)}`)
      .join('\n') || '';

    let problemAwareness = '';
    const insights = getSheetData(INSIGHTS_SHEET).filter(r => r.customer_id === neg.customer_id);
    if (insights.length) {
      insights.sort((a, b) => new Date(b.created_at) - new Date(a.created_at));
      problemAwareness = insights[0].problem_awareness || '';
    }

    const strategy = generatePriceNegotiationStrategyAI({
      companyName: customer ? customer.company_name : neg.customer_id,
      itemName: neg.item_name,
      currentPrice: neg.current_price,
      targetPrice: neg.target_price,
      increaseRate: neg.increase_rate,
      reasonCategory: neg.reason_category,
      reasonDetail: neg.reason_detail,
      customerStatus: customer ? customer.status : '',
      recentSummary,
      problemAwareness,
      pastNegotiations: '',
    });

    updateRow(PRICE_NEGOTIATIONS_SHEET, 'negotiation_id', negotiationId, {
      strategy_json: JSON.stringify(strategy),
      letter_draft: strategy.letter_draft || '',
      updated_at: now(),
    });

    const updated = findRow(PRICE_NEGOTIATIONS_SHEET, 'negotiation_id', negotiationId);
    return successResponse(updated);
  });
}

// ── 一覧・詳細・ステータス管理 ────────────────────────────────

function getAllNegotiations(userEmail) {
  return wrapAction(() => {
    requireFeature('FEATURE_PRICE_NEGOTIATION');
    let data = getSheetData(PRICE_NEGOTIATIONS_SHEET);
    if (!_isAdmin(userEmail)) {
      data = data.filter(r => r.assigned_user === userEmail);
    }
    data.sort((a, b) => new Date(b.created_at) - new Date(a.created_at));
    return successResponse(data);
  });
}

function getNegotiationsByCustomer(customerId) {
  return wrapAction(() => {
    requireFeature('FEATURE_PRICE_NEGOTIATION');
    const data = getSheetData(PRICE_NEGOTIATIONS_SHEET).filter(r => r.customer_id === customerId);
    data.sort((a, b) => new Date(b.created_at) - new Date(a.created_at));
    return successResponse(data);
  });
}

/**
 * ステータス更新（DRAFT→PROPOSED→NEGOTIATING→AGREED/REJECTED）。
 * 合意・不成立になった際は活動履歴にも自動記録する。
 */
function updateNegotiationStatus(data) {
  return wrapAction(() => {
    requireFeature('FEATURE_PRICE_NEGOTIATION');
    const valErr = validateRequired(data, ['negotiation_id', 'status']);
    if (valErr) return errorResponse(valErr.error);

    const neg = findRow(PRICE_NEGOTIATIONS_SHEET, 'negotiation_id', data.negotiation_id);
    if (!neg) return errorResponse(`交渉ID「${data.negotiation_id}」が見つかりません。`);

    const updateData = { status: data.status, updated_at: now() };
    if (data.agreed_price !== undefined && data.agreed_price !== '') updateData.agreed_price = Number(data.agreed_price);
    if (data.agreed_date) updateData.agreed_date = data.agreed_date;

    updateRow(PRICE_NEGOTIATIONS_SHEET, 'negotiation_id', data.negotiation_id, updateData);

    if (data.status === 'AGREED' || data.status === 'REJECTED') {
      const statusLabel = data.status === 'AGREED' ? '合意' : '不成立';
      createActivity({
        customer_id: neg.customer_id,
        deal_id: neg.deal_id || '',
        type: 'OTHER',
        subject: `価格改定交渉: ${neg.item_name}（${statusLabel}）`,
        content: data.status === 'AGREED'
          ? `現行価格${neg.current_price} → 合意価格${updateData.agreed_price || neg.target_price} で合意しました。`
          : `価格改定交渉（${neg.item_name}）は不成立となりました。`,
        activity_date: now(),
      });
    }

    logInfo(`価格交渉ステータス更新: ${data.negotiation_id} → ${data.status}`);
    return successResponse({ negotiation_id: data.negotiation_id, status: data.status });
  });
}

function deleteNegotiation(negotiationId) {
  return wrapAction(() => {
    requireFeature('FEATURE_PRICE_NEGOTIATION');
    const sheet = getSheet(PRICE_NEGOTIATIONS_SHEET);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIdx = headers.indexOf('negotiation_id');
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] === negotiationId) {
        sheet.deleteRow(i + 1);
        return successResponse({ deleted: true });
      }
    }
    return errorResponse(`交渉ID「${negotiationId}」が見つかりません。`);
  });
}
