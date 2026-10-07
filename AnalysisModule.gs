/**
 * AnalysisModule.gs
 * 顧客分析AI - CX研修の手法に基づく環境分析・仮説立案
 *
 * フロー:
 *   ①議事録・活動履歴の蓄積 → ②内容を抽象化してテーマ分類 → ③課題・問題意識を抽出
 *   → ④SWOT/PEST/3C分析で顧客を取り巻く環境を整理 → ⑤自社との接点を探す
 *   → ⑥課題への仮説を立てる → ⑦「自分が顧客だったらやること」を考える
 *   → ⑧提案力・クロージング力強化につながる次アクションを提示
 */

const INSIGHTS_SHEET = '🧭 Insights';

/**
 * 顧客の議事録・活動履歴・商談情報を集約し、AIで分析を生成して保存する
 */
function generateCustomerInsight(customerId) {
  return wrapAction(() => {
    requireFeature('FEATURE_INSIGHTS');

    if (!customerId) return errorResponse('顧客IDが必要です。');
    const customer = findRow(CUSTOMERS_SHEET, 'customer_id', customerId);
    if (!customer) return errorResponse(`顧客ID「${customerId}」が見つかりません。`);

    const meetings = getSheetData(MEETINGS_SHEET).filter(m => m.customer_id === customerId);
    const activities = getSheetData(ACTIVITIES_SHEET).filter(a => a.customer_id === customerId);
    const deals = getSheetData(DEALS_SHEET).filter(d => d.customer_id === customerId);

    if (meetings.length === 0 && activities.length === 0) {
      return errorResponse('分析対象となる議事録・活動履歴がまだありません。先に商談録音または活動記録を作成してください。');
    }

    const context = {
      companyName: customer.company_name,
      industry: customer.industry,
      notes: customer.notes,
      meetings: meetings.map(m => ({
        date: formatDate(m.meeting_date), title: m.title, summary: m.summary, transcript: m.transcript,
      })),
      activities: activities.map(a => ({
        date: formatDate(a.activity_date), type: a.type, subject: a.subject, content: a.content, next_action: a.next_action,
      })),
      deals: deals.map(d => ({ name: d.deal_name, phase: d.phase, amount: d.amount })),
    };

    const result = generateCustomerInsightAI(context);

    const insightId = generateId('INS');
    const row = {
      insight_id: insightId,
      customer_id: customerId,
      source_meeting_count: meetings.length,
      source_activity_count: activities.length,
      themes_json: JSON.stringify(result.themes || []),
      problem_awareness: result.problem_awareness || '',
      swot_json: JSON.stringify(result.swot || {}),
      pest_json: JSON.stringify(result.pest || {}),
      three_c_json: JSON.stringify(result.three_c || {}),
      contact_points: result.contact_points || '',
      hypotheses_json: JSON.stringify(result.hypotheses || []),
      if_i_were_customer_json: JSON.stringify(result.if_i_were_customer || []),
      proposal_actions_json: JSON.stringify((result.proposal_actions || []).map(a => ({ action: a, done: false }))),
      data_sufficiency: result.data_sufficiency || '',
      assigned_user: Session.getActiveUser().getEmail(),
      created_at: now(),
    };
    appendRow(INSIGHTS_SHEET, row);

    logInfo(`顧客分析生成: ${insightId}（顧客${customerId}、議事録${meetings.length}件・活動${activities.length}件を使用）`);
    return successResponse(row);
  });
}

function getInsightsByCustomer(customerId) {
  return wrapAction(() => {
    requireFeature('FEATURE_INSIGHTS');
    const data = getSheetData(INSIGHTS_SHEET).filter(r => r.customer_id === customerId);
    data.sort((a, b) => new Date(b.created_at) - new Date(a.created_at));
    return successResponse(data);
  });
}

function getLatestInsight(customerId) {
  return wrapAction(() => {
    requireFeature('FEATURE_INSIGHTS');
    const data = getSheetData(INSIGHTS_SHEET).filter(r => r.customer_id === customerId);
    data.sort((a, b) => new Date(b.created_at) - new Date(a.created_at));
    return successResponse(data[0] || null);
  });
}

/**
 * 提案アクションの完了チェックを切り替える
 */
function toggleInsightAction(data) {
  return wrapAction(() => {
    requireFeature('FEATURE_INSIGHTS');
    const valErr = validateRequired(data, ['insight_id']);
    if (valErr) return errorResponse(valErr.error);
    if (data.index === undefined || data.index === null) return errorResponse('アクションのインデックスが必要です。');

    const insight = findRow(INSIGHTS_SHEET, 'insight_id', data.insight_id);
    if (!insight) return errorResponse('分析が見つかりません。');

    let actions = [];
    try { actions = JSON.parse(insight.proposal_actions_json || '[]'); } catch (e) { actions = []; }

    const idx = Number(data.index);
    if (!actions[idx]) return errorResponse('アクションが見つかりません。');
    actions[idx].done = !actions[idx].done;

    updateRow(INSIGHTS_SHEET, 'insight_id', data.insight_id, { proposal_actions_json: JSON.stringify(actions) });
    return successResponse({ actions });
  });
}

function deleteInsight(insightId) {
  return wrapAction(() => {
    requireFeature('FEATURE_INSIGHTS');
    const sheet = getSheet(INSIGHTS_SHEET);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIdx = headers.indexOf('insight_id');
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] === insightId) {
        sheet.deleteRow(i + 1);
        return successResponse({ deleted: true });
      }
    }
    return errorResponse(`分析ID「${insightId}」が見つかりません。`);
  });
}
