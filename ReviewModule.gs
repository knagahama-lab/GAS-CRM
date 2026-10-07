/**
 * ReviewModule.gs
 * 営業担当者別 PDCAレビュー - 顧客分析AI（Insights）を横断集計し、
 * 「Plan（注力すべきテーマ）／Do（実施した活動）／Check（繰り返す課題・停滞）／
 *   Act（次の改善アクション）」の形で定期的に振り返りを生成する。
 *
 * また、商談・活動が新しく増えた顧客の分析（Insights）を定期的に自動更新する
 * （＝個々の顧客分析を常に最新化し、PDCAレビューの材料を鮮度高く保つ）。
 */

const PDCA_SHEET = '📈 PdcaReviews';

// ── 既存デプロイへの自動マイグレーション ─────────────────────

function ensureInsightReviewSetup() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let settingsChanged = false;

    if (!ss.getSheetByName(PDCA_SHEET)) {
      const sheet = ss.insertSheet(PDCA_SHEET);
      const headers = ['review_id','assigned_user','period_label','period_days',
        'meeting_count','activity_count','deal_count','won_count','won_amount',
        'insight_count','recurring_themes_json','action_total','action_done','action_completion_rate',
        'plan','do_text','check_text','act_json','summary_comment','created_at'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers])
        .setBackground('#1A56DB').setFontColor('#FFFFFF').setFontWeight('bold');
      sheet.autoResizeColumns(1, headers.length);
      logInfo('PdcaReviews シートを自動補完作成しました');
    }

    const settingsSheet = ss.getSheetByName(SETTINGS_SHEET);
    if (settingsSheet) {
      const data = settingsSheet.getDataRange().getValues();
      const existingKeys = data.slice(1).map(r => r[0]);
      const defaults = [
        ['FEATURE_PDCA',         'TRUE', now(), '', '営業担当者別PDCAレビュー機能'],
        ['INSIGHT_AUTO_REFRESH', 'TRUE', now(), '', '顧客分析AI：新しい議事録・活動がある顧客を定期的に自動再分析するか'],
        ['PDCA_AUTO_REVIEW',     'TRUE', now(), '', 'PDCAレビュー：月次で全担当者分を自動生成・メール送信するか'],
      ];
      const missing = defaults.filter(d => existingKeys.indexOf(d[0]) === -1);
      if (missing.length) {
        settingsSheet.getRange(settingsSheet.getLastRow() + 1, 1, missing.length, missing[0].length).setValues(missing);
        settingsChanged = true;
        logInfo(`Settings にPDCAレビュー設定キーを自動補完しました: ${missing.map(m => m[0]).join(', ')}`);
      }
    }

    if (settingsChanged) _clearCache();
  } catch (e) {
    logError('PDCAレビューの自動セットアップ補完に失敗しました', e);
  }
}

// ── PDCAレビュー生成 ──────────────────────────────────────────

function generateRepPdcaReview(data) {
  return wrapAction(() => {
    requireFeature('FEATURE_PDCA');

    const targetUser = data.assigned_user || Session.getActiveUser().getEmail();
    const periodDays = Number(data.period_days) || 30;
    const since = _daysAgo(periodDays);

    const meetings = getSheetData(MEETINGS_SHEET).filter(m => m.assigned_user === targetUser && new Date(m.created_at) >= since);
    const activities = getSheetData(ACTIVITIES_SHEET).filter(a => a.assigned_user === targetUser && new Date(a.activity_date) >= since);
    const dealsAll = getSheetData(DEALS_SHEET).filter(d => d.assigned_user === targetUser);
    const dealsInPeriod = dealsAll.filter(d => new Date(d.updated_at) >= since);
    const wonInPeriod = dealsInPeriod.filter(d => d.phase === 'CLOSED_WON');
    const insights = getSheetData(INSIGHTS_SHEET).filter(r => r.assigned_user === targetUser && new Date(r.created_at) >= since);

    if (meetings.length === 0 && activities.length === 0 && insights.length === 0) {
      return errorResponse('対象期間に活動データがありません。議事録の記録や活動ログの入力を行ってから再実行してください。');
    }

    // 顧客分析（Insights）から頻出テーマを集計（Check材料・AIに丸投げせず決定的に集計する）
    const themeCounts = {};
    insights.forEach(r => {
      let themes = [];
      try { themes = JSON.parse(r.themes_json || '[]'); } catch (e) { /* noop */ }
      themes.forEach(t => {
        const key = String((t && t.theme) || '').trim();
        if (!key) return;
        themeCounts[key] = (themeCounts[key] || 0) + 1;
      });
    });
    const recurringThemes = Object.entries(themeCounts)
      .sort((a, b) => b[1] - a[1]).slice(0, 8)
      .map(([theme, count]) => ({ theme, count }));

    // 提案アクションの完了状況を集計
    const customerNameById = {};
    getSheetData(CUSTOMERS_SHEET).forEach(c => { customerNameById[c.customer_id] = c.company_name; });

    let actionTotal = 0, actionDone = 0;
    const openActions = [];
    insights.forEach(r => {
      let actions = [];
      try { actions = JSON.parse(r.proposal_actions_json || '[]'); } catch (e) { /* noop */ }
      actions.forEach(a => {
        actionTotal++;
        if (a.done) actionDone++;
        else openActions.push(`${customerNameById[r.customer_id] || r.customer_id}: ${a.action}`);
      });
    });
    const completionRate = actionTotal > 0 ? Math.round(actionDone / actionTotal * 100) : null;

    const activityTypeBreakdown = {};
    activities.forEach(a => { activityTypeBreakdown[a.type] = (activityTypeBreakdown[a.type] || 0) + 1; });

    const customerIds = Array.from(new Set(insights.map(r => r.customer_id)));
    const repUser = findRow(USERS_SHEET, 'email', targetUser);
    const wonAmount = wonInPeriod.reduce((s, d) => s + (Number(d.amount) || 0), 0);

    const context = {
      periodLabel: `直近${periodDays}日間`,
      repName: (repUser && repUser.name) || targetUser,
      meetingCount: meetings.length,
      activityCount: activities.length,
      activityTypeBreakdown,
      dealCount: dealsInPeriod.length,
      wonCount: wonInPeriod.length,
      wonAmount,
      insightCount: insights.length,
      customerNames: customerIds.map(id => customerNameById[id] || id),
      recurringThemes,
      actionTotal, actionDone, completionRate,
      openActions: openActions.slice(0, 15),
    };

    const result = generateRepReviewAI(context);

    const reviewId = generateId('PDCA');
    const row = {
      review_id: reviewId,
      assigned_user: targetUser,
      period_label: context.periodLabel,
      period_days: periodDays,
      meeting_count: meetings.length,
      activity_count: activities.length,
      deal_count: dealsInPeriod.length,
      won_count: wonInPeriod.length,
      won_amount: wonAmount,
      insight_count: insights.length,
      recurring_themes_json: JSON.stringify(recurringThemes),
      action_total: actionTotal,
      action_done: actionDone,
      action_completion_rate: completionRate === null ? '' : completionRate,
      plan: result.plan || '',
      do_text: result.do || '',
      check_text: result.check || '',
      act_json: JSON.stringify(result.act || []),
      summary_comment: result.summary_comment || '',
      created_at: now(),
    };
    appendRow(PDCA_SHEET, row);

    logInfo(`PDCAレビュー生成: ${reviewId}（担当者${targetUser}、対象期間${periodDays}日）`);
    // レスポンスには保存していない未完了アクション一覧も併せて返す（画面表示用）
    return successResponse(Object.assign({}, row, { open_actions: openActions }));
  });
}

function getRepPdcaReviews(assignedUser) {
  return wrapAction(() => {
    requireFeature('FEATURE_PDCA');
    const data = getSheetData(PDCA_SHEET).filter(r => r.assigned_user === assignedUser);
    data.sort((a, b) => new Date(b.created_at) - new Date(a.created_at));
    return successResponse(data);
  });
}

function getLatestRepPdcaReview(assignedUser) {
  return wrapAction(() => {
    requireFeature('FEATURE_PDCA');
    const data = getSheetData(PDCA_SHEET).filter(r => r.assigned_user === assignedUser);
    data.sort((a, b) => new Date(b.created_at) - new Date(a.created_at));
    return successResponse(data[0] || null);
  });
}

// ── 定期自動実行（トリガー） ───────────────────────────────────

/**
 * 新しい議事録・活動が増えた顧客の分析（Insights）を自動更新する。
 * 1回の実行あたり最大15件に制限し、実行時間超過を防ぐ。
 */
function refreshStaleCustomerInsightsTrigger() {
  if (!isEnabled('FEATURE_INSIGHTS') || !isEnabled('INSIGHT_AUTO_REFRESH')) return;
  try {
    const customers = getSheetData(CUSTOMERS_SHEET);
    const insights = getSheetData(INSIGHTS_SHEET);
    const latestInsightByCustomer = {};
    insights.forEach(r => {
      const prev = latestInsightByCustomer[r.customer_id];
      if (!prev || new Date(r.created_at) > new Date(prev.created_at)) latestInsightByCustomer[r.customer_id] = r;
    });

    const meetings = getSheetData(MEETINGS_SHEET);
    const activities = getSheetData(ACTIVITIES_SHEET);
    const MAX_PER_RUN = 15;
    let processed = 0;

    for (let i = 0; i < customers.length && processed < MAX_PER_RUN; i++) {
      const c = customers[i];
      const latest = latestInsightByCustomer[c.customer_id];
      const lastInsightDate = latest ? new Date(latest.created_at) : null;

      const hasNewMeeting = meetings.some(m => m.customer_id === c.customer_id && (!lastInsightDate || new Date(m.created_at) > lastInsightDate));
      const hasNewActivity = activities.some(a => a.customer_id === c.customer_id && (!lastInsightDate || new Date(a.created_at) > lastInsightDate));
      if (!hasNewMeeting && !hasNewActivity) continue;

      try {
        generateCustomerInsight(c.customer_id);
        processed++;
      } catch (e) {
        logError(`顧客分析の自動更新に失敗: ${c.customer_id}`, e);
      }
    }
    logInfo(`顧客分析の自動更新完了: ${processed}件`);
  } catch (e) {
    logError('顧客分析の自動更新処理に失敗しました', e);
  }
}

/**
 * 全アクティブユーザー分のPDCAレビューを月次で自動生成し、本人にメール送信する。
 */
function generateMonthlyPdcaReviewsTrigger() {
  if (!isEnabled('FEATURE_PDCA') || !isEnabled('PDCA_AUTO_REVIEW')) return;
  try {
    const users = getSheetData(USERS_SHEET).filter(u => String(u.active).toUpperCase() === 'TRUE');
    users.forEach(u => {
      try {
        const res = JSON.parse(generateRepPdcaReview({ assigned_user: u.email, period_days: 30 }));
        if (res.success) _emailPdcaReview(u, res.data);
      } catch (e) {
        logError(`月次PDCAレビュー生成に失敗: ${u.email}`, e);
      }
    });
    logInfo('月次PDCAレビュー一括生成完了');
  } catch (e) {
    logError('月次PDCAレビュー一括生成処理に失敗しました', e);
  }
}

function _emailPdcaReview(user, review) {
  try {
    let actList = [];
    try { actList = JSON.parse(review.act_json || '[]'); } catch (e) { /* noop */ }
    const html = `
      <div style="font-family:Arial,sans-serif;max-width:700px;color:#1F2937">
        <h2 style="color:#1A56DB">📈 月次PDCAレビュー（${sanitizeString(review.period_label)}）</h2>
        <p>${sanitizeString(review.summary_comment)}</p>
        <h3>Plan（注力すべきテーマ）</h3><p>${sanitizeString(review.plan)}</p>
        <h3>Do（実施した活動）</h3><p>${sanitizeString(review.do_text)}</p>
        <h3>Check（繰り返す課題・進捗）</h3><p>${sanitizeString(review.check_text)}</p>
        <h3>Act（次の改善アクション）</h3>
        <ul>${actList.map(a => `<li>${sanitizeString(a)}</li>`).join('')}</ul>
        <p style="color:#9CA3AF;font-size:12px;margin-top:24px">このメールはCRMのPDCAレビュー自動生成機能から送信されました。詳細はCRM画面「📈 PDCAレビュー」からご確認ください。</p>
      </div>`;
    GmailApp.sendEmail(user.email, `【CRM】月次PDCAレビュー（${review.period_label}）`, '', { htmlBody: html });
  } catch (e) {
    logError(`PDCAレビューメール送信失敗: ${user.email}`, e);
  }
}

function _daysAgo(days) {
  const d = new Date();
  d.setDate(d.getDate() - days);
  return d;
}
