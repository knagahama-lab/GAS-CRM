/**
 * AnalyticsModule.gs
 * 分析ラボ - キーエンスKI風の分析機能
 *   ① 要因ツリー: 指標を軸でドリルダウン分解
 *   ② マトリックス: 行×列のクロス集計（ピボット）
 *   ③ ターゲットリスト: 条件に合う顧客・商談をワンクリック抽出
 *   ④ AIアシスト: 自然文の質問にCRM集計データを根拠に回答
 *
 * 集計（件数・金額・頻出テーマなど）は全てこのファイル側で決定的に計算し、
 * AIには「解釈・自然文化」のみを担わせる（数値のハルシネーション防止）。
 */

// ── 既存デプロイへの自動マイグレーション ─────────────────────
// 分析ラボは新規シートを必要としない（既存シートの集計のみ）ため、設定キーの補完のみ行う。

function ensureAnalyticsSetup() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName(SETTINGS_SHEET);
    if (!settingsSheet) return;

    const data = settingsSheet.getDataRange().getValues();
    const existingKeys = data.slice(1).map(r => r[0]);
    const defaults = [
      ['FEATURE_ANALYTICS', 'TRUE', now(), '', '分析ラボ（要因ツリー・マトリックス・ターゲットリスト・AIアシスト）'],
    ];
    const missing = defaults.filter(d => existingKeys.indexOf(d[0]) === -1);
    if (missing.length) {
      settingsSheet.getRange(settingsSheet.getLastRow() + 1, 1, missing.length, missing[0].length).setValues(missing);
      _clearCache();
      logInfo(`Settings に分析ラボ設定キーを自動補完しました: ${missing.map(m => m[0]).join(', ')}`);
    }
  } catch (e) {
    logError('分析ラボの自動セットアップ補完に失敗しました', e);
  }
}

// ── 共通: 分解軸アクセサ ───────────────────────────────────────

function _dimensionAccessor(dimension) {
  if (dimension === 'phase') return r => r.phase;
  if (dimension === 'assigned_user') return r => r.assigned_user;
  if (dimension === 'activity_type') return r => r.type;
  if (dimension === 'month') return r => {
    const d = r.activity_date || r.created_at;
    return d ? formatDate(d).slice(0, 7) : '';
  };
  if (dimension === 'industry' || dimension === 'customer_status') {
    const customers = getSheetData(CUSTOMERS_SHEET);
    const map = {};
    customers.forEach(c => { map[c.customer_id] = dimension === 'industry' ? c.industry : c.status; });
    return r => map[r.customer_id] || '';
  }
  return r => r[dimension] || '';
}

function _applyFilters(data, filters) {
  if (!filters || !filters.length) return data;
  const accessors = filters.map(f => ({ fn: _dimensionAccessor(f.dimension), value: f.value }));
  return data.filter(r => accessors.every(a => String(a.fn(r) || '') === String(a.value)));
}

// ── ① 要因ツリー ────────────────────────────────────────────

function getFactorBreakdown(params) {
  return wrapAction(() => {
    requireFeature('FEATURE_ANALYTICS');

    const metric = params.metric || 'deal_amount';
    const dimension = params.dimension || 'phase';
    const filters = params.filters || [];

    const sourceSheet = metric === 'activity_count' ? ACTIVITIES_SHEET : DEALS_SHEET;
    let data = getSheetData(sourceSheet);
    data = _applyFilters(data, filters);

    const dimFn = _dimensionAccessor(dimension);
    const groups = {};
    data.forEach(r => {
      const key = String(dimFn(r) || '(未設定)');
      if (!groups[key]) groups[key] = { label: key, count: 0, amount: 0, wonAmount: 0 };
      groups[key].count++;
      groups[key].amount += Number(r.amount) || 0;
      if (r.phase === 'CLOSED_WON') groups[key].wonAmount += Number(r.amount) || 0;
    });

    const metricKey = (metric === 'deal_count' || metric === 'activity_count') ? 'count' : (metric === 'won_amount' ? 'wonAmount' : 'amount');
    const items = Object.values(groups)
      .map(g => ({ label: g.label, value: g[metricKey], count: g.count }))
      .sort((a, b) => b.value - a.value);

    const total = items.reduce((s, r) => s + r.value, 0);
    return successResponse({ items, total, metric, dimension });
  });
}

// ── ② マトリックス（クロス集計） ─────────────────────────────

function getPivotMatrix(params) {
  return wrapAction(() => {
    requireFeature('FEATURE_ANALYTICS');

    const sourceSheet = params.sheet === 'activities' ? ACTIVITIES_SHEET : DEALS_SHEET;
    const rowDim = params.rowDim || 'phase';
    const colDim = params.colDim || 'assigned_user';
    const metric = params.metric || 'count';

    const data = getSheetData(sourceSheet);
    const rowFn = _dimensionAccessor(rowDim);
    const colFn = _dimensionAccessor(colDim);

    const rowKeysSet = new Set(), colKeysSet = new Set();
    const cellMap = {};
    data.forEach(r => {
      const rk = String(rowFn(r) || '(未設定)');
      const ck = String(colFn(r) || '(未設定)');
      rowKeysSet.add(rk); colKeysSet.add(ck);
      const key = rk + '\u0001' + ck;
      if (!cellMap[key]) cellMap[key] = { count: 0, amount: 0, wonAmount: 0 };
      cellMap[key].count++;
      cellMap[key].amount += Number(r.amount) || 0;
      if (r.phase === 'CLOSED_WON') cellMap[key].wonAmount += Number(r.amount) || 0;
    });

    const rowKeys = Array.from(rowKeysSet).sort();
    const colKeys = Array.from(colKeysSet).sort();
    const metricKey = metric === 'amount' ? 'amount' : (metric === 'won_amount' ? 'wonAmount' : 'count');

    const matrix = rowKeys.map(rk => colKeys.map(ck => {
      const cell = cellMap[rk + '\u0001' + ck];
      return cell ? cell[metricKey] : 0;
    }));
    const rowTotals = matrix.map(row => row.reduce((s, v) => s + v, 0));
    const colTotals = colKeys.map((_, ci) => matrix.reduce((s, row) => s + row[ci], 0));
    const grandTotal = rowTotals.reduce((s, v) => s + v, 0);

    return successResponse({ rowKeys, colKeys, matrix, rowTotals, colTotals, grandTotal, metric, rowDim, colDim });
  });
}

// ── ③ ターゲットリスト ──────────────────────────────────────

const TARGET_STATUS_LABELS = { PROSPECT: '見込み客', ACTIVE: '取引中', INACTIVE: '休眠' };

function generateTargetList(params) {
  return wrapAction(() => {
    requireFeature('FEATURE_ANALYTICS');

    const preset = params.preset;
    const days = Number(params.days) || 14;
    const today = new Date();

    const customers = getSheetData(CUSTOMERS_SHEET);
    const customerById = {};
    customers.forEach(c => { customerById[c.customer_id] = c; });
    const deals = getSheetData(DEALS_SHEET);
    const activities = getSheetData(ACTIVITIES_SHEET);

    let items = [];

    if (preset === 'STALLED') {
      const activeDeals = deals.filter(d => d.phase !== 'CLOSED_WON' && d.phase !== 'CLOSED_LOST');
      const lastActivityByDeal = {};
      activities.forEach(a => {
        if (!a.deal_id) return;
        const d = new Date(a.activity_date);
        if (!lastActivityByDeal[a.deal_id] || d > lastActivityByDeal[a.deal_id]) lastActivityByDeal[a.deal_id] = d;
      });
      const found = [];
      activeDeals.forEach(d => {
        const last = lastActivityByDeal[d.deal_id] || new Date(d.created_at);
        const diffDays = Math.floor((today - last) / 86400000);
        if (diffDays >= days) found.push({ d, diffDays });
      });
      found.sort((a, b) => Number(b.d.amount) - Number(a.d.amount));
      items = found.map(({ d, diffDays }) => ({
        customer_id: d.customer_id, company_name: d.company_name,
        deal_id: d.deal_id, deal_name: d.deal_name,
        reason: `${diffDays}日間 活動記録なし`,
        detail: `フェーズ: ${d.phase} / 金額: ¥${Number(d.amount || 0).toLocaleString()}`,
      }));

    } else if (preset === 'CLOSING_SOON') {
      const limit = new Date(today.getTime() + days * 86400000);
      const found = [];
      deals.filter(d => d.close_date && d.phase !== 'CLOSED_WON' && d.phase !== 'CLOSED_LOST').forEach(d => {
        const cd = new Date(d.close_date);
        if (cd >= today && cd <= limit) found.push({ d, cd });
      });
      found.sort((a, b) => a.cd - b.cd);
      items = found.map(({ d, cd }) => ({
        customer_id: d.customer_id, company_name: d.company_name,
        deal_id: d.deal_id, deal_name: d.deal_name,
        reason: `クローズ予定日まで${Math.ceil((cd - today) / 86400000)}日`,
        detail: `金額: ¥${Number(d.amount || 0).toLocaleString()} / 確度: ${d.probability}%`,
      }));

    } else if (preset === 'DORMANT') {
      const found = customers
        .filter(c => c.status === 'ACTIVE' || c.status === 'PROSPECT')
        .map(c => {
          const last = c.last_contact_date ? new Date(c.last_contact_date) : null;
          const diffDays = last ? Math.floor((today - last) / 86400000) : null;
          return { c, diffDays };
        })
        .filter(x => x.diffDays === null || x.diffDays >= days);
      found.sort((a, b) => (b.diffDays === null ? 99999 : b.diffDays) - (a.diffDays === null ? 99999 : a.diffDays));
      items = found.map(({ c, diffDays }) => ({
        customer_id: c.customer_id, company_name: c.company_name,
        deal_id: '', deal_name: '',
        reason: diffDays === null ? 'コンタクト履歴なし' : `${diffDays}日間 未接触`,
        detail: `ステータス: ${TARGET_STATUS_LABELS[c.status] || c.status} / 担当: ${c.assigned_user}`,
      }));

    } else if (preset === 'HIGH_RISK') {
      const activeDeals = deals.filter(d => d.phase !== 'CLOSED_WON' && d.phase !== 'CLOSED_LOST');
      const amounts = activeDeals.map(d => Number(d.amount) || 0).sort((a, b) => a - b);
      const p75 = amounts.length ? amounts[Math.floor(amounts.length * 0.75)] : 0;
      const threshold = Math.max(p75, 500000);
      const found = activeDeals.filter(d => Number(d.amount || 0) >= threshold && Number(d.probability || 0) <= 30);
      found.sort((a, b) => Number(b.amount) - Number(a.amount));
      items = found.map(d => ({
        customer_id: d.customer_id, company_name: d.company_name,
        deal_id: d.deal_id, deal_name: d.deal_name,
        reason: `高額（¥${Number(d.amount).toLocaleString()}）なのに確度${d.probability}%と低い`,
        detail: `フェーズ: ${d.phase}`,
      }));

    } else if (preset === 'OVERDUE_ACTION') {
      const found = activities.filter(a => a.next_action && a.next_action_date && new Date(a.next_action_date) < today);
      found.sort((a, b) => new Date(a.next_action_date) - new Date(b.next_action_date));
      items = found.map(a => {
        const c = customerById[a.customer_id];
        const overdueDays = Math.floor((today - new Date(a.next_action_date)) / 86400000);
        return {
          customer_id: a.customer_id, company_name: c ? c.company_name : a.customer_id,
          deal_id: a.deal_id, deal_name: '',
          reason: `予定日を${overdueDays}日超過: ${a.next_action}`,
          detail: `担当: ${a.assigned_user}`,
        };
      });

    } else {
      return errorResponse(`不明なプリセット: ${preset}`);
    }

    return successResponse({ preset, items: items.slice(0, 100) });
  });
}

// ── ④ AIアシスト ────────────────────────────────────────────

function runAiAssistantQuery(question) {
  return wrapAction(() => {
    requireFeature('FEATURE_ANALYTICS');
    if (!question || !String(question).trim()) return errorResponse('質問を入力してください。');

    const kpi = JSON.parse(getKpiSummary());
    const pipeline = JSON.parse(getPipelineSummary());
    const team = JSON.parse(getTeamPerformance());
    const recentActivities = getSheetData(ACTIVITIES_SHEET)
      .sort((a, b) => new Date(b.activity_date) - new Date(a.activity_date)).slice(0, 30);
    const nearClose = JSON.parse(getDealsNearClose(14));
    const insights = getSheetData(INSIGHTS_SHEET).slice(-50);

    const themeCounts = {};
    insights.forEach(r => {
      let themes = [];
      try { themes = JSON.parse(r.themes_json || '[]'); } catch (e) { /* noop */ }
      themes.forEach(t => {
        const key = String((t && t.theme) || '').trim();
        if (key) themeCounts[key] = (themeCounts[key] || 0) + 1;
      });
    });
    const recurringThemes = Object.entries(themeCounts).sort((a, b) => b[1] - a[1]).slice(0, 10)
      .map(([theme, count]) => ({ theme, count }));

    const context = {
      kpi: kpi.data,
      pipeline: pipeline.data,
      team: team.data,
      recentActivities: recentActivities.map(a => ({ date: formatDate(a.activity_date), type: a.type, subject: a.subject, assigned_user: a.assigned_user })),
      nearClose: (nearClose.data || []).map(d => ({ name: d.deal_name, company: d.company_name, amount: d.amount, closeDate: d.close_date, assignedUser: d.assigned_user })),
      recurringThemes,
    };

    const result = askCrmAssistantAI(context, question);
    logInfo(`AIアシスト質問: ${question}`);
    return successResponse(result);
  });
}
