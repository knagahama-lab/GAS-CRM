/**
 * ReportModule.gs
 * KPI集計・定期レポート自動送信
 */

const REPORTS_SHEET = '📊 Reports';

/**
 * 定期レポート生成・送信（トリガーで実行）
 */
function generateWeeklyReport() {
  if (!isEnabled('FEATURE_REPORT_AUTO')) return;
  _generateAndSendReport('WEEKLY');
}

function generateMonthlyReport() {
  if (!isEnabled('FEATURE_REPORT_AUTO')) return;
  _generateAndSendReport('MONTHLY');
}

function _generateAndSendReport(type) {
  logInfo(`${type}レポート生成開始`);

  const recipients = getSetting('REPORT_RECIPIENTS', '');
  if (!recipients) { logInfo('送信先未設定のためスキップ'); return; }

  const kpi = JSON.parse(getKpiSummary());
  const pipeline = JSON.parse(getPipelineSummary());
  if (!kpi.success || !pipeline.success) return;

  const subject = `【CRM】${type === 'WEEKLY' ? '週次' : '月次'}営業レポート（${formatDate(new Date())}）`;
  const htmlBody = _buildReportHtml(kpi.data, pipeline.data, type);

  recipients.split(',').forEach(email => {
    email = email.trim();
    if (!email) return;
    try {
      GmailApp.sendEmail(email, subject, '', { htmlBody });
      logInfo(`レポート送信: ${email}`);
    } catch (e) {
      logError(`レポート送信失敗: ${email}`, e);
    }
  });

  appendRow(REPORTS_SHEET, {
    report_id: generateId('RPT'),
    type,
    period: formatDate(new Date()),
    generated_at: now(),
    sent_to: recipients,
    status: 'SENT',
  });
}

function getKpiSummary() {
  return wrapAction(() => {
    const customers = getSheetData('👥 Customers');
    const deals = getSheetData('💼 Deals');
    const activities = getSheetData('📝 Activities');

    const today = new Date();
    const monthStart = new Date(today.getFullYear(), today.getMonth(), 1);
    const lastMonthStart = new Date(today.getFullYear(), today.getMonth() - 1, 1);
    const lastMonthEnd = new Date(today.getFullYear(), today.getMonth(), 0, 23, 59, 59);

    const activeDeals = deals.filter(d => d.phase !== 'CLOSED_WON' && d.phase !== 'CLOSED_LOST');
    const wonThisMonth = deals.filter(d => {
      if (d.phase !== 'CLOSED_WON') return false;
      const ud = new Date(d.updated_at);
      return ud >= monthStart;
    });
    const wonLastMonth = deals.filter(d => {
      if (d.phase !== 'CLOSED_WON') return false;
      const ud = new Date(d.updated_at);
      return ud >= lastMonthStart && ud <= lastMonthEnd;
    });
    const activitiesThisMonth = activities.filter(a => new Date(a.activity_date) >= monthStart);

    const wonAmountThisMonth = wonThisMonth.reduce((s, d) => s + (Number(d.amount) || 0), 0);
    const wonAmountLastMonth = wonLastMonth.reduce((s, d) => s + (Number(d.amount) || 0), 0);
    const revenueGrowthRate = wonAmountLastMonth > 0
      ? Math.round((wonAmountThisMonth - wonAmountLastMonth) / wonAmountLastMonth * 100)
      : null;

    const totalDeals = deals.length;
    const wonDeals = deals.filter(d => d.phase === 'CLOSED_WON').length;
    const closedDeals = deals.filter(d => d.phase === 'CLOSED_WON' || d.phase === 'CLOSED_LOST').length;
    const winRate = closedDeals > 0 ? Math.round(wonDeals / closedDeals * 100) : 0;

    return successResponse({
      totalCustomers: customers.length,
      activeDeals: activeDeals.length,
      pipelineAmount: activeDeals.reduce((s, d) => s + (Number(d.amount) || 0), 0),
      weightedAmount: activeDeals.reduce((s, d) => s + (Number(d.weighted_amount) || 0), 0),
      wonThisMonth: wonThisMonth.length,
      wonAmountThisMonth,
      wonAmountLastMonth,
      revenueGrowthRate,
      activitiesThisMonth: activitiesThisMonth.length,
      winRate,
      totalDeals,
    });
  });
}

/**
 * 過去N月の月別売上時系列データ（グラフ用）
 */
function getSalesTrend(months) {
  return wrapAction(() => {
    const n = Math.max(1, Math.min(Number(months) || 6, 24));
    const deals = getSheetData('💼 Deals');
    const today = new Date();
    const result = [];

    for (let i = n - 1; i >= 0; i--) {
      const start = new Date(today.getFullYear(), today.getMonth() - i, 1);
      const end   = new Date(today.getFullYear(), today.getMonth() - i + 1, 0, 23, 59, 59);
      const label = Utilities.formatDate(start, 'Asia/Tokyo', 'yyyy/MM');

      const wonDeals = deals.filter(d => {
        if (d.phase !== 'CLOSED_WON') return false;
        const ud = new Date(d.updated_at);
        return ud >= start && ud <= end;
      });
      result.push({
        month: label,
        count: wonDeals.length,
        amount: wonDeals.reduce((s, d) => s + (Number(d.amount) || 0), 0),
      });
    }
    return successResponse(result);
  });
}

/**
 * 担当者別件数・売上ランキング
 */
function getTeamPerformance() {
  return wrapAction(() => {
    const deals = getSheetData('💼 Deals');
    const today = new Date();
    const monthStart = new Date(today.getFullYear(), today.getMonth(), 1);

    const byUser = {};
    deals.forEach(d => {
      const user = d.assigned_user || 'unknown';
      if (!byUser[user]) {
        byUser[user] = { assigned_user: user, totalDeals: 0, wonDeals: 0, wonAmount: 0, wonThisMonth: 0, wonAmountThisMonth: 0 };
      }
      byUser[user].totalDeals++;
      if (d.phase === 'CLOSED_WON') {
        byUser[user].wonDeals++;
        byUser[user].wonAmount += Number(d.amount) || 0;
        const ud = new Date(d.updated_at);
        if (ud >= monthStart) {
          byUser[user].wonThisMonth++;
          byUser[user].wonAmountThisMonth += Number(d.amount) || 0;
        }
      }
    });

    const ranking = Object.values(byUser).sort((a, b) => b.wonAmountThisMonth - a.wonAmountThisMonth);
    return successResponse(ranking);
  });
}

function updateDashboard() {
  if (!isEnabled('FEATURE_DASHBOARD')) return;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('🏠 Dashboard');
    if (!sheet) return;
    sheet.getRange('B2').setValue(now());
    logInfo('ダッシュボード更新完了');
  } catch (e) {
    logError('ダッシュボード更新失敗', e);
  }
}

function _buildReportHtml(kpi, pipeline, type) {
  const label = type === 'WEEKLY' ? '週次' : '月次';

  const phaseLabels = {
    PROSPECT:'見込み', APPROACH:'アプローチ', PROPOSAL:'提案',
    NEGOTIATION:'交渉', CLOSED_WON:'受注', CLOSED_LOST:'失注',
  };

  let pipelineRows = '';
  Object.entries(pipeline).forEach(([phase, data]) => {
    pipelineRows += `<tr>
      <td>${phaseLabels[phase] || phase}</td>
      <td style="text-align:right">${data.count}</td>
      <td style="text-align:right">¥${data.total_amount.toLocaleString()}</td>
      <td style="text-align:right">¥${data.weighted_amount.toLocaleString()}</td>
    </tr>`;
  });

  return `
<!DOCTYPE html><html><head><meta charset="UTF-8">
<style>
  body{font-family:Arial,sans-serif;max-width:800px;margin:0 auto;padding:20px;color:#1F2937}
  h1{color:#1A56DB;border-bottom:3px solid #1A56DB;padding-bottom:8px}
  h2{color:#374151;margin-top:30px}
  .kpi-grid{display:grid;grid-template-columns:repeat(2,1fr);gap:16px;margin:20px 0}
  .kpi-card{background:#EFF6FF;border-left:4px solid #1A56DB;padding:16px;border-radius:4px}
  .kpi-value{font-size:28px;font-weight:bold;color:#1A56DB}
  .kpi-label{color:#6B7280;font-size:14px;margin-top:4px}
  table{border-collapse:collapse;width:100%;margin-top:12px}
  th{background:#1A56DB;color:white;padding:8px 12px;text-align:left}
  td{padding:8px 12px;border-bottom:1px solid #E5E7EB}
  tr:nth-child(even){background:#F9FAFB}
  .footer{margin-top:40px;color:#9CA3AF;font-size:12px;border-top:1px solid #E5E7EB;padding-top:12px}
</style></head><body>
<h1>📊 ${label}営業レポート</h1>
<p>集計日時: ${formatDateTime(new Date())}</p>

<h2>📈 主要KPI</h2>
<div class="kpi-grid">
  <div class="kpi-card"><div class="kpi-value">${kpi.totalCustomers}</div><div class="kpi-label">👥 総顧客数</div></div>
  <div class="kpi-card"><div class="kpi-value">${kpi.activeDeals}</div><div class="kpi-label">💼 進行中商談数</div></div>
  <div class="kpi-card"><div class="kpi-value">¥${kpi.wonAmountThisMonth.toLocaleString()}</div><div class="kpi-label">✅ 今月受注金額</div></div>
  <div class="kpi-card"><div class="kpi-value">${kpi.activitiesThisMonth}</div><div class="kpi-label">📝 今月活動数</div></div>
  <div class="kpi-card"><div class="kpi-value">¥${kpi.pipelineAmount.toLocaleString()}</div><div class="kpi-label">💰 パイプライン合計</div></div>
  <div class="kpi-card"><div class="kpi-value">¥${kpi.weightedAmount.toLocaleString()}</div><div class="kpi-label">⚖️ 加重合計</div></div>
</div>

<h2>🔄 パイプライン状況</h2>
<table>
  <tr><th>フェーズ</th><th>件数</th><th>合計金額</th><th>加重金額</th></tr>
  ${pipelineRows}
</table>

<div class="footer">このメールはCRM自動レポートシステムから送信されました。</div>
</body></html>`;
}
