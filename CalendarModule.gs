/**
 * CalendarModule.gs
 * Google カレンダー連携
 */

const CALENDAR_SHEET = '📅 Calendar';

function createDealCalendarEvent(deal) {
  if (!isEnabled('FEATURE_CAL_CREATE')) return null;

  const calendar = CalendarApp.getDefaultCalendar();
  const closeDate = new Date(deal.close_date);
  closeDate.setHours(10, 0, 0, 0);
  const endDate = new Date(closeDate.getTime() + 60 * 60 * 1000);

  const event = calendar.createEvent(
    `[商談] ${deal.deal_name}（${deal.company_name}）`,
    closeDate, endDate,
    {
      description: [
        `商談ID: ${deal.deal_id}`,
        `顧客: ${deal.company_name}`,
        `金額: ${Number(deal.amount).toLocaleString()}円`,
        `確度: ${deal.probability}%`,
        `担当: ${deal.assigned_user}`,
        `メモ: ${deal.notes || ''}`,
      ].join('\n'),
    }
  );

  // Calendarシートに記録
  appendRow(CALENDAR_SHEET, {
    event_id: event.getId(),
    deal_id: deal.deal_id,
    customer_id: deal.customer_id,
    type: 'MEETING',
    title: `[商談] ${deal.deal_name}`,
    start_datetime: closeDate,
    end_datetime: endDate,
    description: deal.notes || '',
    assigned_user: deal.assigned_user,
    created_at: now(),
  });

  logInfo(`カレンダー登録: ${event.getId()} - ${deal.deal_name}`);
  return event.getId();
}

/**
 * 毎朝リマインダー送信（トリガーで実行）
 */
function sendDailyReminders() {
  if (!isEnabled('FEATURE_CAL_REMINDER')) return;

  logInfo('リマインダー送信開始');
  const days = Number(getSetting('REMINDER_DAYS_BEFORE', 3));
  const dealsResult = JSON.parse(getDealsNearClose(days));
  if (!dealsResult.success || !dealsResult.data.length) return;

  const dealsByUser = {};
  dealsResult.data.forEach(deal => {
    const u = deal.assigned_user;
    if (!dealsByUser[u]) dealsByUser[u] = [];
    dealsByUser[u].push(deal);
  });

  Object.entries(dealsByUser).forEach(([email, deals]) => {
    if (!email) return;
    try {
      _sendReminderEmail(email, deals);
      logInfo(`リマインダー送信: ${email} - ${deals.length}件`);
    } catch (e) {
      logError(`リマインダー送信失敗: ${email}`, e);
    }
  });

  // 次回アクションのリマインダー
  _sendNextActionReminders();
}

function _sendReminderEmail(email, deals) {
  const subject = `【CRM】クローズ予定の商談リマインダー（${deals.length}件）`;
  let body = `<h2>クローズ予定の商談リマインダー</h2>`;
  body += `<p>以下の商談のクローズ予定日が近づいています。</p>`;
  body += `<table border="1" cellpadding="6" style="border-collapse:collapse;width:100%">`;
  body += `<tr style="background:#1A56DB;color:white"><th>商談名</th><th>顧客名</th><th>金額</th><th>確度</th><th>クローズ予定日</th></tr>`;
  deals.forEach(d => {
    body += `<tr>
      <td>${d.deal_name}</td>
      <td>${d.company_name}</td>
      <td>¥${Number(d.amount).toLocaleString()}</td>
      <td>${d.probability}%</td>
      <td>${formatDate(d.close_date)}</td>
    </tr>`;
  });
  body += `</table>`;
  body += `<p style="color:#666;font-size:12px">このメールはCRMシステムから自動送信されています。</p>`;

  GmailApp.sendEmail(email, subject, '', { htmlBody: body });
}

function _sendNextActionReminders() {
  const users = getSheetData('👤 Users').filter(u => u.active === 'TRUE' || u.active === true);
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const tomorrow = new Date(today.getTime() + 24 * 60 * 60 * 1000);

  users.forEach(user => {
    const result = JSON.parse(getPendingNextActions(user.email));
    if (!result.success) return;
    const todayActions = result.data.filter(a => {
      const d = new Date(a.next_action_date);
      return d >= today && d < tomorrow;
    });
    if (!todayActions.length) return;

    try {
      let body = `<h2>本日のフォローアップアクション（${todayActions.length}件）</h2><ul>`;
      todayActions.forEach(a => {
        body += `<li><strong>${a.next_action}</strong>（顧客ID: ${a.customer_id}）<br>${a.content}</li>`;
      });
      body += `</ul>`;
      GmailApp.sendEmail(user.email, '【CRM】本日のフォローアップリスト', '', { htmlBody: body });
    } catch (e) {
      logError(`フォローリマインダー失敗: ${user.email}`, e);
    }
  });
}
