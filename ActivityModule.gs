/**
 * ActivityModule.gs
 * 営業活動ログ記録・次回アクション管理
 */

const ACTIVITIES_SHEET = '📝 Activities';

function createActivity(data) {
  requireFeature('FEATURE_ACTIVITIES');

  const id = generateId('ACT');
  const row = {
    activity_id: id,
    customer_id: data.customer_id || '',
    deal_id: data.deal_id || '',
    type: data.type || 'OTHER',
    activity_date: data.activity_date || now(),
    subject: data.subject || '',
    content: data.content || '',
    next_action: data.next_action || '',
    next_action_date: data.next_action_date || '',
    assigned_user: data.assigned_user || Session.getActiveUser().getEmail(),
    created_at: now(),
  };

  appendRow(ACTIVITIES_SHEET, row);

  // 顧客の最終コンタクト日を更新
  if (data.customer_id) {
    updateRow('👥 Customers', 'customer_id', data.customer_id, {
      last_contact_date: row.activity_date,
      updated_at: now(),
    });
  }

  // カレンダー登録（次回アクション日がある場合）
  if (isEnabled('FEATURE_CAL_CREATE') && data.next_action_date && data.next_action) {
    try {
      _createActivityCalendarEvent(row);
    } catch (e) {
      logError('活動カレンダー登録失敗', e);
    }
  }

  logInfo(`活動記録: ${id} - ${data.subject}`);
  return successResponse({ activity_id: id });
}

function getActivitiesByCustomer(customerId) {
  requireFeature('FEATURE_ACTIVITIES');
  const data = getSheetData(ACTIVITIES_SHEET);
  return successResponse(
    data.filter(r => r.customer_id === customerId)
      .sort((a, b) => new Date(b.activity_date) - new Date(a.activity_date))
  );
}

function getActivitiesByDeal(dealId) {
  requireFeature('FEATURE_ACTIVITIES');
  const data = getSheetData(ACTIVITIES_SHEET);
  return successResponse(
    data.filter(r => r.deal_id === dealId)
      .sort((a, b) => new Date(b.activity_date) - new Date(a.activity_date))
  );
}

function getRecentActivities(limit, userEmail) {
  requireFeature('FEATURE_ACTIVITIES');
  let data = getSheetData(ACTIVITIES_SHEET);
  if (!_isAdmin(userEmail)) {
    data = data.filter(r => r.assigned_user === userEmail);
  }
  data.sort((a, b) => new Date(b.activity_date) - new Date(a.activity_date));
  return successResponse(data.slice(0, limit || 20));
}

function getPendingNextActions(userEmail) {
  requireFeature('FEATURE_ACTIVITIES');
  let data = getSheetData(ACTIVITIES_SHEET);
  const today = new Date();
  if (!_isAdmin(userEmail)) {
    data = data.filter(r => r.assigned_user === userEmail);
  }
  const pending = data.filter(r => {
    if (!r.next_action_date || !r.next_action) return false;
    return new Date(r.next_action_date) >= today;
  });
  pending.sort((a, b) => new Date(a.next_action_date) - new Date(b.next_action_date));
  return successResponse(pending);
}

function getActivityStats(userEmail) {
  let data = getSheetData(ACTIVITIES_SHEET);
  if (!_isAdmin(userEmail)) {
    data = data.filter(r => r.assigned_user === userEmail);
  }
  const thisMonth = data.filter(r => {
    const d = new Date(r.activity_date);
    return d.getMonth() === new Date().getMonth() && d.getFullYear() === new Date().getFullYear();
  });
  const byType = {};
  thisMonth.forEach(r => { byType[r.type] = (byType[r.type] || 0) + 1; });
  return successResponse({ total: data.length, thisMonth: thisMonth.length, byType });
}

function _createActivityCalendarEvent(activity) {
  const calendar = CalendarApp.getDefaultCalendar();
  const start = new Date(activity.next_action_date);
  const end = new Date(start.getTime() + 60 * 60 * 1000); // 1時間
  calendar.createEvent(
    `[フォロー] ${activity.next_action}`,
    start, end,
    { description: `活動ID: ${activity.activity_id}\n顧客ID: ${activity.customer_id}\n内容: ${activity.content}` }
  );
}
