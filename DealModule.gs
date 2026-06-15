/**
 * DealModule.gs
 * 商談管理・フェーズ管理・パイプライン集計
 */

const DEALS_SHEET = '💼 Deals';

function createDeal(data) {
  requireFeature('FEATURE_DEALS');

  // 顧客名を取得
  let companyName = data.company_name || '';
  if (!companyName && data.customer_id) {
    const cust = JSON.parse(getCustomer(data.customer_id));
    if (cust.success) companyName = cust.data.company_name;
  }

  const id = generateId('DEAL');
  const weighted = (Number(data.amount) || 0) * (Number(data.probability) || 0) / 100;

  const row = {
    deal_id: id,
    customer_id: data.customer_id || '',
    company_name: companyName,
    deal_name: data.deal_name || '',
    phase: data.phase || 'PROSPECT',
    amount: Number(data.amount) || 0,
    probability: Number(data.probability) || 0,
    weighted_amount: Math.round(weighted),
    close_date: data.close_date || '',
    assigned_user: data.assigned_user || Session.getActiveUser().getEmail(),
    calendar_event_id: '',
    notes: data.notes || '',
    created_at: now(),
    updated_at: now(),
  };

  appendRow(DEALS_SHEET, row);

  // カレンダー自動登録
  if (isEnabled('FEATURE_CAL_CREATE') && data.close_date) {
    try {
      const eventId = createDealCalendarEvent(row);
      updateRow(DEALS_SHEET, 'deal_id', id, { calendar_event_id: eventId });
      row.calendar_event_id = eventId;
    } catch (e) {
      logError('カレンダー登録失敗', e);
    }
  }

  logInfo(`商談作成: ${id} - ${data.deal_name}`);
  return successResponse({ deal_id: id });
}

function updateDeal(id, data) {
  requireFeature('FEATURE_DEALS');

  // 加重金額の再計算
  if (data.amount !== undefined || data.probability !== undefined) {
    const existing = findRow(DEALS_SHEET, 'deal_id', id);
    if (existing) {
      const amount = Number(data.amount !== undefined ? data.amount : existing.amount);
      const prob   = Number(data.probability !== undefined ? data.probability : existing.probability);
      data.weighted_amount = Math.round(amount * prob / 100);
    }
  }

  data.updated_at = now();
  const result = updateRow(DEALS_SHEET, 'deal_id', id, data);
  if (!result) return errorResponse(`商談ID「${id}」が見つかりません。`);

  logInfo(`商談更新: ${id}`);
  return successResponse({ deal_id: id });
}

function updateDealPhase(id, phase) {
  requireFeature('FEATURE_DEALS');
  const result = updateRow(DEALS_SHEET, 'deal_id', id, { phase, updated_at: now() });
  if (!result) return errorResponse(`商談ID「${id}」が見つかりません。`);

  // 活動ログに自動追記
  const deal = findRow(DEALS_SHEET, 'deal_id', id);
  if (deal) {
    createActivity({
      customer_id: deal.customer_id,
      deal_id: id,
      type: 'OTHER',
      subject: `フェーズ変更: → ${phase}`,
      content: `商談「${deal.deal_name}」のフェーズを ${phase} に変更しました。`,
      activity_date: now(),
    });
  }
  logInfo(`フェーズ更新: ${id} → ${phase}`);
  return successResponse({ deal_id: id, phase });
}

function getAllDeals(userEmail) {
  requireFeature('FEATURE_DEALS');
  let data = getSheetData(DEALS_SHEET);
  if (!_isAdmin(userEmail)) {
    data = data.filter(r => r.assigned_user === userEmail);
  }
  return successResponse(data);
}

function getDealsByCustomer(customerId) {
  requireFeature('FEATURE_DEALS');
  const data = getSheetData(DEALS_SHEET);
  return successResponse(data.filter(r => r.customer_id === customerId));
}

function getPipelineSummary() {
  requireFeature('FEATURE_PIPELINE');
  const data = getSheetData(DEALS_SHEET);
  const phases = ['PROSPECT','APPROACH','PROPOSAL','NEGOTIATION','CLOSED_WON','CLOSED_LOST'];
  const summary = {};
  phases.forEach(p => {
    const items = data.filter(r => r.phase === p);
    summary[p] = {
      count: items.length,
      total_amount: items.reduce((s, r) => s + (Number(r.amount) || 0), 0),
      weighted_amount: items.reduce((s, r) => s + (Number(r.weighted_amount) || 0), 0),
    };
  });
  return successResponse(summary);
}

function getKanbanData(userEmail) {
  requireFeature('FEATURE_KANBAN');
  let data = getSheetData(DEALS_SHEET);
  if (!_isAdmin(userEmail)) {
    data = data.filter(r => r.assigned_user === userEmail);
  }
  const phases = ['PROSPECT','APPROACH','PROPOSAL','NEGOTIATION','CLOSED_WON','CLOSED_LOST'];
  const kanban = {};
  phases.forEach(p => { kanban[p] = data.filter(r => r.phase === p); });
  return successResponse(kanban);
}

function getDealsNearClose(days) {
  requireFeature('FEATURE_DEALS');
  const limit = daysFromNow(days);
  const today = new Date();
  const data = getSheetData(DEALS_SHEET);
  const near = data.filter(r => {
    if (!r.close_date || r.phase === 'CLOSED_WON' || r.phase === 'CLOSED_LOST') return false;
    const d = new Date(r.close_date);
    return d >= today && d <= limit;
  });
  return successResponse(near);
}
