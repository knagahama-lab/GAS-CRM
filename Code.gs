/**
 * Code.gs
 * Web Appエントリーポイント・ルーティング
 * GAS Web App として公開するメインファイル
 */

/**
 * GET リクエスト処理 - メインUI表示
 */
function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || 'main';

  try {
    if (page === 'admin') {
      return HtmlService.createTemplateFromFile('AdminPanel')
        .evaluate()
        .setTitle('GAS CRM 管理画面')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }
    // kanban は Index.html 内で navigate('kanban') により表示するため、
    // 独立した KanbanView ファイルは存在しない。Index にフォールバックする。
    return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('GAS CRM 営業管理システム')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (err) {
    return HtmlService.createHtmlOutput(`<h2>エラー: ${err.message}</h2>`);
  }
}

/**
 * POST リクエスト処理 - API エンドポイント
 */
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action  = payload.action;
    const data    = payload.data || {};

    const result = _dispatchAction(action, data);
    return ContentService
      .createTextOutput(result)
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(errorResponse(err.message))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * クライアントサイドから直接呼び出せるサーバー関数群
 */
function serverAction(action, data) {
  return _dispatchAction(action, data || {});
}

function _dispatchAction(action, data) {
  const userEmail = Session.getActiveUser().getEmail();

  switch (action) {
    // ── 初期データ ─────────────────────────────
    case 'getInitData':
      return _getInitData(userEmail);

    // ── 顧客管理 ──────────────────────────────
    case 'getAllCustomers':    return getAllCustomers(userEmail);
    case 'getCustomer':       return getCustomer(data.customer_id);
    case 'createCustomer':    return createCustomer(data);
    case 'updateCustomer':    return updateCustomer(data.customer_id, data);
    case 'searchCustomers':   return searchCustomers(data.query);
    case 'getCustomerStats':  return getCustomerStats();

    // ── 商談管理 ──────────────────────────────
    case 'getAllDeals':        return getAllDeals(userEmail);
    case 'createDeal':        return createDeal(data);
    case 'updateDeal':        return updateDeal(data.deal_id, data);
    case 'updateDealPhase':   return updateDealPhase(data.deal_id, data.phase);
    case 'getDealsByCustomer':return getDealsByCustomer(data.customer_id);
    case 'getPipelineSummary':return getPipelineSummary();
    case 'getKanbanData':     return getKanbanData(userEmail);
    case 'getDealsNearClose': return getDealsNearClose(data.days || 7);

    // ── 活動ログ ──────────────────────────────
    case 'createActivity':          return createActivity(data);
    case 'getActivitiesByCustomer': return getActivitiesByCustomer(data.customer_id);
    case 'getActivitiesByDeal':     return getActivitiesByDeal(data.deal_id);
    case 'getRecentActivities':     return getRecentActivities(data.limit, userEmail);
    case 'getPendingNextActions':   return getPendingNextActions(userEmail);
    case 'getActivityStats':        return getActivityStats(userEmail);

    // ── メール ────────────────────────────────
    case 'sendEmail':         return sendEmail(data);
    case 'getEmailHistory':   return getEmailHistory(data.customer_id);

    // ── レポート ──────────────────────────────
    case 'getKpiSummary':     return getKpiSummary();

    // ── ユーザー・マスタ ──────────────────────
    case 'getCurrentUser':    return successResponse(getCurrentUser());
    case 'getAllUsers':        return getAllUsers();
    case 'createUser':        return createUser(data);
    case 'updateUser':        return updateUser(data.user_id, data);
    case 'getMasterData':     return getMasterData(data.category);

    // ── 設定管理 ──────────────────────────────
    case 'getAllSettings':    return successResponse(getAllSettings());
    case 'setFeature':        return (setFeature(data.key, data.value), successResponse({ key: data.key }));

    // ── システム ──────────────────────────────
    case 'getTriggerStatus':  return getTriggerStatus();
    case 'setupTriggers':     return (setupTriggers(), successResponse({ done: true }));

    default:
      return errorResponse(`不明なアクション: ${action}`);
  }
}

/**
 * 画面初期表示に必要なデータを一括取得
 */
function _getInitData(userEmail) {
  const user = getCurrentUser();
  const settings = getAllSettings();
  const kpi = JSON.parse(getKpiSummary());
  const industries = JSON.parse(getMasterData('INDUSTRY'));
  const phases = JSON.parse(getMasterData('DEAL_PHASE'));
  const actTypes = JSON.parse(getMasterData('ACTIVITY_TYPE'));
  const users = JSON.parse(getAllUsers());

  return successResponse({
    user,
    settings,
    kpi: kpi.data,
    masters: {
      industries: industries.data,
      phases: phases.data,
      activityTypes: actTypes.data,
    },
    users: users.data,
    isAdmin: _isAdmin(userEmail),
    isManager: _isManager(userEmail),
  });
}

/**
 * デプロイ済みWeb Appの実URLを返す。
 * サンドボックスiframeからトップフレームを正しいURLへ遷移させるために使用する。
 */
function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

/**
 * HTMLファイルをincludeするヘルパー
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
