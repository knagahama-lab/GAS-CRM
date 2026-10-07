/**
 * TriggerManager.gs
 * 定期トリガーの登録・管理
 */

function setupTriggers() {
  // 既存トリガーを全削除してから再登録
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  const interval = Number(getSetting('EMAIL_SCAN_INTERVAL', 15));

  // T1: Gmailスキャン（15分ごと）
  ScriptApp.newTrigger('scanEmails')
    .timeBased().everyMinutes(interval).create();

  // T2: 毎朝リマインダー（8:00）
  ScriptApp.newTrigger('sendDailyReminders')
    .timeBased().everyDays(1).atHour(8).create();

  // T3: 週次レポート（月曜9:00）
  ScriptApp.newTrigger('generateWeeklyReport')
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(9).create();

  // T4: ダッシュボード更新（1時間ごと）
  ScriptApp.newTrigger('updateDashboard')
    .timeBased().everyHours(1).create();

  // T5: 月次レポート（毎月1日9:00）
  ScriptApp.newTrigger('monthlyReportTrigger')
    .timeBased().onMonthDay(1).atHour(9).create();

  // T6: 議事録AI - 録音フォルダ監視（デフォルト10分ごと。未設定時は関数内で何もしない）
  const meetingInterval = Number(getSetting('MEETING_SCAN_INTERVAL', 10));
  if (meetingInterval > 0) {
    ScriptApp.newTrigger('scanRecordingFolderTrigger')
      .timeBased().everyMinutes(meetingInterval).create();
  }

  // T7: 顧客分析AI - 新しい議事録・活動がある顧客を自動で再分析（毎朝6:00）
  ScriptApp.newTrigger('refreshStaleCustomerInsightsTrigger')
    .timeBased().everyDays(1).atHour(6).create();

  // T8: PDCAレビュー - 全担当者分を月次で自動生成・メール送信（毎月1日8:00）
  ScriptApp.newTrigger('generateMonthlyPdcaReviewsTrigger')
    .timeBased().onMonthDay(1).atHour(8).create();

  logInfo('トリガー設定完了');
}

function monthlyReportTrigger() {
  generateMonthlyReport();
}

function removeTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  logInfo('全トリガーを削除しました');
}

function getTriggerStatus() {
  const triggers = ScriptApp.getProjectTriggers();
  return successResponse(triggers.map(t => ({
    id: t.getUniqueId(),
    function: t.getHandlerFunction(),
    type: t.getEventType().toString(),
  })));
}
