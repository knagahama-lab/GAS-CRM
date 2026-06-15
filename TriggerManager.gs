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
