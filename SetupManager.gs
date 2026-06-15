/**
 * SetupManager.gs
 * スプレッドシートの初期セットアップ
 * 初回実行時に全シートを自動作成する
 */

function setupCRM() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  logInfo('CRM セットアップ開始');

  try {
    _createSettingsSheet(ss);
    _createUsersSheet(ss);
    _createMasterDataSheet(ss);
    _createCustomersSheet(ss);
    _createContactsSheet(ss);
    _createDealsSheet(ss);
    _createActivitiesSheet(ss);
    _createEmailTrackingSheet(ss);
    _createCalendarSheet(ss);
    _createTargetsSheet(ss);
    _createNotificationsSheet(ss);
    _createReportsSheet(ss);
    _createDashboardSheet(ss);

    // デフォルトシートを削除
    const defaultSheet = ss.getSheetByName('シート1');
    if (defaultSheet) ss.deleteSheet(defaultSheet);

    // Dashboardを最初に移動
    ss.setActiveSheet(ss.getSheetByName('🏠 Dashboard'));

    // トリガー設定
    setupTriggers();

    ui.alert('✅ CRM セットアップ完了！\n\n全シートとトリガーを設定しました。\n管理者メールアドレスをSettingsシートに入力してください。');
    logInfo('CRM セットアップ完了');
  } catch (e) {
    logError('セットアップエラー', e);
    ui.alert(`❌ エラー: ${e.message}`);
  }
}

// ── 各シート作成関数 ──────────────────────────────────────

function _createSettingsSheet(ss) {
  let sheet = ss.getSheetByName('Settings');
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet('Settings');

  const headers = ['設定キー', '設定値', '更新日時', '更新者', '説明'];
  const defaults = [
    ['ADMIN_EMAILS',         Session.getActiveUser().getEmail(), now(), '', '管理者メールアドレス（カンマ区切り）'],
    ['FEATURE_CUSTOMERS',   'TRUE',  now(), '', '顧客管理機能'],
    ['FEATURE_DEALS',       'TRUE',  now(), '', '商談管理機能'],
    ['FEATURE_ACTIVITIES',  'TRUE',  now(), '', '活動ログ機能'],
    ['FEATURE_CONTACTS',    'TRUE',  now(), '', '担当者管理機能'],
    ['FEATURE_EMAIL_TRACKING','TRUE',now(), '', 'Gmail自動追跡'],
    ['FEATURE_EMAIL_SEND',  'TRUE',  now(), '', 'CRMからメール送信'],
    ['FEATURE_CAL_CREATE',  'TRUE',  now(), '', 'カレンダー自動登録'],
    ['FEATURE_CAL_REMINDER','TRUE',  now(), '', 'リマインダー通知'],
    ['FEATURE_DASHBOARD',   'TRUE',  now(), '', 'KPIダッシュボード'],
    ['FEATURE_PIPELINE',    'TRUE',  now(), '', 'パイプライン分析'],
    ['FEATURE_REPORT_AUTO', 'TRUE',  now(), '', '定期レポート自動送信'],
    ['FEATURE_TARGETS',     'TRUE',  now(), '', '目標管理'],
    ['FEATURE_DEDUP',       'TRUE',  now(), '', '重複顧客チェック'],
    ['FEATURE_KANBAN',      'TRUE',  now(), '', 'カンバンボード'],
    ['FEATURE_SEARCH',      'TRUE',  now(), '', 'グローバル検索'],
    ['EMAIL_SCAN_INTERVAL', '15',    now(), '', 'Gmailスキャン間隔（分）'],
    ['REPORT_RECIPIENTS',   Session.getActiveUser().getEmail(), now(), '', 'レポート送信先メール（カンマ区切り）'],
    ['REPORT_SCHEDULE',     'WEEKLY',now(), '', 'レポート頻度 DAILY/WEEKLY/MONTHLY'],
    ['REMINDER_DAYS_BEFORE','3',     now(), '', 'クローズ予定日の何日前にリマインダー'],
    ['MAX_SHEET_ROWS',      '10000', now(), '', 'シート最大行数'],
    ['SYSTEM_VERSION',      '1.0.0', now(), '', 'システムバージョン'],
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#1A56DB').setFontColor('#FFFFFF').setFontWeight('bold');
  sheet.getRange(2, 1, defaults.length, defaults[0].length).setValues(defaults);
  _autoResizeColumns(sheet);
  logInfo('Settings シート作成完了');
}

function _createUsersSheet(ss) {
  let sheet = ss.getSheetByName('👤 Users');
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet('👤 Users');

  const headers = ['user_id','name','email','role','area','active','created_at'];
  const adminEmail = Session.getActiveUser().getEmail();
  const sample = [
    [generateId('USR'), '管理者', adminEmail, 'ADMIN', '全国', 'TRUE', now()],
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#1A56DB').setFontColor('#FFFFFF').setFontWeight('bold');
  sheet.getRange(2, 1, sample.length, sample[0].length).setValues(sample);
  _autoResizeColumns(sheet);
  logInfo('Users シート作成完了');
}

function _createMasterDataSheet(ss) {
  let sheet = ss.getSheetByName('📋 MasterData');
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet('📋 MasterData');

  const headers = ['category','code','label','sort_order','active'];
  const data = [
    // 業種
    ['INDUSTRY','IT','IT・システム',1,'TRUE'],
    ['INDUSTRY','MFCT','製造業',2,'TRUE'],
    ['INDUSTRY','RETAIL','小売・流通',3,'TRUE'],
    ['INDUSTRY','FIN','金融・保険',4,'TRUE'],
    ['INDUSTRY','MED','医療・介護',5,'TRUE'],
    ['INDUSTRY','CONST','建設・不動産',6,'TRUE'],
    ['INDUSTRY','SVC','サービス業',7,'TRUE'],
    ['INDUSTRY','OTHER','その他',8,'TRUE'],
    // 商談フェーズ
    ['DEAL_PHASE','PROSPECT','見込み',1,'TRUE'],
    ['DEAL_PHASE','APPROACH','アプローチ',2,'TRUE'],
    ['DEAL_PHASE','PROPOSAL','提案',3,'TRUE'],
    ['DEAL_PHASE','NEGOTIATION','交渉',4,'TRUE'],
    ['DEAL_PHASE','CLOSED_WON','受注',5,'TRUE'],
    ['DEAL_PHASE','CLOSED_LOST','失注',6,'TRUE'],
    // 活動種別
    ['ACTIVITY_TYPE','CALL','電話',1,'TRUE'],
    ['ACTIVITY_TYPE','EMAIL','メール',2,'TRUE'],
    ['ACTIVITY_TYPE','VISIT','訪問',3,'TRUE'],
    ['ACTIVITY_TYPE','MEETING','商談',4,'TRUE'],
    ['ACTIVITY_TYPE','DEMO','デモ',5,'TRUE'],
    ['ACTIVITY_TYPE','OTHER','その他',6,'TRUE'],
    // 顧客ステータス
    ['CUSTOMER_STATUS','PROSPECT','見込み客',1,'TRUE'],
    ['CUSTOMER_STATUS','ACTIVE','取引中',2,'TRUE'],
    ['CUSTOMER_STATUS','INACTIVE','休眠',3,'TRUE'],
    // カレンダー種別
    ['CAL_TYPE','MEETING','商談',1,'TRUE'],
    ['CAL_TYPE','FOLLOWUP','フォロー',2,'TRUE'],
    ['CAL_TYPE','DEMO','デモ',3,'TRUE'],
    ['CAL_TYPE','OTHER','その他',4,'TRUE'],
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#1A56DB').setFontColor('#FFFFFF').setFontWeight('bold');
  sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  _autoResizeColumns(sheet);
  logInfo('MasterData シート作成完了');
}

function _createCustomersSheet(ss) {
  let sheet = ss.getSheetByName('👥 Customers');
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet('👥 Customers');

  const headers = ['customer_id','company_name','industry','phone','email','address',
    'assigned_user','status','last_contact_date','notes','created_at','updated_at'];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#1A56DB').setFontColor('#FFFFFF').setFontWeight('bold');
  _autoResizeColumns(sheet);
  logInfo('Customers シート作成完了');
}

function _createContactsSheet(ss) {
  let sheet = ss.getSheetByName('📇 Contacts');
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet('📇 Contacts');

  const headers = ['contact_id','customer_id','name','title','email','phone','notes','created_at'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#1A56DB').setFontColor('#FFFFFF').setFontWeight('bold');
  _autoResizeColumns(sheet);
  logInfo('Contacts シート作成完了');
}

function _createDealsSheet(ss) {
  let sheet = ss.getSheetByName('💼 Deals');
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet('💼 Deals');

  const headers = ['deal_id','customer_id','company_name','deal_name','phase',
    'amount','probability','weighted_amount','close_date','assigned_user',
    'calendar_event_id','notes','created_at','updated_at'];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#1A56DB').setFontColor('#FFFFFF').setFontWeight('bold');
  _autoResizeColumns(sheet);
  logInfo('Deals シート作成完了');
}

function _createActivitiesSheet(ss) {
  let sheet = ss.getSheetByName('📝 Activities');
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet('📝 Activities');

  const headers = ['activity_id','customer_id','deal_id','type','activity_date',
    'subject','content','next_action','next_action_date','assigned_user','created_at'];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#1A56DB').setFontColor('#FFFFFF').setFontWeight('bold');
  _autoResizeColumns(sheet);
  logInfo('Activities シート作成完了');
}

function _createEmailTrackingSheet(ss) {
  let sheet = ss.getSheetByName('📧 EmailTracking');
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet('📧 EmailTracking');

  const headers = ['thread_id','customer_id','company_name','subject',
    'last_email_date','last_sender','direction','tracked_at'];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#1A56DB').setFontColor('#FFFFFF').setFontWeight('bold');
  _autoResizeColumns(sheet);
  logInfo('EmailTracking シート作成完了');
}

function _createCalendarSheet(ss) {
  let sheet = ss.getSheetByName('📅 Calendar');
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet('📅 Calendar');

  const headers = ['event_id','deal_id','customer_id','type','title',
    'start_datetime','end_datetime','description','assigned_user','created_at'];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#1A56DB').setFontColor('#FFFFFF').setFontWeight('bold');
  _autoResizeColumns(sheet);
  logInfo('Calendar シート作成完了');
}

function _createTargetsSheet(ss) {
  let sheet = ss.getSheetByName('🎯 Targets');
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet('🎯 Targets');

  const headers = ['target_id','assigned_user','period','period_label',
    'amount_target','deal_count_target','activity_count_target','created_at'];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#1A56DB').setFontColor('#FFFFFF').setFontWeight('bold');
  _autoResizeColumns(sheet);
  logInfo('Targets シート作成完了');
}

function _createNotificationsSheet(ss) {
  let sheet = ss.getSheetByName('🔔 Notifications');
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet('🔔 Notifications');

  const headers = ['notification_id','recipient','channel','subject','message',
    'status','created_at','sent_at'];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#1A56DB').setFontColor('#FFFFFF').setFontWeight('bold');
  _autoResizeColumns(sheet);
  logInfo('Notifications シート作成完了');
}

function _createReportsSheet(ss) {
  let sheet = ss.getSheetByName('📊 Reports');
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet('📊 Reports');

  const headers = ['report_id','type','period','generated_at','sent_to','status'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#1A56DB').setFontColor('#FFFFFF').setFontWeight('bold');
  _autoResizeColumns(sheet);
  logInfo('Reports シート作成完了');
}

function _createDashboardSheet(ss) {
  let sheet = ss.getSheetByName('🏠 Dashboard');
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet('🏠 Dashboard', 0);

  sheet.getRange('A1').setValue('📊 GAS CRM ダッシュボード')
    .setFontSize(18).setFontWeight('bold').setFontColor('#1A56DB');
  sheet.getRange('A2').setValue('最終更新：').setFontWeight('bold');
  sheet.getRange('B2').setValue(now());

  // KPIブロック
  const kpiLabels = [
    ['👥 顧客数（全体）','💼 商談数（進行中）','✅ 今月受注金額','📝 今月活動数'],
    ['=COUNTA(Customers!A:A)-1',
     '=COUNTIF(Deals!E:E,"<>CLOSED_WON")-COUNTIF(Deals!E:E,"CLOSED_LOST")-COUNTIF(Deals!E:E,"")',
     '=SUMPRODUCT((MONTH(Deals!I:I)=MONTH(TODAY()))*(YEAR(Deals!I:I)=YEAR(TODAY()))*(Deals!E:E="CLOSED_WON")*Deals!F:F)',
     '=COUNTIFS(Activities!E:E,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),Activities!E:E,"<="&TODAY())']
  ];

  sheet.getRange('A4:D4').setValues([kpiLabels[0]])
    .setBackground('#1A56DB').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('A5:D5').setValues([kpiLabels[1]])
    .setFontSize(20).setFontWeight('bold').setHorizontalAlignment('center').setBackground('#EFF6FF');

  sheet.getRange('A7').setValue('🔥 クローズ予定（直近30日）').setFontWeight('bold').setFontColor('#1A56DB');
  sheet.getRange('A8:F8').setValues([['商談名','顧客名','金額','確度','担当者','クローズ予定日']])
    .setBackground('#374151').setFontColor('#FFFFFF').setFontWeight('bold');

  sheet.getRange('A13').setValue('📈 フェーズ別パイプライン').setFontWeight('bold').setFontColor('#1A56DB');
  sheet.getRange('A14:C14').setValues([['フェーズ','件数','合計金額']])
    .setBackground('#374151').setFontColor('#FFFFFF').setFontWeight('bold');
  const phases = [['見込み','=COUNTIF(Deals!E:E,"PROSPECT")','=SUMIF(Deals!E:E,"PROSPECT",Deals!F:F)'],
    ['アプローチ','=COUNTIF(Deals!E:E,"APPROACH")','=SUMIF(Deals!E:E,"APPROACH",Deals!F:F)'],
    ['提案','=COUNTIF(Deals!E:E,"PROPOSAL")','=SUMIF(Deals!E:E,"PROPOSAL",Deals!F:F)'],
    ['交渉','=COUNTIF(Deals!E:E,"NEGOTIATION")','=SUMIF(Deals!E:E,"NEGOTIATION",Deals!F:F)'],
    ['受注','=COUNTIF(Deals!E:E,"CLOSED_WON")','=SUMIF(Deals!E:E,"CLOSED_WON",Deals!F:F)'],
    ['失注','=COUNTIF(Deals!E:E,"CLOSED_LOST")','=SUMIF(Deals!E:E,"CLOSED_LOST",Deals!F:F)']];
  sheet.getRange('A15:C20').setValues(phases);

  _autoResizeColumns(sheet);
  logInfo('Dashboard シート作成完了');
}

function _autoResizeColumns(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  sheet.autoResizeColumns(1, lastCol);
}
