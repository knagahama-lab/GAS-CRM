/**
 * MeetingModule.gs
 * 議事録AI - 録音（ブラウザ／Driveフォルダ）からの文字起こし・CRM自動紐付け
 *
 * フロー（JAPAN AI SALESの「記録と転記」「紐付け」を踏襲）:
 *   録音 → 文字起こし・項目抽出（AI） → 顧客/商談の紐付け候補を提示 → 営業が確認・承認 → CRMへ反映
 *   抽出できない項目は推測で埋めず、CRMと食い違う項目は人が選ぶまで更新しない。
 */

const MEETINGS_SHEET = '🎙️ Meetings';
const GLOSSARY_SHEET = '📚 Glossary';
const PRODUCT_DB_SHEET = '📦 ProductDB';

// ── 既存デプロイへの自動マイグレーション ─────────────────────
// 議事録AI機能のリリース前に setupCRM() 済みのスプレッドシートには
// Meetingsシートや新設定キーが存在しない。既存データを一切削除せず、
// 不足分だけを安全に追加する。画面初期表示のたびに呼ばれる想定（コストは軽微）。

function ensureMeetingsSetup() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let settingsChanged = false;

    // Meetingsシートが無ければ新規作成のみ行う（既存シートには一切触れない）
    if (!ss.getSheetByName(MEETINGS_SHEET)) {
      const sheet = ss.insertSheet(MEETINGS_SHEET);
      const headers = ['meeting_id','title','meeting_date','source','audio_file_id','audio_file_name',
        'duration_sec','status','customer_id','deal_id','company_name_guess','project_name_guess',
        'transcript','summary','extracted_json','next_action','next_action_date','amount_guess',
        'assigned_user','error_message','created_at','updated_at'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers])
        .setBackground('#1A56DB').setFontColor('#FFFFFF').setFontWeight('bold');
      sheet.autoResizeColumns(1, headers.length);
      logInfo('Meetings シートを自動補完作成しました');
    }

    // Meetingsシートに商談採点・次回提案AI用の列が無ければ末尾に追加する（既存列・データは一切変更しない）
    const meetingsSheet = ss.getSheetByName(MEETINGS_SHEET);
    if (meetingsSheet) {
      const meetingHeaders = meetingsSheet.getRange(1, 1, 1, meetingsSheet.getLastColumn()).getValues()[0];
      ['score_json', 'score_total', 'next_proposal_json', 'is_success_case'].forEach(col => {
        if (meetingHeaders.indexOf(col) === -1) {
          meetingsSheet.getRange(1, meetingsSheet.getLastColumn() + 1).setValue(col)
            .setBackground('#1A56DB').setFontColor('#FFFFFF').setFontWeight('bold');
          logInfo(`Meetings シートに列を自動補完しました: ${col}`);
        }
      });
    }

    // Glossary（専門用語辞書）シートが無ければ新規作成のみ行う
    if (!ss.getSheetByName(GLOSSARY_SHEET)) {
      const sheet = ss.insertSheet(GLOSSARY_SHEET);
      const headers = ['term', 'reading', 'definition', 'category'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers])
        .setBackground('#1A56DB').setFontColor('#FFFFFF').setFontWeight('bold');
      sheet.autoResizeColumns(1, headers.length);
      logInfo('Glossary シートを自動補完作成しました');
    }

    // ProductDB（商品DB）シートが無ければ新規作成のみ行う
    if (!ss.getSheetByName(PRODUCT_DB_SHEET)) {
      const sheet = ss.insertSheet(PRODUCT_DB_SHEET);
      const headers = ['product_id', 'product_name', 'category', 'description', 'features', 'target_use', 'created_at'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers])
        .setBackground('#1A56DB').setFontColor('#FFFFFF').setFontWeight('bold');
      sheet.autoResizeColumns(1, headers.length);
      logInfo('ProductDB シートを自動補完作成しました');
    }

    // Insightsシートが無ければ新規作成のみ行う（既存シートには一切触れない）
    if (!ss.getSheetByName(INSIGHTS_SHEET)) {
      const sheet = ss.insertSheet(INSIGHTS_SHEET);
      const headers = ['insight_id','customer_id','source_meeting_count','source_activity_count',
        'themes_json','problem_awareness','swot_json','pest_json','three_c_json','contact_points',
        'hypotheses_json','if_i_were_customer_json','proposal_actions_json','data_sufficiency',
        'assigned_user','created_at'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers])
        .setBackground('#1A56DB').setFontColor('#FFFFFF').setFontWeight('bold');
      sheet.autoResizeColumns(1, headers.length);
      logInfo('Insights シートを自動補完作成しました');
    }

    // Settingsシートに新設定キーが無ければ末尾に追記する（既存行は一切変更しない）
    const settingsSheet = ss.getSheetByName(SETTINGS_SHEET);
    if (settingsSheet) {
      const data = settingsSheet.getDataRange().getValues();
      const existingKeys = data.slice(1).map(r => r[0]);
      const defaults = [
        ['FEATURE_MEETINGS',      'TRUE', now(), '', '議事録AI（録音からの文字起こし・CRM自動反映）'],
        ['RECORDING_FOLDER_ID',   '',     now(), '', '議事録AI：録音データを監視するDriveフォルダID（空欄なら自動作成フォルダを使用）'],
        ['GEMINI_MODEL',          'gemini-2.5-flash', now(), '', '議事録AI：使用するGeminiモデル名'],
        ['MEETING_SCAN_INTERVAL', '10',   now(), '', '議事録AI：フォルダ監視の実行間隔（分）'],
        ['FEATURE_INSIGHTS',      'TRUE', now(), '', '顧客分析AI（SWOT/PEST/3C分析・仮説立案）'],
        ['FEATURE_MEETING_SCORING','TRUE', now(), '', '商談採点AI・次回提案サジェスト（成功事例比較・商品DB参照）'],
      ];
      const missing = defaults.filter(d => existingKeys.indexOf(d[0]) === -1);
      if (missing.length) {
        settingsSheet.getRange(settingsSheet.getLastRow() + 1, 1, missing.length, missing[0].length).setValues(missing);
        settingsChanged = true;
        logInfo(`Settings に議事録AI設定キーを自動補完しました: ${missing.map(m => m[0]).join(', ')}`);
      }
    }

    if (settingsChanged) _clearCache();
  } catch (e) {
    logError('議事録AIの自動セットアップ補完に失敗しました', e);
  }
}

// ── 録音アップロード（ブラウザ録音） ──────────────────────────

function uploadMeetingRecording(data) {
  return wrapAction(() => {
    requireFeature('FEATURE_MEETINGS');

    const valErr = validateRequired(data, ['base64Data', 'mimeType']);
    if (valErr) return errorResponse(valErr.error);

    const bytes = Utilities.base64Decode(data.base64Data);
    const sizeMb = bytes.length / (1024 * 1024);
    if (sizeMb > 19) {
      return errorResponse(`音声ファイルが大きすぎます（${sizeMb.toFixed(1)}MB）。19MB以下にしてください。`);
    }

    const filename = sanitizeString(data.filename) || `meeting_${Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss')}.webm`;
    const blob = Utilities.newBlob(bytes, data.mimeType, filename);

    const folder = _getRecordingFolder();
    const file = folder.createFile(blob);

    const meetingId = generateId('MTG');
    const row = {
      meeting_id: meetingId,
      title: sanitizeString(data.title) || `商談録音 ${formatDateTime(now())}`,
      meeting_date: data.meeting_date || now(),
      source: 'BROWSER',
      audio_file_id: file.getId(),
      audio_file_name: file.getName(),
      duration_sec: data.duration_sec || '',
      status: 'PROCESSING',
      customer_id: '',
      deal_id: '',
      company_name_guess: '',
      project_name_guess: '',
      transcript: '',
      summary: '',
      extracted_json: '',
      next_action: '',
      next_action_date: '',
      amount_guess: '',
      assigned_user: Session.getActiveUser().getEmail(),
      error_message: '',
      created_at: now(),
      updated_at: now(),
    };
    appendRow(MEETINGS_SHEET, row);
    logInfo(`議事録受付: ${meetingId}（ブラウザ録音）`);

    const result = _processMeeting(meetingId, blob);
    return successResponse(result);
  });
}

// ── Driveフォルダ自動取込 ─────────────────────────────────────

/**
 * 録音フォルダ内の未処理音声ファイルをスキャンし、順次AI処理する。
 * 管理画面からの手動実行、および定期トリガーの両方から呼ばれる。
 */
function scanRecordingFolder() {
  return wrapAction(() => {
    requireFeature('FEATURE_MEETINGS');

    const folder = _getRecordingFolder();
    const files = folder.getFiles();
    const existingIds = getSheetData(MEETINGS_SHEET).map(r => r.audio_file_id);
    const audioExt = /\.(mp3|wav|m4a|webm|ogg|flac|aac|mp4)$/i;

    let processed = 0;
    let failed = 0;
    const targets = [];

    while (files.hasNext()) {
      const file = files.next();
      if (existingIds.indexOf(file.getId()) !== -1) continue;
      const isAudio = audioExt.test(file.getName()) || String(file.getMimeType()).indexOf('audio/') === 0;
      if (!isAudio) continue;
      targets.push(file);
    }

    targets.forEach(file => {
      const meetingId = generateId('MTG');
      const row = {
        meeting_id: meetingId,
        title: file.getName(),
        meeting_date: file.getDateCreated(),
        source: 'FOLDER',
        audio_file_id: file.getId(),
        audio_file_name: file.getName(),
        duration_sec: '',
        status: 'PROCESSING',
        customer_id: '', deal_id: '',
        company_name_guess: '', project_name_guess: '',
        transcript: '', summary: '', extracted_json: '',
        next_action: '', next_action_date: '', amount_guess: '',
        assigned_user: Session.getActiveUser().getEmail(),
        error_message: '',
        created_at: now(), updated_at: now(),
      };
      appendRow(MEETINGS_SHEET, row);

      try {
        _processMeeting(meetingId, file.getBlob());
        processed++;
      } catch (e) {
        failed++; // _processMeeting内でERROR状態に記録済み
      }
    });

    logInfo(`フォルダ取込完了: 成功${processed}件 / 失敗${failed}件`);
    return successResponse({ processed, failed, total: targets.length });
  });
}

/**
 * 定期トリガーから呼ばれる薄いラッパー。未設定時は何もしない。
 */
function scanRecordingFolderTrigger() {
  if (!isEnabled('FEATURE_MEETINGS')) return;
  if (!getSetting('RECORDING_FOLDER_ID', '')) return;
  try {
    scanRecordingFolder();
  } catch (e) {
    logError('フォルダ自動取込に失敗', e);
  }
}

function _getRecordingFolder() {
  const folderId = getSetting('RECORDING_FOLDER_ID', '');
  if (folderId) {
    try {
      return DriveApp.getFolderById(folderId);
    } catch (e) {
      logError('RECORDING_FOLDER_IDが無効です。自動作成フォルダにフォールバックします', e);
    }
  }
  const existing = DriveApp.getFoldersByName('GAS-CRM 商談録音');
  if (existing.hasNext()) return existing.next();
  return DriveApp.createFolder('GAS-CRM 商談録音');
}

// ── AI処理本体 ────────────────────────────────────────────────

function _processMeeting(meetingId, blob) {
  try {
    const glossaryTerms = _getGlossaryTerms();
    const extracted = transcribeAndExtractAudio(blob, glossaryTerms);
    const matches = _findLinkCandidates(extracted);

    updateRow(MEETINGS_SHEET, 'meeting_id', meetingId, {
      status: 'REVIEW',
      transcript: extracted.transcript || '',
      summary: extracted.summary || '',
      company_name_guess: extracted.company_name || '',
      project_name_guess: extracted.project_name || '',
      next_action: extracted.next_action || '',
      next_action_date: extracted.next_action_date || '',
      amount_guess: extracted.amount_guess || '',
      extracted_json: JSON.stringify(extracted),
      updated_at: now(),
    });

    if (isEnabled('FEATURE_MEETING_SCORING')) {
      try {
        _scoreAndSuggestForMeeting(meetingId, extracted);
      } catch (e) {
        logError(`商談採点・次回提案の生成に失敗: ${meetingId}`, e);
      }
    }

    logInfo(`議事録AI処理完了: ${meetingId}（候補${matches.length}件）`);
    const meeting = findRow(MEETINGS_SHEET, 'meeting_id', meetingId);
    return { meeting, extracted, candidates: matches };
  } catch (e) {
    updateRow(MEETINGS_SHEET, 'meeting_id', meetingId, {
      status: 'ERROR',
      error_message: e.message,
      updated_at: now(),
    });
    logError(`議事録AI処理失敗: ${meetingId}`, e);
    throw e;
  }
}

// ── 商談採点・次回提案サジェスト ──────────────────────────────
// 「商談内容の採点（過去の成功事例と比較・育成に活用）」
// 「次回の商談内容を提案（商品DBを参照・提案書に反映）」のステップ

function _scoreAndSuggestForMeeting(meetingId, extracted) {
  const successExamples = _getSuccessCaseSummaries(5);
  const scoreResult = scoreMeetingAI({
    transcript: extracted.transcript,
    summary: extracted.summary,
    companyName: extracted.company_name,
    successExamples,
  });

  const products = _getProductDbList();
  const suggestResult = suggestNextMeetingAI({
    summary: extracted.summary,
    companyName: extracted.company_name,
    nextAction: extracted.next_action,
    products,
  });

  updateRow(MEETINGS_SHEET, 'meeting_id', meetingId, {
    score_json: JSON.stringify(scoreResult),
    score_total: scoreResult.total_score !== undefined ? scoreResult.total_score : '',
    next_proposal_json: JSON.stringify(suggestResult),
  });

  logInfo(`商談採点・次回提案サジェスト完了: ${meetingId}（総合点${scoreResult.total_score}）`);
}

/**
 * 処理済みだが採点未実施の議事録を後から採点する（機能追加前のデータへの対応・再採点用）
 */
function rescoreMeeting(meetingId) {
  return wrapAction(() => {
    requireFeature('FEATURE_MEETINGS');
    const meeting = findRow(MEETINGS_SHEET, 'meeting_id', meetingId);
    if (!meeting) return errorResponse(`議事録ID「${meetingId}」が見つかりません。`);
    if (!meeting.extracted_json) return errorResponse('文字起こしデータがまだありません。処理完了後に再度お試しください。');

    let extracted;
    try { extracted = JSON.parse(meeting.extracted_json); } catch (e) { return errorResponse('抽出データの解析に失敗しました。'); }

    _scoreAndSuggestForMeeting(meetingId, extracted);
    const updated = findRow(MEETINGS_SHEET, 'meeting_id', meetingId);
    return successResponse(updated);
  });
}

function _getGlossaryTerms() {
  try {
    return getSheetData(GLOSSARY_SHEET)
      .filter(r => r.term)
      .map(r => `${r.term}${r.reading ? '（' + r.reading + '）' : ''}: ${r.definition || ''}`)
      .slice(0, 100);
  } catch (e) {
    return [];
  }
}

function _getProductDbList() {
  try {
    return getSheetData(PRODUCT_DB_SHEET)
      .filter(r => r.product_name)
      .map(r => ({ name: r.product_name, category: r.category, description: r.description, target_use: r.target_use }));
  } catch (e) {
    return [];
  }
}

function _getSuccessCaseSummaries(limit) {
  try {
    const data = getSheetData(MEETINGS_SHEET)
      .filter(r => String(r.is_success_case).toUpperCase() === 'TRUE' && r.summary);
    data.sort((a, b) => new Date(b.created_at) - new Date(a.created_at));
    return data.slice(0, limit).map(r => ({ company: r.company_name_guess, summary: r.summary }));
  } catch (e) {
    return [];
  }
}

/**
 * 商談が受注（CLOSED_WON）になった際、紐付く議事録を「成功事例」としてフラグする。
 * DealModule.gs の updateDealPhase から呼ばれる。
 */
function _markMeetingsAsSuccessCase(dealId) {
  try {
    const sheet = getSheet(MEETINGS_SHEET);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const dealIdx = headers.indexOf('deal_id');
    const successIdx = headers.indexOf('is_success_case');
    if (dealIdx === -1 || successIdx === -1) return;
    for (let i = 1; i < data.length; i++) {
      if (data[i][dealIdx] === dealId) {
        sheet.getRange(i + 1, successIdx + 1).setValue('TRUE');
      }
    }
  } catch (e) {
    logError('成功事例フラグの更新に失敗しました', e);
  }
}

// ── 顧客・商談への紐付け候補推定 ──────────────────────────────
// 「録音前に商談を選ばなくても、AIが内容から紐付け先を推定し、候補ごとに根拠を表示する」

function _findLinkCandidates(extracted) {
  const companyGuess = String(extracted.company_name || '').trim();
  if (!companyGuess) return [];

  const normalizedGuess = _normalizeCompanyName(companyGuess);
  if (!normalizedGuess) return [];

  const customers = getSheetData(CUSTOMERS_SHEET);
  const allDeals = getSheetData(DEALS_SHEET);
  const projectGuessNorm = _normalizeCompanyName(extracted.project_name || '');

  const candidates = [];
  customers.forEach(c => {
    const normalizedName = _normalizeCompanyName(c.company_name);
    if (!normalizedName) return;

    let score = 0;
    let reason = '';
    if (normalizedName === normalizedGuess) {
      score = 95; reason = `会社名が完全一致（「${companyGuess}」）`;
    } else if (normalizedName.indexOf(normalizedGuess) !== -1 || normalizedGuess.indexOf(normalizedName) !== -1) {
      score = 70; reason = `会社名が部分一致（「${companyGuess}」）`;
    }
    if (score === 0) return;

    const dealsForCustomer = allDeals.filter(d => d.customer_id === c.customer_id);
    let bestDeal = null;
    if (projectGuessNorm) {
      bestDeal = dealsForCustomer.find(d => {
        const dn = _normalizeCompanyName(d.deal_name);
        return dn && (dn.indexOf(projectGuessNorm) !== -1 || projectGuessNorm.indexOf(dn) !== -1);
      }) || null;
      if (bestDeal) { score += 3; reason += ` ／ 案件名「${extracted.project_name}」も一致`; }
    }
    if (!bestDeal) {
      // 進行中（未クローズ）の商談があれば参考候補として提示
      bestDeal = dealsForCustomer.find(d => d.phase !== 'CLOSED_WON' && d.phase !== 'CLOSED_LOST') || null;
    }

    candidates.push({
      customer_id: c.customer_id,
      company_name: c.company_name,
      deal_id: bestDeal ? bestDeal.deal_id : '',
      deal_name: bestDeal ? bestDeal.deal_name : '',
      score,
      reason,
    });
  });

  candidates.sort((a, b) => b.score - a.score);
  return candidates.slice(0, 5);
}

function _normalizeCompanyName(name) {
  return String(name || '')
    .toLowerCase()
    .replace(/株式会社|有限会社|合同会社|\(株\)|（株）|\(有\)|（有）/g, '')
    .replace(/[\s　]+/g, '')
    .trim();
}

// ── 一覧・詳細取得 ────────────────────────────────────────────

function getAllMeetings(userEmail) {
  return wrapAction(() => {
    requireFeature('FEATURE_MEETINGS');
    let data = getSheetData(MEETINGS_SHEET);
    if (!_isAdmin(userEmail)) {
      data = data.filter(r => r.assigned_user === userEmail);
    }
    data.sort((a, b) => new Date(b.created_at) - new Date(a.created_at));
    return successResponse(data);
  });
}

function getMeeting(meetingId) {
  return wrapAction(() => {
    requireFeature('FEATURE_MEETINGS');
    const meeting = findRow(MEETINGS_SHEET, 'meeting_id', meetingId);
    if (!meeting) return errorResponse(`議事録ID「${meetingId}」が見つかりません。`);
    let candidates = [];
    if (meeting.extracted_json) {
      try { candidates = _findLinkCandidates(JSON.parse(meeting.extracted_json)); } catch (e) { /* noop */ }
    }
    return successResponse({ meeting, candidates });
  });
}

/**
 * 処理失敗時の再実行
 */
function retryMeetingProcessing(meetingId) {
  return wrapAction(() => {
    requireFeature('FEATURE_MEETINGS');
    const meeting = findRow(MEETINGS_SHEET, 'meeting_id', meetingId);
    if (!meeting) return errorResponse(`議事録ID「${meetingId}」が見つかりません。`);

    const file = DriveApp.getFileById(meeting.audio_file_id);
    updateRow(MEETINGS_SHEET, 'meeting_id', meetingId, { status: 'PROCESSING', error_message: '', updated_at: now() });
    const result = _processMeeting(meetingId, file.getBlob());
    return successResponse(result);
  });
}

function deleteMeeting(meetingId) {
  return wrapAction(() => {
    requireFeature('FEATURE_MEETINGS');
    const sheet = getSheet(MEETINGS_SHEET);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIdx = headers.indexOf('meeting_id');
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] === meetingId) {
        sheet.deleteRow(i + 1);
        logInfo(`議事録削除: ${meetingId}`);
        return successResponse({ deleted: true });
      }
    }
    return errorResponse(`議事録ID「${meetingId}」が見つかりません。`);
  });
}

// ── 確認・承認 → CRM反映 ──────────────────────────────────────
// 「抽出できない項目は推測で埋めない／CRMと食い違う項目は人が選ぶまで更新しない／
//   面談1回につき活動履歴を必ず1件作成」という方針をそのまま実装する。

function confirmMeetingLink(data) {
  return wrapAction(() => {
    requireFeature('FEATURE_MEETINGS');

    const valErr = validateRequired(data, ['meeting_id']);
    if (valErr) return errorResponse(valErr.error);

    const meeting = findRow(MEETINGS_SHEET, 'meeting_id', data.meeting_id);
    if (!meeting) return errorResponse(`議事録ID「${data.meeting_id}」が見つかりません。`);

    let customerId = data.customer_id || '';

    // 新規顧客として登録する場合
    if (customerId === 'NEW') {
      const custRes = JSON.parse(createCustomer({
        company_name: data.new_customer_name || meeting.company_name_guess,
        email: data.new_customer_email || '',
        industry: data.new_customer_industry || '',
        phone: data.new_customer_phone || '',
      }));
      if (!custRes.success) return errorResponse('顧客の新規作成に失敗しました: ' + custRes.error);
      customerId = custRes.data.customer_id;
    }

    let dealId = data.deal_id || '';

    // 新規商談として登録する場合
    if (dealId === 'NEW') {
      if (!customerId) return errorResponse('商談を新規作成するには顧客の紐付けが必要です。');
      const dealRes = JSON.parse(createDeal({
        customer_id: customerId,
        deal_name: data.new_deal_name || meeting.project_name_guess || meeting.title,
        phase: data.deal_phase || 'PROSPECT',
        amount: data.amount || meeting.amount_guess || 0,
        probability: data.probability || 20,
      }));
      if (!dealRes.success) return errorResponse('商談の新規作成に失敗しました: ' + dealRes.error);
      dealId = dealRes.data.deal_id;
    }

    // 面談1回につき活動履歴を必ず1件作成
    const actRes = JSON.parse(createActivity({
      customer_id: customerId,
      deal_id: dealId,
      type: data.activity_type || 'MEETING',
      activity_date: meeting.meeting_date,
      subject: data.subject || meeting.title,
      content: data.summary !== undefined ? data.summary : meeting.summary,
      next_action: data.next_action !== undefined ? data.next_action : meeting.next_action,
      next_action_date: data.next_action_date !== undefined ? data.next_action_date : meeting.next_action_date,
    }));

    updateRow(MEETINGS_SHEET, 'meeting_id', data.meeting_id, {
      status: 'LINKED',
      customer_id: customerId,
      deal_id: dealId,
      updated_at: now(),
    });

    logInfo(`議事録紐付け確定: ${data.meeting_id} → 顧客${customerId} / 商談${dealId || '(なし)'}`);
    return successResponse({
      meeting_id: data.meeting_id,
      customer_id: customerId,
      deal_id: dealId,
      activity_id: actRes.success ? actRes.data.activity_id : '',
    });
  });
}

/**
 * 紐付けを保留のままにする（未紐付けステータスに戻す等は不要。REVIEWのまま維持）
 */
function skipMeetingLink(meetingId) {
  return wrapAction(() => {
    requireFeature('FEATURE_MEETINGS');
    const result = updateRow(MEETINGS_SHEET, 'meeting_id', meetingId, { status: 'UNLINKED', updated_at: now() });
    if (!result) return errorResponse(`議事録ID「${meetingId}」が見つかりません。`);
    return successResponse({ meeting_id: meetingId });
  });
}
