/**
 * MeetingModule.gs
 * 議事録AI - 録音（ブラウザ／Driveフォルダ）からの文字起こし・CRM自動紐付け
 *
 * フロー（JAPAN AI SALESの「記録と転記」「紐付け」を踏襲）:
 *   録音 → 文字起こし・項目抽出（AI） → 顧客/商談の紐付け候補を提示 → 営業が確認・承認 → CRMへ反映
 *   抽出できない項目は推測で埋めず、CRMと食い違う項目は人が選ぶまで更新しない。
 */

const MEETINGS_SHEET = '🎙️ Meetings';

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
    const extracted = transcribeAndExtractAudio(blob);
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
