/**
 * AIService.gs
 * Gemini API 連携 - 商談音声の文字起こし・話者分離・CRM項目抽出
 */

const GEMINI_API_KEY_PROP = 'GEMINI_API_KEY';
const GEMINI_INLINE_MAX_BYTES = 19 * 1024 * 1024; // Gemini inlineData の実用上限（約19MB）

/**
 * Gemini APIキーを保存（管理者のみ）。
 * キー自体はスプレッドシートに書かず、スクリプトのプロパティストアに保存する。
 */
function setGeminiApiKey(apiKey) {
  return wrapAction(() => {
    _requireAdmin();
    const key = String(apiKey || '').trim();
    if (!key) return errorResponse('APIキーが空です。');
    PropertiesService.getScriptProperties().setProperty(GEMINI_API_KEY_PROP, key);
    logInfo('Gemini APIキーを更新しました');
    return successResponse({ saved: true });
  });
}

/**
 * APIキーが設定済みかどうかのみを返す（値自体は返さない）
 */
function hasGeminiApiKey() {
  return successResponse({ hasKey: !!_getGeminiApiKey() });
}

function _getGeminiApiKey() {
  return PropertiesService.getScriptProperties().getProperty(GEMINI_API_KEY_PROP) || '';
}

/**
 * 音声Blobを渡すと、文字起こし・要約・CRM項目候補をJSONで返す
 */
function transcribeAndExtractAudio(audioBlob) {
  const apiKey = _getGeminiApiKey();
  if (!apiKey) {
    throw new Error('Gemini APIキーが未設定です。管理画面「議事録AI」タブから設定してください。');
  }

  const bytes = audioBlob.getBytes();
  if (bytes.length > GEMINI_INLINE_MAX_BYTES) {
    throw new Error(`音声ファイルが大きすぎます（${(bytes.length / 1024 / 1024).toFixed(1)}MB）。19MB以下（目安：30〜40分程度）に分割してください。`);
  }

  const model = getSetting('GEMINI_MODEL', 'gemini-2.5-flash');
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

  const base64Audio = Utilities.base64Encode(bytes);
  const mimeType = audioBlob.getContentType() || 'audio/webm';

  const prompt = [
    'あなたは日本企業の営業商談・打ち合わせの音声を解析する専門アシスタントです。',
    '添付された音声を文字起こしし、話者を「話者A」「話者B」のように分離してください。',
    '聞き取れない箇所は[不明瞭]と記載し、絶対に内容を創作しないでください。',
    '',
    '出力は以下のキーを持つJSONオブジェクト1つのみとし、説明文やコードブロック記法は付けないでください。',
    '抽出できない・自信がない項目は空文字（amount_guessは空文字または数値）にしてください。推測で埋めないこと。',
    '',
    '{',
    '  "transcript": "話者ラベル付きの全文文字起こし",',
    '  "summary": "商談内容の要約（250〜350字程度、日本語）",',
    '  "company_name": "先方の会社名（文字起こしから明確に読み取れる場合のみ）",',
    '  "contact_person": "先方の担当者名（読み取れる場合のみ）",',
    '  "project_name": "案件名・商材名（読み取れる場合のみ）",',
    '  "deal_phase_guess": "PROSPECT, APPROACH, PROPOSAL, NEGOTIATION, CLOSED_WON, CLOSED_LOST のいずれか、または空文字",',
    '  "amount_guess": "会話中に言及された金額（半角数字のみ、不明なら空文字）",',
    '  "next_action": "合意された次回アクション（例: 見積送付、再訪問）",',
    '  "next_action_date": "次回アクション予定日（YYYY-MM-DD形式、不明なら空文字）",',
    '  "activity_type_guess": "CALL, EMAIL, VISIT, MEETING, DEMO, OTHER のいずれか"',
    '}',
  ].join('\n');

  const payload = {
    contents: [{
      parts: [
        { text: prompt },
        { inline_data: { mime_type: mimeType, data: base64Audio } },
      ],
    }],
    generationConfig: {
      responseMimeType: 'application/json',
      temperature: 0.2,
    },
  };

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  const code = res.getResponseCode();
  const body = res.getContentText();

  if (code !== 200) {
    logError('Gemini APIエラー', { message: `HTTP ${code}: ${body.substring(0, 500)}` });
    throw new Error(`AI解析でエラーが発生しました（HTTP ${code}）。APIキーや利用枠をご確認ください。`);
  }

  let json;
  try { json = JSON.parse(body); } catch (e) {
    throw new Error('Gemini応答の解析に失敗しました。');
  }

  const candidate = json.candidates && json.candidates[0];
  const text = candidate && candidate.content && candidate.content.parts && candidate.content.parts[0] && candidate.content.parts[0].text;
  if (!text) {
    const finishReason = candidate && candidate.finishReason;
    throw new Error(`AIから有効な応答が得られませんでした（理由: ${finishReason || '不明'}）。音声の長さや形式をご確認ください。`);
  }

  let extracted;
  try {
    extracted = JSON.parse(text);
  } catch (e) {
    throw new Error('AI出力のJSON解析に失敗しました: ' + e.message);
  }

  return {
    transcript: extracted.transcript || '',
    summary: extracted.summary || '',
    company_name: extracted.company_name || '',
    contact_person: extracted.contact_person || '',
    project_name: extracted.project_name || '',
    deal_phase_guess: extracted.deal_phase_guess || '',
    amount_guess: extracted.amount_guess || '',
    next_action: extracted.next_action || '',
    next_action_date: extracted.next_action_date || '',
    activity_type_guess: extracted.activity_type_guess || 'MEETING',
  };
}
