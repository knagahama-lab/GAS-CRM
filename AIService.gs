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

/**
 * 顧客分析AI - CX研修の手法（議事録・活動記録の抽象化→SWOT/PEST/3C分析→
 * 顧客の問題意識と自社接点の発見→仮説立案→「自分が顧客だったら」視点での
 * 提案アクション抽出）に沿って、蓄積された議事録・活動履歴から顧客分析を生成する。
 */
function generateCustomerInsightAI(context) {
  const apiKey = _getGeminiApiKey();
  if (!apiKey) {
    throw new Error('Gemini APIキーが未設定です。管理画面「議事録AI」タブから設定してください。');
  }

  const model = getSetting('GEMINI_MODEL', 'gemini-2.5-flash');
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

  const meetingsText = (context.meetings || []).map((m, i) =>
    `[議事録${i + 1}] ${m.date || ''} 「${m.title || ''}」\n要約: ${m.summary || '(なし)'}\n文字起こし抜粋: ${(m.transcript || '(なし)').slice(0, 1500)}`
  ).join('\n\n');

  const activitiesText = (context.activities || []).map((a, i) =>
    `[活動${i + 1}] ${a.date || ''} ${a.type || ''} 「${a.subject || ''}」: ${a.content || ''}${a.next_action ? ' / 次回アクション: ' + a.next_action : ''}`
  ).join('\n');

  const dealsText = (context.deals || []).map(d => `${d.name}（${d.phase}, ${d.amount}円）`).join('、');

  const prompt = [
    'あなたは法人営業のCX（カスタマーエクスペリエンス）研修で学んだ手法を用いて、',
    '営業担当者が記録した商談議事録・活動履歴から顧客分析を行う専門アシスタントです。',
    '',
    '分析の考え方（研修内容。この順序・観点に沿って分析すること）:',
    '1. 商談で話していた内容を整理・分類し、個別の発言を「これは要するに〇〇のことだ」と抽象化してテーマごとにまとめる',
    '2. 抽象化した内容から、顧客自身もまだ明確に言語化できていないかもしれない課題・問題意識を洗い出す',
    '3. SWOT分析・PEST分析・3C分析を使い、顧客を取り巻く環境を整理する',
    '   （3C = 「顧客（顧客の市場・顧客の顧客）」「競合（顧客のライバル）」「自社（ここでは顧客自身の会社の強み弱み）」を指す）',
    '   （PEST = 顧客を取り巻くマクロ環境＝政治的・経済的・社会文化的・技術的要因）',
    '   （SWOT = 顧客自身の強み・弱み・機会・脅威）',
    '4. 顧客の問題意識と、弊社（営業担当の所属企業）が提供できることとの接点を探す（顧客の発言だけで判断せず、また自社が提案できるかどうかだけでも判断しない）',
    '5. 顧客の課題に対する仮説を立てる（顧客は自らの課題を正確に把握していないことが多い、という前提に立つ）',
    '6. 「自分が顧客だったら何をするか」を具体的に考える（必ずしも弊社で提供できることとは限らない）',
    '7. 顧客に寄り添っていると感じてもらい、提案力・クロージング力を強化するための次の提案アクションを挙げる',
    '',
    '厳守事項: 入力データに無い事実を創作しないこと。業界の一般的な傾向など常識的な推測は許容するが、推測である旨が分かる書き方にすること。',
    '入力情報が乏しい項目は無理に埋めず、該当項目に「情報不足」と明記すること。',
    '',
    '【顧客情報】',
    `会社名: ${context.companyName || '不明'}`,
    `業種: ${context.industry || '不明'}`,
    `メモ: ${context.notes || 'なし'}`,
    `商談中の案件: ${dealsText || 'なし'}`,
    '',
    '【商談議事録】',
    meetingsText || '(議事録なし)',
    '',
    '【活動履歴】',
    activitiesText || '(活動履歴なし)',
    '',
    '出力は以下のキーを持つJSONオブジェクト1つのみとし、説明文やコードブロック記法は付けないこと:',
    '{',
    '  "themes": [ { "theme": "テーマ名", "summary": "要約", "evidence": "該当する発言・記録の抜粋や要約" } ],',
    '  "problem_awareness": "抽象化・整理した後の、顧客の課題・問題意識",',
    '  "swot": { "strengths": [""], "weaknesses": [""], "opportunities": [""], "threats": [""] },',
    '  "pest": { "political": "", "economic": "", "social": "", "technological": "" },',
    '  "three_c": { "customer": "顧客の市場・顧客の顧客についての整理", "competitor": "顧客のライバルについての整理", "company": "顧客自身（強み弱み）についての整理" },',
    '  "contact_points": "顧客の問題意識と自社（営業担当の会社）が提供できることの接点",',
    '  "hypotheses": [ { "hypothesis": "仮説", "rationale": "根拠" } ],',
    '  "if_i_were_customer": ["自分が顧客だったらやること"],',
    '  "proposal_actions": ["次に取るべき提案アクション"],',
    '  "data_sufficiency": "入力情報が十分かどうかのコメント（不足があれば具体的に何が足りないか）"',
    '}',
  ].join('\n');

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { responseMimeType: 'application/json', temperature: 0.4 },
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
    logError('Gemini APIエラー（顧客分析）', { message: `HTTP ${code}: ${body.substring(0, 500)}` });
    throw new Error(`AI分析でエラーが発生しました（HTTP ${code}）。APIキーや利用枠をご確認ください。`);
  }

  let json;
  try { json = JSON.parse(body); } catch (e) {
    throw new Error('Gemini応答の解析に失敗しました。');
  }

  const candidate = json.candidates && json.candidates[0];
  const text = candidate && candidate.content && candidate.content.parts && candidate.content.parts[0] && candidate.content.parts[0].text;
  if (!text) {
    const finishReason = candidate && candidate.finishReason;
    throw new Error(`AIから有効な応答が得られませんでした（理由: ${finishReason || '不明'}）。`);
  }

  let extracted;
  try {
    extracted = JSON.parse(text);
  } catch (e) {
    throw new Error('AI出力のJSON解析に失敗しました: ' + e.message);
  }

  return extracted;
}
