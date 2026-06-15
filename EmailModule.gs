/**
 * EmailModule.gs
 * Gmail追跡・送信・自動スキャン
 */

const EMAIL_TRACKING_SHEET = '📧 EmailTracking';

/**
 * Gmail自動スキャン（トリガーで定期実行）
 */
function scanEmails() {
  if (!isEnabled('FEATURE_EMAIL_TRACKING')) return;

  logInfo('Gmailスキャン開始');
  const customers = getSheetData('👥 Customers');
  if (!customers.length) return;

  let tracked = 0;
  customers.forEach(cust => {
    if (!cust.email) return;
    try {
      _trackEmailsForCustomer(cust);
      tracked++;
    } catch (e) {
      logError(`メール追跡失敗: ${cust.customer_id}`, e);
    }
  });
  logInfo(`Gmailスキャン完了: ${tracked}件処理`);
}

function _trackEmailsForCustomer(customer) {
  const query = `from:${customer.email} OR to:${customer.email}`;
  const threads = GmailApp.search(query, 0, 5);
  if (!threads.length) return;

  const existingData = getSheetData(EMAIL_TRACKING_SHEET);
  const existingIds = new Set(existingData.map(r => r.thread_id));

  threads.forEach(thread => {
    const threadId = thread.getId();
    const messages = thread.getMessages();
    if (!messages.length) return;

    const lastMsg = messages[messages.length - 1];
    const lastDate = lastMsg.getDate();
    const lastFrom = lastMsg.getFrom();
    const direction = lastFrom.includes(customer.email) ? 'RECEIVED' : 'SENT';

    if (existingIds.has(threadId)) {
      // 既存スレッドを更新
      updateRow(EMAIL_TRACKING_SHEET, 'thread_id', threadId, {
        last_email_date: lastDate,
        last_sender: lastFrom,
        direction,
        tracked_at: now(),
      });
    } else {
      // 新規スレッドを追加
      const row = {
        thread_id: threadId,
        customer_id: customer.customer_id,
        company_name: customer.company_name,
        subject: thread.getFirstMessageSubject(),
        last_email_date: lastDate,
        last_sender: lastFrom,
        direction,
        tracked_at: now(),
      };
      appendRow(EMAIL_TRACKING_SHEET, row);
    }

    // 顧客の最終コンタクト日を更新
    const existingContact = customer.last_contact_date
      ? new Date(customer.last_contact_date) : new Date(0);
    if (lastDate > existingContact) {
      updateRow('👥 Customers', 'customer_id', customer.customer_id, {
        last_contact_date: lastDate,
        updated_at: now(),
      });
    }
  });
}

/**
 * CRMからメール送信
 */
function sendEmail(data) {
  if (!isEnabled('FEATURE_EMAIL_SEND')) return errorResponse('メール送信機能は無効です。');
  if (!data.to || !data.subject || !data.body) return errorResponse('宛先・件名・本文は必須です。');

  try {
    GmailApp.sendEmail(data.to, data.subject, data.body, {
      htmlBody: data.htmlBody || data.body,
      name: data.senderName || 'CRMシステム',
    });

    // 活動ログに記録
    if (data.customer_id) {
      createActivity({
        customer_id: data.customer_id,
        deal_id: data.deal_id || '',
        type: 'EMAIL',
        subject: `メール送信: ${data.subject}`,
        content: data.body,
        activity_date: now(),
      });
    }
    logInfo(`メール送信: ${data.to} - ${data.subject}`);
    return successResponse({ sent: true, to: data.to });
  } catch (e) {
    logError('メール送信エラー', e);
    return errorResponse(`送信失敗: ${e.message}`);
  }
}

function getEmailHistory(customerId) {
  const data = getSheetData(EMAIL_TRACKING_SHEET);
  return successResponse(
    data.filter(r => r.customer_id === customerId)
      .sort((a, b) => new Date(b.last_email_date) - new Date(a.last_email_date))
  );
}
