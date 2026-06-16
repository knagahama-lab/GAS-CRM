/**
 * CustomerModule.gs
 * 顧客管理 CRUD・重複チェック
 */

const CUSTOMERS_SHEET = '👥 Customers';

function createCustomer(data) {
  return wrapAction(() => {
    requireFeature('FEATURE_CUSTOMERS');

    // バリデーション
    const valErr = validateRequired(data, ['company_name', 'email']);
    if (valErr) return errorResponse(valErr.error);

    // サニタイズ
    data.company_name = sanitizeString(data.company_name);
    data.industry     = sanitizeString(data.industry);
    data.address      = sanitizeString(data.address);
    data.notes        = sanitizeString(data.notes);

    // 重複チェック
    if (isEnabled('FEATURE_DEDUP')) {
      const dup = checkDuplicateCustomer(data);
      if (dup) return errorResponse(`重複する顧客が存在します: ${dup.company_name}（ID: ${dup.customer_id}）`);
    }

    const id = generateId('CUST');
    const row = {
      customer_id: id,
      company_name: data.company_name,
      industry: data.industry || '',
      phone: normalizePhone(data.phone),
      email: data.email || '',
      address: data.address || '',
      assigned_user: data.assigned_user || Session.getActiveUser().getEmail(),
      status: data.status || 'PROSPECT',
      last_contact_date: '',
      notes: data.notes || '',
      created_at: now(),
      updated_at: now(),
    };

    appendRow(CUSTOMERS_SHEET, row);
    logInfo(`顧客作成: ${id} - ${data.company_name}`);
    return successResponse({ customer_id: id });
  });
}

function getCustomer(id) {
  return wrapAction(() => {
    requireFeature('FEATURE_CUSTOMERS');
    const customer = findRow(CUSTOMERS_SHEET, 'customer_id', id);
    if (!customer) return errorResponse(`顧客ID「${id}」が見つかりません。`);
    return successResponse(customer);
  });
}

function updateCustomer(id, data) {
  return wrapAction(() => {
    requireFeature('FEATURE_CUSTOMERS');

    const valErr = validateRequired({ id }, ['id']);
    if (valErr) return errorResponse(valErr.error);

    // サニタイズ
    if (data.company_name !== undefined) data.company_name = sanitizeString(data.company_name);
    if (data.industry !== undefined)     data.industry     = sanitizeString(data.industry);
    if (data.address !== undefined)      data.address      = sanitizeString(data.address);
    if (data.notes !== undefined)        data.notes        = sanitizeString(data.notes);

    data.updated_at = now();
    const result = updateRow(CUSTOMERS_SHEET, 'customer_id', id, data);
    if (!result) return errorResponse(`顧客ID「${id}」が見つかりません。`);
    logInfo(`顧客更新: ${id}`);
    return successResponse({ customer_id: id });
  });
}

function getAllCustomers(userEmail) {
  return wrapAction(() => {
    requireFeature('FEATURE_CUSTOMERS');
    let data = getSheetData(CUSTOMERS_SHEET);
    if (!_isAdmin(userEmail)) {
      data = data.filter(r => r.assigned_user === userEmail);
    }
    return successResponse(data);
  });
}

/**
 * 検索・フィルタ・ソート・ページネーション付き顧客一覧
 */
function getCustomersFiltered(params) {
  return wrapAction(() => {
    requireFeature('FEATURE_CUSTOMERS');
    const { search, status, sortKey, sortAsc, page, pageSize } = params || {};
    const userEmail = Session.getActiveUser().getEmail();
    let data = getSheetData(CUSTOMERS_SHEET);

    if (!_isAdmin(userEmail)) {
      data = data.filter(r => r.assigned_user === userEmail);
    }
    if (search) {
      const q = String(search).toLowerCase();
      data = data.filter(r =>
        String(r.company_name).toLowerCase().includes(q) ||
        String(r.email).toLowerCase().includes(q) ||
        String(r.phone).includes(q)
      );
    }
    if (status) {
      data = data.filter(r => r.status === status);
    }
    if (sortKey) {
      data = sortBy(data, sortKey, sortAsc !== false);
    }
    const paged = paginate(data, page, pageSize);
    return successResponse(paged);
  });
}

function searchCustomers(query) {
  return wrapAction(() => {
    requireFeature('FEATURE_CUSTOMERS');
    const q = String(query).toLowerCase();
    const data = getSheetData(CUSTOMERS_SHEET);
    const results = data.filter(r =>
      String(r.company_name).toLowerCase().includes(q) ||
      String(r.email).toLowerCase().includes(q) ||
      String(r.phone).includes(q)
    );
    return successResponse(results);
  });
}

function checkDuplicateCustomer(data) {
  const allData = getSheetData(CUSTOMERS_SHEET);
  const phone = normalizePhone(data.phone);
  return allData.find(r =>
    (data.company_name && r.company_name === data.company_name) ||
    (data.email && r.email === data.email) ||
    (phone && normalizePhone(r.phone) === phone)
  ) || null;
}

function getCustomerStats() {
  return wrapAction(() => {
    const data = getSheetData(CUSTOMERS_SHEET);
    const total = data.length;
    const byStatus = {};
    data.forEach(r => {
      byStatus[r.status] = (byStatus[r.status] || 0) + 1;
    });
    return successResponse({ total, byStatus });
  });
}
