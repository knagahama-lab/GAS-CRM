/**
 * UserModule.gs
 * ユーザー管理・権限チェック
 */

const USERS_SHEET = '👤 Users';

function getCurrentUser() {
  const email = Session.getActiveUser().getEmail();
  const users = getSheetData(USERS_SHEET);
  const user = users.find(u => u.email === email && (u.active === 'TRUE' || u.active === true));
  if (!user) return { email, role: 'VIEWER', name: email };
  return user;
}

function _isAdmin(email) {
  const adminEmails = getSetting('ADMIN_EMAILS', '');
  if (adminEmails.includes(email)) return true;
  const users = getSheetData(USERS_SHEET);
  const user = users.find(u => u.email === email);
  return user && user.role === 'ADMIN';
}

function _isManager(email) {
  if (_isAdmin(email)) return true;
  const users = getSheetData(USERS_SHEET);
  const user = users.find(u => u.email === email);
  return user && (user.role === 'MANAGER');
}

const ALLOWED_ROLES = ['ADMIN', 'MANAGER', 'SALES', 'VIEWER'];

function createUser(data) {
  return wrapAction(() => {
    _requireAdmin();

    const valErr = validateRequired(data, ['name', 'email']);
    if (valErr) return errorResponse(valErr.error);

    const role = (data.role || 'SALES').toUpperCase();
    if (!ALLOWED_ROLES.includes(role)) {
      return errorResponse(`ロール「${role}」は無効です。許可値: ${ALLOWED_ROLES.join(', ')}`);
    }

    const id = generateId('USR');
    appendRow(USERS_SHEET, {
      user_id: id,
      name: sanitizeString(data.name),
      email: data.email || '',
      role,
      area: sanitizeString(data.area),
      active: 'TRUE',
      created_at: now(),
    });
    return successResponse({ user_id: id });
  });
}

function getAllUsers() {
  return wrapAction(() => {
    return successResponse(getSheetData(USERS_SHEET));
  });
}

function updateUser(id, data) {
  return wrapAction(() => {
    _requireAdmin();
    if (data.role !== undefined) {
      const role = String(data.role).toUpperCase();
      if (!ALLOWED_ROLES.includes(role)) {
        return errorResponse(`ロール「${role}」は無効です。許可値: ${ALLOWED_ROLES.join(', ')}`);
      }
      data.role = role;
    }
    if (data.name !== undefined) data.name = sanitizeString(data.name);
    updateRow(USERS_SHEET, 'user_id', id, data);
    return successResponse({ user_id: id });
  });
}

function getMasterData(category) {
  return wrapAction(() => {
    const data = getSheetData('📋 MasterData');
    const filtered = data.filter(r =>
      r.category === category && (r.active === 'TRUE' || r.active === true)
    );
    filtered.sort((a, b) => Number(a.sort_order) - Number(b.sort_order));
    return successResponse(filtered);
  });
}
