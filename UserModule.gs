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

function createUser(data) {
  _requireAdmin();
  const id = generateId('USR');
  appendRow(USERS_SHEET, {
    user_id: id,
    name: data.name || '',
    email: data.email || '',
    role: data.role || 'SALES',
    area: data.area || '',
    active: 'TRUE',
    created_at: now(),
  });
  return successResponse({ user_id: id });
}

function getAllUsers() {
  return successResponse(getSheetData(USERS_SHEET));
}

function updateUser(id, data) {
  _requireAdmin();
  updateRow(USERS_SHEET, 'user_id', id, data);
  return successResponse({ user_id: id });
}

function getMasterData(category) {
  const data = getSheetData('📋 MasterData');
  const filtered = data.filter(r =>
    r.category === category && (r.active === 'TRUE' || r.active === true)
  );
  filtered.sort((a, b) => Number(a.sort_order) - Number(b.sort_order));
  return successResponse(filtered);
}
