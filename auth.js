// ══════════════════════════════════════════════════════════
//  🔐 Auth & Roles — مركزي لجميع الصفحات
// ══════════════════════════════════════════════════════════

const Auth = {
  // الصلاحيات لكل دور
  ROLES: {
    admin: {
      label: 'مسؤول',
      pages: ['admin.html','dashboard.html','scanner.html'],
      canCreateWorkshop : true,
      canDeleteWorkshop : true,
      canArchiveWorkshop: true,
      canManageQR       : true,
      canViewReports    : true,
      canExport         : true,
      canViewAllStats   : true,
    },
    organizer: {
      label: 'منظم',
      pages: ['scanner.html','dashboard.html'],
      canCreateWorkshop : false,
      canDeleteWorkshop : false,
      canArchiveWorkshop: false,
      canManageQR       : false,
      canViewReports    : false,
      canExport         : true,
      canViewAllStats   : false,
    },
  },

  // تسجيل دخول
  login(password) {
    for (const [role, pass] of Object.entries(CONFIG.PASSWORDS)) {
      if (password === pass) {
        sessionStorage.setItem('role', role);
        sessionStorage.setItem('auth', btoa(password + ':' + Date.now()));
        return role;
      }
    }
    return null;
  },

  // التحقق من الجلسة
  check(requiredRole) {
    const role = sessionStorage.getItem('role');
    const auth = sessionStorage.getItem('auth');
    if (!role || !auth) { this.redirect(); return false; }
    if (requiredRole && role !== requiredRole && !(role === 'admin')) {
      this.redirect(); return false;
    }
    return true;
  },

  // صلاحية محددة
  can(permission) {
    const role = sessionStorage.getItem('role');
    return this.ROLES[role]?.[permission] ?? false;
  },

  getRole() { return sessionStorage.getItem('role'); },

  logout() { sessionStorage.clear(); window.location.href = 'index.html'; },

  redirect() { sessionStorage.clear(); window.location.href = 'index.html'; },
};

// API مشتركة لجميع الصفحات
async function apiCall(payload) {
  const url = CONFIG.APPS_SCRIPT_URL + '?data=' + encodeURIComponent(JSON.stringify(payload));
  const r   = await fetch(url, { method: 'GET', redirect: 'follow' });
  if (!r.ok) throw new Error('HTTP ' + r.status);
  const text = await r.text();
  return JSON.parse(text);
}
