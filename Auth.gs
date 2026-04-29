function getCurrentUser() {
  const email = Session.getActiveUser().getEmail();
  return {
    email,
    name: email ? email.split('@')[0] : 'unknown'
  };
}

function assertWorkspaceUser(email) {
  if (!email) {
    throw new Error('Google Workspaceアカウントでログインしてください。');
  }

  const allowedDomain = getSetting('allowed_domain');
  if (allowedDomain && !email.toLowerCase().endsWith('@' + allowedDomain.toLowerCase())) {
    throw new Error('許可された社内ドメインのユーザーのみ利用できます。');
  }
}

function assertProjectMember(projectId, email) {
  assertWorkspaceUser(email);

  const members = readTable(SHEETS.members);
  const allowed = members.some((member) =>
    member.project_id === projectId &&
    normalizeEmail(member.user_email) === normalizeEmail(email) &&
    member.active !== 'FALSE'
  );

  if (!allowed) {
    throw new Error('このプロジェクトへのアクセス権がありません。');
  }
}

function isProjectAdmin(projectId, email) {
  const members = readTable(SHEETS.members);
  return members.some((member) =>
    member.project_id === projectId &&
    normalizeEmail(member.user_email) === normalizeEmail(email) &&
    ['admin', 'owner'].indexOf(String(member.role).toLowerCase()) >= 0 &&
    member.active !== 'FALSE'
  );
}

function assertProjectAdmin(projectId, email) {
  assertProjectMember(projectId, email);

  if (!isProjectAdmin(projectId, email)) {
    throw new Error('この操作にはプロジェクト管理者権限が必要です。');
  }
}

function assertDeletedProjectAdmin(projectId, email) {
  assertWorkspaceUser(email);

  const members = readAllRows(SHEETS.members);
  const allowed = members.some((member) =>
    member.project_id === projectId &&
    normalizeEmail(member.user_email) === normalizeEmail(email) &&
    ['admin', 'owner'].indexOf(String(member.role).toLowerCase()) >= 0 &&
    member.active !== 'FALSE'
  );

  if (!allowed) {
    throw new Error('この操作にはプロジェクト管理者権限が必要です。');
  }
}

function normalizeEmail(email) {
  return String(email || '').trim().toLowerCase();
}
