function validateProjectPayload(payload, fallbackOwnerEmail) {
  const project = payload || {};
  const required = ['project_name', 'client_name', 'status', 'start_date', 'end_date'];

  required.forEach((field) => {
    if (!String(project[field] || '').trim()) {
      throw new Error(field + ' は必須です。');
    }
  });

  if (project.start_date > project.end_date) {
    throw new Error('開始日は終了日以前にしてください。');
  }

  return {
    project_id: String(project.project_id || ''),
    project_name: sanitizeText(project.project_name, 120),
    client_name: sanitizeText(project.client_name, 80),
    status: sanitizeText(project.status || '計画中', 20),
    start_date: String(project.start_date),
    end_date: String(project.end_date),
    owner_email: normalizeEmail(project.owner_email || fallbackOwnerEmail)
  };
}

function validatePhasePayload(payload) {
  const phase = payload || {};
  const required = ['project_id', 'phase_name', 'planned_start', 'planned_end', 'assignee_email'];

  required.forEach((field) => {
    if (!String(phase[field] || '').trim()) {
      throw new Error(field + ' は必須です。');
    }
  });

  if (phase.planned_start > phase.planned_end) {
    throw new Error('予定開始日は予定終了日以前にしてください。');
  }

  return {
    phase_id: String(phase.phase_id || ''),
    project_id: String(phase.project_id),
    phase_name: sanitizeText(phase.phase_name, 80),
    planned_start: String(phase.planned_start),
    planned_end: String(phase.planned_end),
    actual_start: String(phase.actual_start || ''),
    actual_end: String(phase.actual_end || ''),
    progress_percent: String(clampNumber(Number(phase.progress_percent || 0), 0, 100)),
    status: sanitizeText(phase.status || '未着手', 20),
    assignee_email: normalizeEmail(phase.assignee_email),
    sort_order: String(clampNumber(Number(phase.sort_order || 999), 0, 9999))
  };
}

function validateTaskPayload(payload) {
  const task = payload || {};
  const required = ['phase_id', 'task_name', 'planned_start', 'planned_end', 'assignee_email'];

  required.forEach((field) => {
    if (!String(task[field] || '').trim()) {
      throw new Error(field + ' は必須です。');
    }
  });

  if (task.planned_start > task.planned_end) {
    throw new Error('予定開始日は予定終了日以前にしてください。');
  }

  return {
    task_id: String(task.task_id || ''),
    phase_id: String(task.phase_id),
    task_name: sanitizeText(task.task_name, 120),
    planned_start: String(task.planned_start),
    planned_end: String(task.planned_end),
    actual_start: String(task.actual_start || ''),
    actual_end: String(task.actual_end || ''),
    progress_percent: String(clampNumber(Number(task.progress_percent || 0), 0, 100)),
    status: sanitizeText(task.status || '未着手', 20),
    priority: sanitizeText(task.priority || 'Medium', 20),
    blocked_reason: sanitizeText(task.blocked_reason || '', 300),
    depends_on: sanitizeText(task.depends_on || '', 36),
    assignee_email: normalizeEmail(task.assignee_email)
  };
}

function validateMemberPayload(projectId, payload) {
  const member = payload || {};
  const email = normalizeEmail(member.user_email);

  if (!email || email.indexOf('@') < 0) {
    throw new Error('有効なメールアドレスを入力してください。');
  }

  const allowedDomain = getSetting('allowed_domain');
  if (allowedDomain && !email.endsWith('@' + allowedDomain.toLowerCase())) {
    throw new Error('社内ドメインのユーザーのみ追加できます。');
  }

  return {
    project_id: String(projectId),
    user_email: email,
    role: sanitizeText(member.role || 'member', 20),
    active: member.active === 'FALSE' ? 'FALSE' : 'TRUE'
  };
}

function validateIssuePayload(payload) {
  const issue = payload || {};
  const required = ['project_id', 'title', 'severity', 'status', 'owner_email', 'due_date'];

  required.forEach((field) => {
    if (!String(issue[field] || '').trim()) {
      throw new Error(field + ' は必須です。');
    }
  });

  return {
    issue_id: String(issue.issue_id || ''),
    project_id: String(issue.project_id),
    title: sanitizeText(issue.title, 160),
    severity: sanitizeText(issue.severity || 'Medium', 20),
    status: sanitizeText(issue.status || '未対応', 20),
    owner_email: normalizeEmail(issue.owner_email),
    due_date: String(issue.due_date)
  };
}

function validateIssueCommentPayload(payload) {
  const comment = payload || {};
  if (!String(comment.issue_id || '').trim()) {
    throw new Error('issue_id は必須です。');
  }
  if (!String(comment.comment || '').trim()) {
    throw new Error('コメントを入力してください。');
  }
  return {
    issue_id: String(comment.issue_id),
    comment: sanitizeText(comment.comment, 1000)
  };
}

function validateNotificationSettings(payload) {
  const settings = payload || {};
  return {
    notify_enabled: settings.notify_enabled === 'TRUE' ? 'TRUE' : 'FALSE',
    notify_task_create: settings.notify_task_create === 'TRUE' ? 'TRUE' : 'FALSE',
    notify_task_update: settings.notify_task_update === 'TRUE' ? 'TRUE' : 'FALSE',
    notify_task_delete: settings.notify_task_delete === 'TRUE' ? 'TRUE' : 'FALSE',
    notify_issue_change: settings.notify_issue_change === 'TRUE' ? 'TRUE' : 'FALSE',
    notify_member_change: settings.notify_member_change === 'TRUE' ? 'TRUE' : 'FALSE',
    notify_overdue_daily: settings.notify_overdue_daily === 'TRUE' ? 'TRUE' : 'FALSE',
    notify_overdue_hour: String(clampNumber(Number(settings.notify_overdue_hour || 9), 0, 23)),
    notify_recipients: ['all', 'assignee_admins', 'admins'].indexOf(settings.notify_recipients) >= 0 ? settings.notify_recipients : 'assignee_admins',
    notify_task_subject_template: sanitizeTemplate(settings.notify_task_subject_template, 180),
    notify_task_body_template: sanitizeTemplate(settings.notify_task_body_template, 2000),
    notify_issue_subject_template: sanitizeTemplate(settings.notify_issue_subject_template, 180),
    notify_issue_body_template: sanitizeTemplate(settings.notify_issue_body_template, 2000),
    notify_overdue_subject_template: sanitizeTemplate(settings.notify_overdue_subject_template, 180),
    notify_overdue_body_template: sanitizeTemplate(settings.notify_overdue_body_template, 3000)
  };
}

function sanitizeTemplate(value, maxLength) {
  return String(value || '')
    .replace(/[<>]/g, '')
    .trim()
    .slice(0, maxLength);
}

function sanitizeText(value, maxLength) {
  return String(value || '')
    .replace(/[<>]/g, '')
    .trim()
    .slice(0, maxLength);
}

function clampNumber(value, min, max) {
  if (Number.isNaN(value)) {
    return min;
  }
  return Math.max(min, Math.min(max, value));
}

function todayString() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
