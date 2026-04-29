function notifyTaskChange(projectId, task, action, actorEmail) {
  if (!shouldNotifyTask(action)) {
    return;
  }

  const recipients = getNotificationRecipients(projectId, task.assignee_email);
  const context = {
    action,
    task_name: task.task_name,
    assignee_email: task.assignee_email,
    progress_percent: task.progress_percent,
    status: task.status,
    actor_email: actorEmail
  };
  const subject = applyTemplate(
    getSetting('notify_task_subject_template') || '[Progress Board] タスク更新: {{task_name}}',
    context
  );
  const body = applyTemplate(
    getSetting('notify_task_body_template') || [
      'タスクが更新されました。',
      '',
      '操作: {{action}}',
      'タスク: {{task_name}}',
      '担当: {{assignee_email}}',
      '進捗: {{progress_percent}}%',
      '状態: {{status}}',
      '更新者: {{actor_email}}'
    ].join('\n'),
    context
  );

  sendMailSafely(recipients, subject, body);
}

function notifyIssueChange(issue, actorEmail) {
  if (!isNotifyEnabled() || getSetting('notify_issue_change') === 'FALSE') {
    return;
  }

  const recipients = getNotificationRecipients(issue.project_id, issue.owner_email);
  const context = {
    issue_title: issue.title,
    severity: issue.severity,
    status: issue.status,
    due_date: issue.due_date,
    owner_email: issue.owner_email,
    actor_email: actorEmail
  };
  const subject = applyTemplate(
    getSetting('notify_issue_subject_template') || '[Progress Board] 課題更新: {{issue_title}}',
    context
  );
  const body = applyTemplate(
    getSetting('notify_issue_body_template') || [
      '課題が更新されました。',
      '',
      '課題: {{issue_title}}',
      '重要度: {{severity}}',
      '状態: {{status}}',
      '期限: {{due_date}}',
      '担当: {{owner_email}}',
      '更新者: {{actor_email}}'
    ].join('\n'),
    context
  );

  sendMailSafely(recipients, subject, body);
}

function notifyMemberChange(projectId, member, actorEmail) {
  if (!isNotifyEnabled() || getSetting('notify_member_change') === 'FALSE') {
    return;
  }

  const subject = '[Progress Board] メンバー更新: ' + member.user_email;
  const body = [
    'プロジェクトメンバーが更新されました。',
    '',
    'ユーザー: ' + member.user_email,
    'ロール: ' + member.role,
    '更新者: ' + actorEmail
  ].join('\n');

  sendMailSafely(uniqueEmails([member.user_email].concat(getProjectAdminEmails(projectId))), subject, body);
}

function notifyOverdueItems(projectId) {
  if (!isNotifyEnabled() || getSetting('notify_overdue_daily') === 'FALSE') {
    return;
  }

  const detail = getProjectDetail(projectId);
  const today = todayString();
  const overdueTasks = detail.tasks.filter((task) => task.planned_end && task.planned_end < today && Number(task.progress_percent || 0) < 100);
  const overdueIssues = detail.issues.filter((issue) => issue.due_date && issue.due_date < today && issue.status !== '完了');

  if (!overdueTasks.length && !overdueIssues.length) {
    return;
  }

  const context = {
    project_name: detail.project.project_name,
    overdue_task_count: String(overdueTasks.length),
    overdue_issue_count: String(overdueIssues.length),
    overdue_tasks: overdueTasks.length ? overdueTasks.map((task) => '- ' + task.task_name + ' / ' + task.assignee_email + ' / ' + task.planned_end).join('\n') : '- なし',
    overdue_issues: overdueIssues.length ? overdueIssues.map((issue) => '- ' + issue.title + ' / ' + issue.owner_email + ' / ' + issue.due_date).join('\n') : '- なし'
  };
  const subject = applyTemplate(
    getSetting('notify_overdue_subject_template') || '[Progress Board] 期限超過通知: {{project_name}}',
    context
  );
  const body = applyTemplate(
    getSetting('notify_overdue_body_template') || [
      '期限超過の項目があります。',
      '',
      'プロジェクト: {{project_name}}',
      '',
      '期限超過タスク:',
      '{{overdue_tasks}}',
      '',
      '期限超過課題:',
      '{{overdue_issues}}'
    ].join('\n'),
    context
  );

  sendMailSafely(getProjectAdminEmails(projectId), subject, body);
  writeAuditLog('system', 'SEND_OVERDUE_NOTIFICATION', 'project', projectId, null, { tasks: overdueTasks.length, issues: overdueIssues.length });
}

function shouldNotifyTask(action) {
  if (!isNotifyEnabled()) {
    return false;
  }
  if (action === 'CREATE_TASK') {
    return getSetting('notify_task_create') !== 'FALSE';
  }
  if (action === 'DELETE_TASK') {
    return getSetting('notify_task_delete') !== 'FALSE';
  }
  return getSetting('notify_task_update') !== 'FALSE';
}

function getNotificationRecipients(projectId, ownerEmail) {
  const mode = getSetting('notify_recipients') || 'assignee_admins';
  if (mode === 'all') {
    return getProjectEmails(projectId);
  }
  if (mode === 'admins') {
    return getProjectAdminEmails(projectId);
  }
  return uniqueEmails([ownerEmail].concat(getProjectAdminEmails(projectId)));
}

function applyTemplate(template, context) {
  return String(template || '').replace(/\{\{([a-zA-Z0-9_]+)\}\}/g, (match, key) => {
    return Object.prototype.hasOwnProperty.call(context, key) ? String(context[key]) : match;
  });
}

function sendMailSafely(recipients, subject, body) {
  const to = uniqueEmails(recipients).join(',');
  if (!to) {
    return;
  }

  try {
    MailApp.sendEmail({
      to,
      subject,
      body,
      name: 'Progress Board'
    });
  } catch (error) {
    writeAuditLog('system', 'MAIL_FAILED', 'notification', subject, null, { message: error.message });
  }
}

function isNotifyEnabled() {
  return getSetting('notify_enabled') !== 'FALSE';
}

function uniqueEmails(emails) {
  const seen = {};
  return (emails || [])
    .map(normalizeEmail)
    .filter((email) => {
      if (!email || seen[email]) {
        return false;
      }
      seen[email] = true;
      return true;
    });
}

function extractMentions(text) {
  const pattern = /@([\w.+\-]+@[\w.\-]+\.[a-zA-Z]{2,})/g;
  const emails = [];
  let match;
  while ((match = pattern.exec(String(text || ''))) !== null) {
    emails.push(match[1].toLowerCase());
  }
  return emails;
}

function notifyMentionInComment(issue, commentText, mentionedEmails, actorEmail) {
  if (!isNotifyEnabled()) return;
  const to = uniqueEmails(
    mentionedEmails.filter((e) => normalizeEmail(e) !== normalizeEmail(actorEmail))
  );
  if (!to.length) return;

  const subject = '[Progress Board] メンションされました: ' + issue.title;
  const body = [
    actorEmail + ' さんが課題コメントであなたをメンションしました。',
    '',
    '課題: ' + issue.title,
    '',
    'コメント:',
    commentText
  ].join('\n');

  sendMailSafely(to, subject, body);
  writeAuditLog('system', 'NOTIFY_MENTION', 'issue', issue.issue_id, null, { to: to.join(',') });
}
