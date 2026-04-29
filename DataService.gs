const HEADERS = {
  Projects: ['project_id', 'project_name', 'client_name', 'status', 'start_date', 'end_date', 'owner_email', 'created_at', 'updated_at', 'deleted'],
  Members: ['project_id', 'user_email', 'role', 'active'],
  Phases: ['phase_id', 'project_id', 'phase_name', 'planned_start', 'planned_end', 'actual_start', 'actual_end', 'progress_percent', 'status', 'assignee_email', 'sort_order', 'deleted'],
  Tasks: ['task_id', 'phase_id', 'task_name', 'planned_start', 'planned_end', 'actual_start', 'actual_end', 'progress_percent', 'status', 'priority', 'blocked_reason', 'depends_on', 'assignee_email', 'summary', 'created_at', 'updated_at', 'deleted'],
  Issues: ['issue_id', 'project_id', 'title', 'severity', 'status', 'owner_email', 'due_date', 'summary', 'created_at', 'updated_at', 'deleted'],
  AuditLogs: ['timestamp', 'user_email', 'action', 'target_type', 'target_id', 'before', 'after'],
  IssueComments: ['comment_id', 'issue_id', 'user_email', 'comment', 'created_at', 'updated_at', 'deleted'],
  Archives: ['archive_id', 'project_id', 'project_name', 'archived_at', 'archived_by', 'snapshot'],
  Settings: ['key', 'value']
};

function getSheet(name) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(name);
  if (!sheet) {
    throw new Error(name + ' シートが見つかりません。setupSampleSpreadsheet を実行してください。');
  }
  return sheet;
}

function ensureSheet(name) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(name);
    sheet.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS[name].length).setFontWeight('bold').setBackground('#eef2f7');
  }
  return sheet;
}

function readTable(name) {
  return readAllRows(name).filter((record) => record.deleted !== 'TRUE' && record.deleted !== 'ARCHIVED');
}

function readAllRows(name) {
  ensureSheet(name);
  const sheet = getSheet(name);
  const values = sheet.getDataRange().getDisplayValues();

  if (values.length < 2) {
    return [];
  }

  const headers = values[0];
  return values.slice(1)
    .filter((row) => row.some(Boolean))
    .map((row) => headers.reduce((record, header, index) => {
      record[header] = row[index] || '';
      return record;
    }, {}));
}

function appendRecord(sheet, record) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  sheet.appendRow(headers.map((header) => record[header] || ''));
}

function writeRecordAt(sheet, rowNumber, record) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  sheet.getRange(rowNumber, 1, 1, headers.length).setValues([headers.map((header) => record[header] || '')]);
}

function findRecordById(sheetName, idColumn, id) {
  return readTable(sheetName).find((record) => record[idColumn] === id);
}

function getVisibleProjectsForUser(email) {
  const projectIds = readTable(SHEETS.members)
    .filter((member) => normalizeEmail(member.user_email) === normalizeEmail(email) && member.active !== 'FALSE')
    .map((member) => member.project_id);

  return readTable(SHEETS.projects)
    .filter((project) => projectIds.indexOf(project.project_id) >= 0)
    .sort((a, b) => a.project_name.localeCompare(b.project_name, 'ja'));
}

function getDeletedProjectsForUser(email) {
  const adminProjectIds = readAllRows(SHEETS.members)
    .filter((member) =>
      normalizeEmail(member.user_email) === normalizeEmail(email) &&
      member.active !== 'FALSE' &&
      ['admin', 'owner'].indexOf(String(member.role).toLowerCase()) >= 0
    )
    .map((member) => member.project_id);

  return readAllRows(SHEETS.projects)
    .filter((project) => project.deleted === 'TRUE' && adminProjectIds.indexOf(project.project_id) >= 0)
    .sort((a, b) => a.project_name.localeCompare(b.project_name, 'ja'));
}

function getProjectDetail(projectId) {

  const project = findRecordById(SHEETS.projects, 'project_id', projectId);
  if (!project) {
    throw new Error('プロジェクトが見つかりません。');
  }

  const phases = readTable(SHEETS.phases)
    .filter((phase) => phase.project_id === projectId)
    .sort((a, b) => Number(a.sort_order || 0) - Number(b.sort_order || 0));

  const phaseIds = phases.map((phase) => phase.phase_id);
  const tasks = readTable(SHEETS.tasks)
    .filter((task) => phaseIds.indexOf(task.phase_id) >= 0)
    .sort((a, b) => String(a.planned_start).localeCompare(String(b.planned_start)));
  const issues = readTable(SHEETS.issues)
    .filter((issue) => issue.project_id === projectId)
    .sort((a, b) => String(a.due_date).localeCompare(String(b.due_date)));
  const issueIds = issues.map((issue) => issue.issue_id);
  const issueComments = readTable(SHEETS.issueComments)
    .filter((comment) => issueIds.indexOf(comment.issue_id) >= 0)
    .sort((a, b) => String(a.created_at).localeCompare(String(b.created_at)));
  const members = readTable(SHEETS.members)
    .filter((member) => member.project_id === projectId && member.active !== 'FALSE')
    .sort((a, b) => a.user_email.localeCompare(b.user_email));

  return {
    project,
    phases,
    tasks,
    issues,
    issueComments,
    members,
    summary: buildSummary(project, phases, tasks, issues)
  };
}

function buildSummary(project, phases, tasks, issues) {
  const activeTasks = tasks.filter((task) => task.deleted !== 'TRUE');
  const doneTasks = activeTasks.filter((task) => Number(task.progress_percent || 0) >= 100);
  const blockedTasks = activeTasks.filter((task) => task.blocked_reason);
  const delayedTasks = activeTasks.filter((task) => task.planned_end && task.planned_end < todayString() && Number(task.progress_percent || 0) < 100);
  const openIssues = issues.filter((issue) => issue.status !== '完了');
  const averageProgress = activeTasks.length
    ? Math.round(activeTasks.reduce((sum, task) => sum + Number(task.progress_percent || 0), 0) / activeTasks.length)
    : 0;

  return {
    progress: averageProgress,
    taskCount: activeTasks.length,
    doneTaskCount: doneTasks.length,
    delayedTaskCount: delayedTasks.length,
    blockedTaskCount: blockedTasks.length,
    openIssueCount: openIssues.length,
    phaseCount: phases.length,
    projectStatus: project.status
  };
}

function recalculatePhaseProgress(projectId) {
  const phaseSheet = getSheet(SHEETS.phases);
  const phases = readAllRows(SHEETS.phases);
  const tasks = readTable(SHEETS.tasks);

  phases.forEach((phase, index) => {
    if (phase.project_id !== projectId || phase.deleted === 'TRUE') {
      return;
    }

    const phaseTasks = tasks.filter((task) => task.phase_id === phase.phase_id);
    const progress = phaseTasks.length
      ? Math.round(phaseTasks.reduce((sum, task) => sum + Number(task.progress_percent || 0), 0) / phaseTasks.length)
      : Number(phase.progress_percent || 0);
    const updated = Object.assign({}, phase, {
      progress_percent: String(progress),
      status: progress >= 100 ? '完了' : progress > 0 ? '進行中' : '未着手'
    });
    writeRecordAt(phaseSheet, index + 2, updated);
  });
}

function getProjectEmails(projectId) {
  return readTable(SHEETS.members)
    .filter((member) => member.project_id === projectId && member.active !== 'FALSE')
    .map((member) => member.user_email)
    .filter(Boolean);
}

function getProjectAdminEmails(projectId) {
  return readTable(SHEETS.members)
    .filter((member) =>
      member.project_id === projectId &&
      member.active !== 'FALSE' &&
      ['admin', 'owner'].indexOf(String(member.role).toLowerCase()) >= 0
    )
    .map((member) => member.user_email)
    .filter(Boolean);
}

function buildProjectCsv(detail) {
  const phaseById = detail.phases.reduce((map, phase) => {
    map[phase.phase_id] = phase.phase_name;
    return map;
  }, {});
  const rows = [
    ['type', 'phase', 'name', 'assignee_or_owner', 'planned_start', 'planned_end', 'actual_start', 'actual_end', 'progress', 'status', 'priority_or_severity', 'note']
  ];

  detail.phases.forEach((phase) => rows.push(['phase', phase.phase_name, phase.phase_name, phase.assignee_email, phase.planned_start, phase.planned_end, phase.actual_start, phase.actual_end, phase.progress_percent, phase.status, '', '']));
  detail.tasks.forEach((task) => rows.push(['task', phaseById[task.phase_id] || '', task.task_name, task.assignee_email, task.planned_start, task.planned_end, task.actual_start, task.actual_end, task.progress_percent, task.status, task.priority, task.blocked_reason]));
  detail.issues.forEach((issue) => rows.push(['issue', '', issue.title, issue.owner_email, '', issue.due_date, '', '', '', issue.status, issue.severity, '']));

  return rows.map((row) => row.map(csvEscape).join(',')).join('\n');
}

function csvEscape(value) {
  const text = String(value || '');
  return /[",\n]/.test(text) ? '"' + text.replace(/"/g, '""') + '"' : text;
}

function getSetting(key) {
  const records = readTable(SHEETS.settings);
  const setting = records.find((record) => record.key === key);
  return setting ? setting.value : '';
}

function setSetting(key, value) {
  const sheet = getSheet(SHEETS.settings);
  const rows = readAllRows(SHEETS.settings);
  const index = rows.findIndex((row) => row.key === key);
  const record = { key, value: String(value || '') };

  if (index >= 0) {
    writeRecordAt(sheet, index + 2, record);
  } else {
    appendRecord(sheet, record);
  }
}

function getSettingsMap() {
  return readTable(SHEETS.settings).reduce((settings, row) => {
    settings[row.key] = row.value;
    return settings;
  }, {});
}

function getNotificationSettings() {
  const settings = getSettingsMap();
  return {
    notify_enabled: settings.notify_enabled || 'TRUE',
    notify_task_create: settings.notify_task_create || 'TRUE',
    notify_task_update: settings.notify_task_update || 'TRUE',
    notify_task_delete: settings.notify_task_delete || 'TRUE',
    notify_issue_change: settings.notify_issue_change || 'TRUE',
    notify_member_change: settings.notify_member_change || 'TRUE',
    notify_overdue_daily: settings.notify_overdue_daily || 'FALSE',
    notify_overdue_hour: settings.notify_overdue_hour || '9',
    notify_recipients: settings.notify_recipients || 'assignee_admins',
    notify_overdue_project_id: settings.notify_overdue_project_id || '',
    notify_task_subject_template: settings.notify_task_subject_template || '[Progress Board] タスク更新: {{task_name}}',
    notify_task_body_template: settings.notify_task_body_template || 'タスクが更新されました。\n\n操作: {{action}}\nタスク: {{task_name}}\n担当: {{assignee_email}}\n進捗: {{progress_percent}}%\n状態: {{status}}\n更新者: {{actor_email}}',
    notify_issue_subject_template: settings.notify_issue_subject_template || '[Progress Board] 課題更新: {{issue_title}}',
    notify_issue_body_template: settings.notify_issue_body_template || '課題が更新されました。\n\n課題: {{issue_title}}\n重要度: {{severity}}\n状態: {{status}}\n期限: {{due_date}}\n担当: {{owner_email}}\n更新者: {{actor_email}}',
    notify_overdue_subject_template: settings.notify_overdue_subject_template || '[Progress Board] 期限超過通知: {{project_name}}',
    notify_overdue_body_template: settings.notify_overdue_body_template || '期限超過の項目があります。\n\nプロジェクト: {{project_name}}\n\n期限超過タスク:\n{{overdue_tasks}}\n\n期限超過課題:\n{{overdue_issues}}'
  };
}

function initializeSheets() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(HEADERS).forEach((name) => {
    let sheet = spreadsheet.getSheetByName(name);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(name);
    }
    sheet.clear();
    sheet.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS[name].length).setFontWeight('bold').setBackground('#eef2f7');
  });
}

function seedSampleData(ownerEmail) {
  const now = new Date().toISOString();
  const domain = ownerEmail.indexOf('@') >= 0 ? ownerEmail.split('@')[1] : '';
  const projectId = 'pj-sample';
  const phaseIds = ['phase-plan', 'phase-design', 'phase-build', 'phase-test', 'phase-release'];

  [
    ['allowed_domain', domain],
    ['notify_enabled', 'TRUE'],
    ['notify_task_create', 'TRUE'],
    ['notify_task_update', 'TRUE'],
    ['notify_task_delete', 'TRUE'],
    ['notify_issue_change', 'TRUE'],
    ['notify_member_change', 'TRUE'],
    ['notify_overdue_daily', 'FALSE'],
    ['notify_overdue_hour', '9'],
    ['notify_recipients', 'assignee_admins'],
    ['notify_task_subject_template', '[Progress Board] タスク更新: {{task_name}}'],
    ['notify_task_body_template', 'タスクが更新されました。\n\n操作: {{action}}\nタスク: {{task_name}}\n担当: {{assignee_email}}\n進捗: {{progress_percent}}%\n状態: {{status}}\n更新者: {{actor_email}}'],
    ['notify_issue_subject_template', '[Progress Board] 課題更新: {{issue_title}}'],
    ['notify_issue_body_template', '課題が更新されました。\n\n課題: {{issue_title}}\n重要度: {{severity}}\n状態: {{status}}\n期限: {{due_date}}\n担当: {{owner_email}}\n更新者: {{actor_email}}'],
    ['notify_overdue_subject_template', '[Progress Board] 期限超過通知: {{project_name}}'],
    ['notify_overdue_body_template', '期限超過の項目があります。\n\nプロジェクト: {{project_name}}\n\n期限超過タスク:\n{{overdue_tasks}}\n\n期限超過課題:\n{{overdue_issues}}']
  ].forEach((setting) => appendRecord(getSheet(SHEETS.settings), { key: setting[0], value: setting[1] }));

  appendRecord(getSheet(SHEETS.projects), {
    project_id: projectId,
    project_name: '社内進捗管理MVP',
    client_name: 'Internal',
    status: '進行中',
    start_date: '2026-05-01',
    end_date: '2026-06-14',
    owner_email: ownerEmail,
    created_at: now,
    updated_at: now,
    deleted: ''
  });
  appendRecord(getSheet(SHEETS.members), { project_id: projectId, user_email: ownerEmail, role: 'owner', active: 'TRUE' });

  const phases = [
    ['要件定義', '2026-05-01', '2026-05-07', '2026-05-01', '2026-05-06', 100, '完了', 1],
    ['基本設計', '2026-05-08', '2026-05-16', '2026-05-08', '', 72, '進行中', 2],
    ['実装', '2026-05-17', '2026-05-31', '', '', 28, '進行中', 3],
    ['テスト', '2026-06-01', '2026-06-09', '', '', 0, '未着手', 4],
    ['リリース', '2026-06-10', '2026-06-14', '', '', 0, '未着手', 5]
  ];

  phases.forEach((phase, index) => appendRecord(getSheet(SHEETS.phases), {
    phase_id: phaseIds[index],
    project_id: projectId,
    phase_name: phase[0],
    planned_start: phase[1],
    planned_end: phase[2],
    actual_start: phase[3],
    actual_end: phase[4],
    progress_percent: String(phase[5]),
    status: phase[6],
    assignee_email: ownerEmail,
    sort_order: String(phase[7]),
    deleted: ''
  }));

  [
    ['phase-plan', '業務フロー整理', '2026-05-01', '2026-05-03', '2026-05-01', '2026-05-03', 100, '完了', 'High', ''],
    ['phase-design', '画面ワイヤー作成', '2026-05-08', '2026-05-11', '2026-05-08', '', 80, '進行中', 'High', ''],
    ['phase-design', '権限設計レビュー', '2026-05-12', '2026-05-16', '', '', 45, '進行中', 'High', '管理者承認待ち'],
    ['phase-build', 'ガントチャート実装', '2026-05-17', '2026-05-24', '', '', 30, '進行中', 'Medium', ''],
    ['phase-build', '監査ログ実装', '2026-05-25', '2026-05-31', '', '', 0, '未着手', 'Medium', '']
  ].forEach((task) => appendRecord(getSheet(SHEETS.tasks), {
    task_id: Utilities.getUuid(),
    phase_id: task[0],
    task_name: task[1],
    planned_start: task[2],
    planned_end: task[3],
    actual_start: task[4],
    actual_end: task[5],
    progress_percent: String(task[6]),
    status: task[7],
    priority: task[8],
    blocked_reason: task[9],
    assignee_email: ownerEmail,
    created_at: now,
    updated_at: now,
    deleted: ''
  }));

  appendRecord(getSheet(SHEETS.issues), {
    issue_id: Utilities.getUuid(),
    project_id: projectId,
    title: 'メンバー追加フローの承認ルール確認',
    severity: 'Medium',
    status: '対応中',
    owner_email: ownerEmail,
    due_date: '2026-05-13',
    created_at: now,
    updated_at: now,
    deleted: ''
  });
}
