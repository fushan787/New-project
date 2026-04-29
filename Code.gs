const SHEETS = {
  projects: 'Projects',
  members: 'Members',
  phases: 'Phases',
  tasks: 'Tasks',
  issues: 'Issues',
  auditLogs: 'AuditLogs',
  issueComments: 'IssueComments',
  archives: 'Archives',
  settings: 'Settings'
};

function doGet() {
  const template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .setTitle('Progress Board')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getAppBootstrap() {
  const user = getCurrentUser();
  assertWorkspaceUser(user.email);

  const projects = getVisibleProjectsForUser(user.email);
  const activeProjectId = projects.length ? projects[0].project_id : '';
  const activeProject = activeProjectId ? getProjectDetail(activeProjectId) : null;

  return {
    user,
    projects,
    activeProject,
    deletedProjects: getDeletedProjectsForUser(user.email),
    notificationSettings: getNotificationSettings(),
    serverDate: todayString()
  };
}

function getProject(projectId) {
  const user = getCurrentUser();
  assertProjectMember(projectId, user.email);
  return getProjectDetail(projectId);
}

function saveProject(payload) {
  const user = getCurrentUser();
  assertWorkspaceUser(user.email);

  const project = validateProjectPayload(payload, user.email);
  const sheet = getSheet(SHEETS.projects);
  const records = readTable(SHEETS.projects);
  const existingIndex = records.findIndex((record) => record.project_id === project.project_id);
  const now = new Date().toISOString();
  let savedProject;
  let before = null;

  if (existingIndex >= 0) {
    before = records[existingIndex];
    assertProjectAdmin(project.project_id, user.email);
    savedProject = Object.assign({}, before, project, {
      updated_at: now,
      deleted: before.deleted || ''
    });
    writeRecordAt(sheet, existingIndex + 2, savedProject);
  } else {
    savedProject = Object.assign({}, project, {
      project_id: Utilities.getUuid(),
      created_at: now,
      updated_at: now,
      deleted: ''
    });
    appendRecord(sheet, savedProject);
    appendRecord(getSheet(SHEETS.members), {
      project_id: savedProject.project_id,
      user_email: user.email,
      role: 'owner',
      active: 'TRUE'
    });
  }

  writeAuditLog(user.email, existingIndex >= 0 ? 'UPDATE_PROJECT' : 'CREATE_PROJECT', 'project', savedProject.project_id, before, savedProject);

  const projects = getVisibleProjectsForUser(user.email);
  return {
    projects,
    activeProject: getProjectDetail(savedProject.project_id)
  };
}

function deleteProject(projectId) {
  const user = getCurrentUser();
  assertProjectAdmin(projectId, user.email);

  const sheet = getSheet(SHEETS.projects);
  const records = readAllRows(SHEETS.projects);
  const index = records.findIndex((record) => record.project_id === projectId && record.deleted !== 'TRUE');

  if (index < 0) {
    throw new Error('プロジェクトが見つかりません。');
  }

  const updated = Object.assign({}, records[index], {
    deleted: 'TRUE',
    updated_at: new Date().toISOString()
  });
  writeRecordAt(sheet, index + 2, updated);
  writeAuditLog(user.email, 'DELETE_PROJECT', 'project', projectId, records[index], updated);

  const projects = getVisibleProjectsForUser(user.email);
  return {
    projects,
    activeProject: projects.length ? getProjectDetail(projects[0].project_id) : null
  };
}

function restoreProject(projectId) {
  const user = getCurrentUser();
  assertDeletedProjectAdmin(projectId, user.email);

  const sheet = getSheet(SHEETS.projects);
  const records = readAllRows(SHEETS.projects);
  const index = records.findIndex((record) => record.project_id === projectId && record.deleted === 'TRUE');

  if (index < 0) {
    throw new Error('復元対象のプロジェクトが見つかりません。');
  }

  const restored = Object.assign({}, records[index], {
    deleted: '',
    updated_at: new Date().toISOString()
  });
  writeRecordAt(sheet, index + 2, restored);
  writeAuditLog(user.email, 'RESTORE_PROJECT', 'project', projectId, records[index], restored);

  const projects = getVisibleProjectsForUser(user.email);
  return {
    projects,
    activeProject: getProjectDetail(projectId),
    deletedProjects: getDeletedProjectsForUser(user.email)
  };
}

function archiveDeletedProject(projectId) {
  const user = getCurrentUser();
  assertDeletedProjectAdmin(projectId, user.email);

  const project = readAllRows(SHEETS.projects).find((record) => record.project_id === projectId && record.deleted === 'TRUE');
  if (!project) {
    throw new Error('アーカイブ対象の削除済みプロジェクトが見つかりません。');
  }

  ensureSheet(SHEETS.archives);
  const phases = readAllRows(SHEETS.phases).filter((phase) => phase.project_id === projectId);
  const phaseIds = phases.map((phase) => phase.phase_id);
  const issues = readAllRows(SHEETS.issues).filter((issue) => issue.project_id === projectId);
  const issueIds = issues.map((issue) => issue.issue_id);
  const snapshot = {
    project,
    members: readAllRows(SHEETS.members).filter((member) => member.project_id === projectId),
    phases,
    tasks: readAllRows(SHEETS.tasks).filter((task) => phaseIds.indexOf(task.phase_id) >= 0),
    issues,
    issueComments: readAllRows(SHEETS.issueComments).filter((comment) => issueIds.indexOf(comment.issue_id) >= 0),
    archived_at: new Date().toISOString(),
    archived_by: user.email
  };

  appendRecord(getSheet(SHEETS.archives), {
    archive_id: Utilities.getUuid(),
    project_id: projectId,
    project_name: project.project_name,
    archived_at: snapshot.archived_at,
    archived_by: user.email,
    snapshot: JSON.stringify(snapshot)
  });

  const sheet = getSheet(SHEETS.projects);
  const records = readAllRows(SHEETS.projects);
  const index = records.findIndex((record) => record.project_id === projectId);
  const updated = Object.assign({}, records[index], {
    deleted: 'ARCHIVED',
    updated_at: new Date().toISOString()
  });
  writeRecordAt(sheet, index + 2, updated);
  writeAuditLog(user.email, 'ARCHIVE_PROJECT', 'project', projectId, project, updated);

  return {
    deletedProjects: getDeletedProjectsForUser(user.email)
  };
}

function savePhase(payload) {
  const user = getCurrentUser();
  const phase = validatePhasePayload(payload);
  assertProjectAdmin(phase.project_id, user.email);

  const sheet = getSheet(SHEETS.phases);
  const records = readTable(SHEETS.phases);
  const existingIndex = records.findIndex((record) => record.phase_id === phase.phase_id);
  let savedPhase;
  let before = null;

  if (existingIndex >= 0) {
    before = records[existingIndex];
    savedPhase = Object.assign({}, before, phase, { deleted: before.deleted || '' });
    writeRecordAt(sheet, existingIndex + 2, savedPhase);
  } else {
    savedPhase = Object.assign({}, phase, {
      phase_id: Utilities.getUuid(),
      deleted: ''
    });
    appendRecord(sheet, savedPhase);
  }

  recalculatePhaseProgress(phase.project_id);
  writeAuditLog(user.email, existingIndex >= 0 ? 'UPDATE_PHASE' : 'CREATE_PHASE', 'phase', savedPhase.phase_id, before, savedPhase);
  return getProjectDetail(phase.project_id);
}

function deletePhase(phaseId) {
  const user = getCurrentUser();
  const sheet = getSheet(SHEETS.phases);
  const records = readAllRows(SHEETS.phases);
  const index = records.findIndex((record) => record.phase_id === phaseId && record.deleted !== 'TRUE');

  if (index < 0) {
    throw new Error('フェーズが見つかりません。');
  }

  const phase = records[index];
  assertProjectAdmin(phase.project_id, user.email);

  const activeTasks = readTable(SHEETS.tasks).filter((task) => task.phase_id === phaseId);
  if (activeTasks.length) {
    throw new Error('このフェーズには有効なタスクがあります。先にタスクを削除または移動してください。');
  }

  const updated = Object.assign({}, phase, { deleted: 'TRUE' });
  writeRecordAt(sheet, index + 2, updated);
  writeAuditLog(user.email, 'DELETE_PHASE', 'phase', phaseId, phase, updated);
  return getProjectDetail(phase.project_id);
}

function saveTask(payload) {
  const user = getCurrentUser();
  const task = validateTaskPayload(payload);
  const phase = findRecordById(SHEETS.phases, 'phase_id', task.phase_id);

  if (!phase) {
    throw new Error('フェーズが見つかりません。');
  }

  assertProjectMember(phase.project_id, user.email);

  const sheet = getSheet(SHEETS.tasks);
  const records = readTable(SHEETS.tasks);
  const existingIndex = records.findIndex((record) => record.task_id === task.task_id);
  const now = new Date().toISOString();
  let savedTask;
  let action;
  let before = null;

  if (existingIndex >= 0) {
    before = records[existingIndex];
    savedTask = Object.assign({}, before, task, {
      updated_at: now,
      deleted: before.deleted || ''
    });
    writeRecordAt(sheet, existingIndex + 2, savedTask);
    action = 'UPDATE_TASK';
  } else {
    savedTask = Object.assign({}, task, {
      task_id: Utilities.getUuid(),
      created_at: now,
      updated_at: now,
      deleted: ''
    });
    appendRecord(sheet, savedTask);
    action = 'CREATE_TASK';
  }

  recalculatePhaseProgress(phase.project_id);
  writeAuditLog(user.email, action, 'task', savedTask.task_id, before, savedTask);
  notifyTaskChange(phase.project_id, savedTask, action, user.email);
  return getProjectDetail(phase.project_id);
}

function deleteTask(taskId) {
  const user = getCurrentUser();
  const sheet = getSheet(SHEETS.tasks);
  const records = readAllRows(SHEETS.tasks);
  const index = records.findIndex((record) => record.task_id === taskId && record.deleted !== 'TRUE');

  if (index < 0) {
    throw new Error('タスクが見つかりません。');
  }

  const task = records[index];
  const phase = findRecordById(SHEETS.phases, 'phase_id', task.phase_id);
  assertProjectMember(phase.project_id, user.email);

  const updated = Object.assign({}, task, {
    deleted: 'TRUE',
    updated_at: new Date().toISOString()
  });

  writeRecordAt(sheet, index + 2, updated);
  recalculatePhaseProgress(phase.project_id);
  writeAuditLog(user.email, 'DELETE_TASK', 'task', taskId, task, updated);
  notifyTaskChange(phase.project_id, updated, 'DELETE_TASK', user.email);
  return getProjectDetail(phase.project_id);
}

function moveTask(taskId, targetPhaseId) {
  const user = getCurrentUser();
  const taskSheet = getSheet(SHEETS.tasks);
  const tasks = readAllRows(SHEETS.tasks);
  const taskIndex = tasks.findIndex((record) => record.task_id === taskId && record.deleted !== 'TRUE');

  if (taskIndex < 0) {
    throw new Error('タスクが見つかりません。');
  }

  const task = tasks[taskIndex];
  const sourcePhase = findRecordById(SHEETS.phases, 'phase_id', task.phase_id);
  const targetPhase = findRecordById(SHEETS.phases, 'phase_id', targetPhaseId);

  if (!sourcePhase || !targetPhase || sourcePhase.project_id !== targetPhase.project_id) {
    throw new Error('移動先フェーズが不正です。');
  }

  assertProjectMember(sourcePhase.project_id, user.email);

  const updated = Object.assign({}, task, {
    phase_id: targetPhaseId,
    updated_at: new Date().toISOString()
  });
  writeRecordAt(taskSheet, taskIndex + 2, updated);
  recalculatePhaseProgress(sourcePhase.project_id);
  writeAuditLog(user.email, 'MOVE_TASK', 'task', taskId, task, updated);
  return getProjectDetail(sourcePhase.project_id);
}

function moveTasks(taskIds, targetPhaseId) {
  const ids = Array.isArray(taskIds) ? taskIds : [];
  if (!ids.length) {
    throw new Error('移動するタスクを選択してください。');
  }

  let projectId = '';
  ids.forEach((taskId) => {
    const detail = moveTask(taskId, targetPhaseId);
    projectId = detail.project.project_id;
  });
  return getProjectDetail(projectId);
}

function updateSchedule(type, id, plannedStart, plannedEnd) {
  const user = getCurrentUser();
  const target = type === 'phase'
    ? { sheetName: SHEETS.phases, idColumn: 'phase_id' }
    : { sheetName: SHEETS.tasks, idColumn: 'task_id' };
  const sheet = getSheet(target.sheetName);
  const records = readAllRows(target.sheetName);
  const index = records.findIndex((record) => record[target.idColumn] === id && record.deleted !== 'TRUE');

  if (index < 0) {
    throw new Error('更新対象が見つかりません。');
  }

  const before = records[index];
  const projectId = type === 'phase'
    ? before.project_id
    : findRecordById(SHEETS.phases, 'phase_id', before.phase_id).project_id;
  assertProjectMember(projectId, user.email);

  if (plannedStart > plannedEnd) {
    throw new Error('開始日は終了日以前にしてください。');
  }

  const updated = Object.assign({}, before, {
    planned_start: plannedStart,
    planned_end: plannedEnd,
    updated_at: new Date().toISOString()
  });
  writeRecordAt(sheet, index + 2, updated);
  recalculatePhaseProgress(projectId);
  writeAuditLog(user.email, 'UPDATE_SCHEDULE', type, id, before, updated);
  return getProjectDetail(projectId);
}

function updateTaskProgress(taskId, progressPercent) {
  const user = getCurrentUser();
  const task = findRecordById(SHEETS.tasks, 'task_id', taskId);

  if (!task) {
    throw new Error('タスクが見つかりません。');
  }

  const phase = findRecordById(SHEETS.phases, 'phase_id', task.phase_id);
  assertProjectMember(phase.project_id, user.email);

  const progress = clampNumber(Number(progressPercent), 0, 100);
  const records = readTable(SHEETS.tasks);
  const index = records.findIndex((record) => record.task_id === taskId);
  const updated = Object.assign({}, records[index], {
    progress_percent: String(progress),
    status: progress >= 100 ? '完了' : progress > 0 ? '進行中' : '未着手',
    actual_end: progress >= 100 ? todayString() : records[index].actual_end,
    updated_at: new Date().toISOString()
  });

  writeRecordAt(getSheet(SHEETS.tasks), index + 2, updated);
  recalculatePhaseProgress(phase.project_id);
  writeAuditLog(user.email, 'UPDATE_PROGRESS', 'task', taskId, records[index], updated);
  notifyTaskChange(phase.project_id, updated, 'UPDATE_PROGRESS', user.email);
  return getProjectDetail(phase.project_id);
}

function saveMember(projectId, payload) {
  const user = getCurrentUser();
  assertProjectAdmin(projectId, user.email);

  const member = validateMemberPayload(projectId, payload);
  const sheet = getSheet(SHEETS.members);
  const records = readAllRows(SHEETS.members);
  const existingIndex = records.findIndex((record) =>
    record.project_id === projectId &&
    normalizeEmail(record.user_email) === normalizeEmail(member.user_email)
  );
  const before = existingIndex >= 0 ? records[existingIndex] : null;

  if (existingIndex >= 0) {
    writeRecordAt(sheet, existingIndex + 2, member);
  } else {
    appendRecord(sheet, member);
  }

  writeAuditLog(user.email, existingIndex >= 0 ? 'UPDATE_MEMBER' : 'CREATE_MEMBER', 'member', member.user_email, before, member);
  notifyMemberChange(projectId, member, user.email);
  return getProjectDetail(projectId);
}

function deactivateMember(projectId, email) {
  const user = getCurrentUser();
  assertProjectAdmin(projectId, user.email);

  if (normalizeEmail(email) === normalizeEmail(user.email)) {
    throw new Error('自分自身は無効化できません。');
  }

  const sheet = getSheet(SHEETS.members);
  const records = readAllRows(SHEETS.members);
  const index = records.findIndex((record) =>
    record.project_id === projectId &&
    normalizeEmail(record.user_email) === normalizeEmail(email)
  );

  if (index < 0) {
    throw new Error('メンバーが見つかりません。');
  }

  const updated = Object.assign({}, records[index], { active: 'FALSE' });
  writeRecordAt(sheet, index + 2, updated);
  writeAuditLog(user.email, 'DEACTIVATE_MEMBER', 'member', email, records[index], updated);
  return getProjectDetail(projectId);
}

function saveIssue(payload) {
  const user = getCurrentUser();
  const issue = validateIssuePayload(payload);
  assertProjectMember(issue.project_id, user.email);

  const sheet = getSheet(SHEETS.issues);
  const records = readTable(SHEETS.issues);
  const existingIndex = records.findIndex((record) => record.issue_id === issue.issue_id);
  const now = new Date().toISOString();
  let savedIssue;
  let before = null;

  if (existingIndex >= 0) {
    before = records[existingIndex];
    savedIssue = Object.assign({}, before, issue, { updated_at: now, deleted: before.deleted || '' });
    writeRecordAt(sheet, existingIndex + 2, savedIssue);
  } else {
    savedIssue = Object.assign({}, issue, {
      issue_id: Utilities.getUuid(),
      created_at: now,
      updated_at: now,
      deleted: ''
    });
    appendRecord(sheet, savedIssue);
  }

  writeAuditLog(user.email, existingIndex >= 0 ? 'UPDATE_ISSUE' : 'CREATE_ISSUE', 'issue', savedIssue.issue_id, before, savedIssue);
  notifyIssueChange(savedIssue, user.email);
  return getProjectDetail(issue.project_id);
}

function deleteIssue(issueId) {
  const user = getCurrentUser();
  const sheet = getSheet(SHEETS.issues);
  const records = readAllRows(SHEETS.issues);
  const index = records.findIndex((record) => record.issue_id === issueId && record.deleted !== 'TRUE');

  if (index < 0) {
    throw new Error('課題が見つかりません。');
  }

  const issue = records[index];
  assertProjectMember(issue.project_id, user.email);

  const updated = Object.assign({}, issue, {
    deleted: 'TRUE',
    updated_at: new Date().toISOString()
  });

  writeRecordAt(sheet, index + 2, updated);
  writeAuditLog(user.email, 'DELETE_ISSUE', 'issue', issueId, issue, updated);
  notifyIssueChange(updated, user.email);
  return getProjectDetail(issue.project_id);
}

function saveIssueComment(payload) {
  const user = getCurrentUser();
  const comment = validateIssueCommentPayload(payload);
  const issue = findRecordById(SHEETS.issues, 'issue_id', comment.issue_id);

  if (!issue) {
    throw new Error('課題が見つかりません。');
  }

  assertProjectMember(issue.project_id, user.email);
  ensureSheet(SHEETS.issueComments);

  const saved = Object.assign({}, comment, {
    comment_id: Utilities.getUuid(),
    user_email: user.email,
    created_at: new Date().toISOString(),
    updated_at: '',
    deleted: ''
  });
  appendRecord(getSheet(SHEETS.issueComments), saved);
  writeAuditLog(user.email, 'CREATE_ISSUE_COMMENT', 'issue', comment.issue_id, null, saved);

  const mentions = extractMentions(saved.comment);
  if (mentions.length) notifyMentionInComment(issue, saved.comment, mentions, user.email);

  return getProjectDetail(issue.project_id);
}

function updateIssueComment(payload) {
  const user = getCurrentUser();
  const comment = validateIssueCommentPayload(payload);
  const sheet = getSheet(SHEETS.issueComments);
  const records = readAllRows(SHEETS.issueComments);
  const index = records.findIndex((record) => record.comment_id === String(payload.comment_id || '') && record.deleted !== 'TRUE');

  if (index < 0) {
    throw new Error('コメントが見つかりません。');
  }

  const before = records[index];
  const issue = findRecordById(SHEETS.issues, 'issue_id', before.issue_id);
  assertProjectMember(issue.project_id, user.email);

  if (normalizeEmail(before.user_email) !== normalizeEmail(user.email) && !isProjectAdmin(issue.project_id, user.email)) {
    throw new Error('コメントの編集権限がありません。');
  }

  const updated = Object.assign({}, before, {
    comment: comment.comment,
    updated_at: new Date().toISOString()
  });
  writeRecordAt(sheet, index + 2, updated);
  writeAuditLog(user.email, 'UPDATE_ISSUE_COMMENT', 'issue', before.issue_id, before, updated);

  const mentions = extractMentions(updated.comment);
  if (mentions.length) notifyMentionInComment(issue, updated.comment, mentions, user.email);

  return getProjectDetail(issue.project_id);
}

function deleteIssueComment(commentId) {
  const user = getCurrentUser();
  const sheet = getSheet(SHEETS.issueComments);
  const records = readAllRows(SHEETS.issueComments);
  const index = records.findIndex((record) => record.comment_id === commentId && record.deleted !== 'TRUE');

  if (index < 0) {
    throw new Error('コメントが見つかりません。');
  }

  const before = records[index];
  const issue = findRecordById(SHEETS.issues, 'issue_id', before.issue_id);
  assertProjectMember(issue.project_id, user.email);

  if (normalizeEmail(before.user_email) !== normalizeEmail(user.email) && !isProjectAdmin(issue.project_id, user.email)) {
    throw new Error('コメントの削除権限がありません。');
  }

  const updated = Object.assign({}, before, {
    deleted: 'TRUE',
    updated_at: new Date().toISOString()
  });
  writeRecordAt(sheet, index + 2, updated);
  writeAuditLog(user.email, 'DELETE_ISSUE_COMMENT', 'issue', before.issue_id, before, updated);
  return getProjectDetail(issue.project_id);
}

function getPortfolioReport() {
  const user = getCurrentUser();
  assertWorkspaceUser(user.email);

  return getVisibleProjectsForUser(user.email).map((project) => {
    const detail = getProjectDetail(project.project_id);
    return {
      project_id: project.project_id,
      project_name: project.project_name,
      client_name: project.client_name,
      status: project.status,
      owner_email: project.owner_email,
      start_date: project.start_date,
      end_date: project.end_date,
      progress: detail.summary.progress,
      taskCount: detail.summary.taskCount,
      delayedTaskCount: detail.summary.delayedTaskCount,
      openIssueCount: detail.summary.openIssueCount
    };
  });
}

function getDeletedRecords(projectId) {
  const user = getCurrentUser();
  assertProjectAdmin(projectId, user.email);

  const allPhases = readAllRows(SHEETS.phases).filter((phase) => phase.project_id === projectId);
  const phaseIds = allPhases.map((phase) => phase.phase_id);
  return {
    projects: readAllRows(SHEETS.projects).filter((project) => project.project_id === projectId && project.deleted === 'TRUE'),
    phases: allPhases.filter((phase) => phase.deleted === 'TRUE'),
    tasks: readAllRows(SHEETS.tasks).filter((task) => phaseIds.indexOf(task.phase_id) >= 0 && task.deleted === 'TRUE'),
    issues: readAllRows(SHEETS.issues).filter((issue) => issue.project_id === projectId && issue.deleted === 'TRUE')
  };
}

function restoreRecord(projectId, type, id) {
  const user = getCurrentUser();
  assertProjectAdmin(projectId, user.email);

  const map = {
    phase: { sheetName: SHEETS.phases, idColumn: 'phase_id' },
    task: { sheetName: SHEETS.tasks, idColumn: 'task_id' },
    issue: { sheetName: SHEETS.issues, idColumn: 'issue_id' }
  };
  const target = map[type];
  if (!target) {
    throw new Error('復元対象が不正です。');
  }

  const sheet = getSheet(target.sheetName);
  const records = readAllRows(target.sheetName);
  const index = records.findIndex((record) => record[target.idColumn] === id && record.deleted === 'TRUE');

  if (index < 0) {
    throw new Error('復元対象が見つかりません。');
  }

  const before = records[index];
  if (type === 'task') {
    const phase = findRecordById(SHEETS.phases, 'phase_id', before.phase_id);
    if (!phase) {
      throw new Error('移動先フェーズが削除済みです。先にフェーズを復元してください。');
    }
  }

  const restored = Object.assign({}, before, {
    deleted: '',
    updated_at: new Date().toISOString()
  });
  writeRecordAt(sheet, index + 2, restored);
  recalculatePhaseProgress(projectId);
  writeAuditLog(user.email, 'RESTORE_' + type.toUpperCase(), type, id, before, restored);
  return {
    activeProject: getProjectDetail(projectId),
    deletedRecords: getDeletedRecords(projectId)
  };
}

function getAuditLogs(projectId) {
  const user = getCurrentUser();
  assertProjectAdmin(projectId, user.email);
  return getProjectAuditLogs(projectId);
}

function exportAuditLogsCsv(projectId) {
  const user = getCurrentUser();
  assertProjectAdmin(projectId, user.email);
  const logs = getProjectAuditLogs(projectId);
  writeAuditLog(user.email, 'EXPORT_AUDIT_CSV', 'project', projectId, null, { count: logs.length });

  const rows = [['timestamp', 'user_email', 'action', 'target_type', 'target_id', 'before', 'after']]
    .concat(logs.map((log) => [log.timestamp, log.user_email, log.action, log.target_type, log.target_id, log.before, log.after]));
  return rows.map((row) => row.map(csvEscape).join(',')).join('\n');
}

function getProjectAuditLogs(projectId) {
  const detail = getProjectDetail(projectId);
  const ids = {};
  ids[projectId] = true;
  detail.phases.forEach((phase) => ids[phase.phase_id] = true);
  detail.tasks.forEach((task) => ids[task.task_id] = true);
  detail.issues.forEach((issue) => ids[issue.issue_id] = true);
  detail.members.forEach((member) => ids[member.user_email] = true);

  return readTable(SHEETS.auditLogs)
    .filter((log) => ids[log.target_id] || log.target_type === 'notification')
    .slice(-500)
    .reverse();
}

function saveNotificationSettings(payload) {
  const user = getCurrentUser();
  const projectId = String(payload.project_id || '');
  assertProjectAdmin(projectId, user.email);

  const settings = validateNotificationSettings(payload);
  Object.keys(settings).forEach((key) => setSetting(key, settings[key]));
  writeAuditLog(user.email, 'UPDATE_NOTIFICATION_SETTINGS', 'project', projectId, null, settings);
  return getNotificationSettings();
}

function sendNotificationPreview(projectId, payload, previewType) {
  const user = getCurrentUser();
  assertProjectAdmin(projectId, user.email);

  const settings = validateNotificationSettings(payload);
  const context = {
    action: 'PREVIEW',
    task_name: 'サンプルタスク',
    assignee_email: user.email,
    progress_percent: '50',
    status: '進行中',
    actor_email: user.email,
    issue_title: 'サンプル課題',
    severity: 'Medium',
    due_date: todayString(),
    owner_email: user.email,
    project_name: getProjectDetail(projectId).project.project_name,
    overdue_tasks: '- サンプルタスク / ' + user.email + ' / ' + todayString(),
    overdue_issues: '- サンプル課題 / ' + user.email + ' / ' + todayString()
  };
  const type = previewType || 'task';
  const subjectTemplate = type === 'issue'
    ? settings.notify_issue_subject_template
    : type === 'overdue'
      ? settings.notify_overdue_subject_template
      : settings.notify_task_subject_template;
  const bodyTemplate = type === 'issue'
    ? settings.notify_issue_body_template
    : type === 'overdue'
      ? settings.notify_overdue_body_template
      : settings.notify_task_body_template;
  const subject = applyTemplate(subjectTemplate || '[Progress Board] プレビュー: {{project_name}}', context);
  const body = applyTemplate(bodyTemplate || '通知プレビューです。\n{{project_name}}', context);

  sendMailSafely([user.email], subject, body);
  writeAuditLog(user.email, 'SEND_NOTIFICATION_PREVIEW', 'project', projectId, null, { to: user.email, type });
  return true;
}

function installOverdueNotificationTrigger(projectId) {
  const user = getCurrentUser();
  assertProjectAdmin(projectId, user.email);

  ScriptApp.getProjectTriggers()
    .filter((trigger) => trigger.getHandlerFunction() === 'sendOverdueNotifications')
    .forEach((trigger) => ScriptApp.deleteTrigger(trigger));

  ScriptApp.newTrigger('sendOverdueNotifications')
    .timeBased()
    .everyDays(1)
    .atHour(Number(getSetting('notify_overdue_hour') || 9))
    .create();

  setSetting('notify_overdue_project_id', projectId);
  writeAuditLog(user.email, 'INSTALL_OVERDUE_TRIGGER', 'project', projectId, null, { hour: getSetting('notify_overdue_hour') || '9' });
  return getNotificationSettings();
}

function sendOverdueNotifications() {
  const projectId = getSetting('notify_overdue_project_id');
  if (!projectId || getSetting('notify_overdue_daily') === 'FALSE') {
    return;
  }
  notifyOverdueItems(projectId);
}

function exportProjectCsv(projectId) {
  const user = getCurrentUser();
  assertProjectMember(projectId, user.email);
  const detail = getProjectDetail(projectId);
  writeAuditLog(user.email, 'EXPORT_CSV', 'project', projectId, null, { project_id: projectId });
  return buildProjectCsv(detail);
}

function setupSampleSpreadsheet() {
  const user = getCurrentUser();
  const email = user.email || 'owner@example.com';
  initializeSheets();
  seedSampleData(email);
  return getAppBootstrap();
}
