function writeAuditLog(userEmail, action, targetType, targetId, beforeValue, afterValue) {
  appendRecord(getSheet(SHEETS.auditLogs), {
    timestamp: new Date().toISOString(),
    user_email: userEmail,
    action,
    target_type: targetType,
    target_id: targetId,
    before: beforeValue ? JSON.stringify(beforeValue) : '',
    after: afterValue ? JSON.stringify(afterValue) : ''
  });
}
