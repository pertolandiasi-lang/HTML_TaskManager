const SHEET_ID     = '1aUW4iKrUhUQBNISsY2xISgNHlfOc7e6dDjuqVzHgC68';
const DRIVE_FOLDER = '1qZdcgxj1neycYTj4BmTQhI1SbGklgowc';
const BACKUP_FOLDER = '1ct1aOAP0Qapa3AlcGKJmWxqp5w7ilPCp'; // Drive folder for daily Sheet backups
const BACKUP_RETENTION_DAYS = 30;

// ── ENTRY POINTS ─────────────────────────────────────────────────────────────

function doPost(e) {
  try {
    var p = JSON.parse(e.postData.contents);
    return dispatch_(p);
  } catch(err) { return respond_({ error: err.message }); }
}

function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  var result;
  try { result = dispatch_(params); }
  catch(err) { result = respond_({ error: err.message }); }
  var cb = params.callback;
  if (cb) {
    return ContentService
      .createTextOutput(cb + '(' + result.getContent() + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return result;
}

function authorize() {
  UrlFetchApp.fetch('https://www.googleapis.com/oauth2/v3/tokeninfo?access_token=test', { muteHttpExceptions: true });
  SpreadsheetApp.openById(SHEET_ID);
  var tmp = DocumentApp.create('_auth_test_tmp');
  DriveApp.getFileById(tmp.getId()).setTrashed(true);
  Logger.log('Autorizzazione completata');
}

// ── ROUTER ────────────────────────────────────────────────────────────────────

function dispatch_(p) {
  // Called by the Sheet-bound script after a multi-assign — re-syncs one row.
  if (p.action === 'syncRow' && p.key === 'tf_syncrow_v56') {
    syncTaskRow_(parseInt(p.row, 10));
    return respond_({ ok: true });
  }
  if (!p.token) throw new Error('Token mancante');
  var callerEmail = verifyToken_(p.token);

  switch (p.action) {
    case 'getRole':
      return respond_({ role: getUserRole_(callerEmail), email: callerEmail, name: getUserName_(callerEmail) });

    case 'getTeam':
      requireRole_(callerEmail, 'manager');
      return respond_({ team: getTeam_() });

    case 'getTasks':
      return respond_({ tasks: getTasksForUser_(callerEmail) });

    case 'getAllTasks':
      requireRole_(callerEmail, 'manager');
      return respond_({ tasks: getAllTasks_() });

    case 'updateStatus':
      updateStatus_(p.taskId, p.status, callerEmail);
      return respond_({ ok: true });

    case 'deleteTask':
      requireRole_(callerEmail, 'manager');
      deleteTask_(p.taskId);
      return respond_({ ok: true });

    case 'setupValidation':
      requireRole_(callerEmail, 'manager');
      setupSheetValidation();
      return respond_({ ok: true });

    default:
      throw new Error('Azione non riconosciuta');
  }
}

function respond_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── IDENTITY / AUTH ───────────────────────────────────────────────────────────

function verifyToken_(token) {
  var resp = UrlFetchApp.fetch(
    'https://www.googleapis.com/oauth2/v3/tokeninfo?access_token=' + encodeURIComponent(token),
    { muteHttpExceptions: true }
  );
  if (resp.getResponseCode() !== 200) throw new Error('Token non valido');
  var info = JSON.parse(resp.getContentText());
  if (info.error_description || !info.email) throw new Error('Token non valido');
  return info.email;
}

function getUserRole_(email) {
  try {
    var rows = getSheet_('Team').getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][1] && rows[i][1].toString().trim().toLowerCase() === email.toLowerCase())
        return rows[i][2].toString().trim().toLowerCase();
    }
  } catch(e) {}
  return 'unauthorized';
}

function getUserName_(email) {
  try {
    var rows = getSheet_('Team').getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][1] && rows[i][1].toString().trim().toLowerCase() === email.toLowerCase())
        return rows[i][0].toString().trim();
    }
  } catch(e) {}
  return email;
}

function requireRole_(email, role) {
  if (getUserRole_(email) !== role) throw new Error('Non autorizzato');
}

// ── DATA HELPERS ──────────────────────────────────────────────────────────────

function getSheet_(name) {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
}

function formatDate_(val) {
  if (!val) return '';
  if (val instanceof Date) return Utilities.formatDate(val, 'UTC', 'yyyy-MM-dd');
  return val.toString().trim();
}

// Per-assignee status is stored as JSON in col I, e.g. {"a@x":"In lavoro","b@x":"Da fare"}.
// Legacy single-string values ("Da fare"/"In lavoro"/"Completato") are migrated lazily
// to "every assignee has this status" on read.
function parseStatusMap_(raw, assigneeCsv) {
  raw = String(raw || '').trim();
  var map = {};
  if (raw && raw.charAt(0) === '{') {
    try { var m = JSON.parse(raw); if (m && typeof m === 'object') map = m; } catch(e) {}
  } else {
    var legacy = (raw === 'In lavoro' || raw === 'Completato') ? raw : 'Da fare';
    String(assigneeCsv || '').split(',').forEach(function(em) {
      em = em.trim().toLowerCase();
      if (em) map[em] = legacy;
    });
  }
  var norm = {};
  Object.keys(map).forEach(function(k){ norm[k.toLowerCase()] = map[k]; });
  return norm;
}

// User rule: 'Da fare' only if ALL Da fare; 'Completato' only if ALL Completato;
// otherwise 'In lavoro' (including the mixed Completato+Da fare case).
function aggregateStatus_(map, assignees) {
  if (!assignees || !assignees.length) return 'Da fare';
  var allDone = true, allTodo = true;
  for (var i = 0; i < assignees.length; i++) {
    var s = map[assignees[i]] || 'Da fare';
    if (s !== 'Completato') allDone = false;
    if (s !== 'Da fare')    allTodo = false;
  }
  if (allDone) return 'Completato';
  if (allTodo) return 'Da fare';
  return 'In lavoro';
}

function assigneeList_(csv) {
  return String(csv||'').split(',').map(function(s){return s.trim().toLowerCase();}).filter(Boolean);
}

function rowToTask_(r) {
  var statusMap = parseStatusMap_(r[8], r[2]);
  var assignees = assigneeList_(r[2]);
  return {
    id: String(r[0]), company: String(r[1]||''), assignee: String(r[2]||''),
    assignDate: formatDate_(r[3]), deadline: formatDate_(r[4]),
    brief: String(r[5]||''), driveUrl: String(r[6]||''), docUrl: String(r[7]||''),
    statusMap: statusMap,
    status: aggregateStatus_(statusMap, assignees)
  };
}

function getTeam_() {
  var rows = getSheet_('Team').getDataRange().getValues(), team = [];
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][1]) team.push({
      name:  String(rows[i][0]||''),
      email: rows[i][1].toString().trim(),
      role:  (rows[i][2]||'').toString().trim().toLowerCase()
    });
  }
  return team;
}

function emailInCsv_(csv, email) {
  if (!csv) return false;
  var target = email.toLowerCase();
  var list = csv.toString().toLowerCase().split(',');
  for (var i = 0; i < list.length; i++) {
    if (list[i].trim() === target) return true;
  }
  return false;
}

function namesForCsv_(csv) {
  if (!csv) return '';
  var rows = getSheet_('Team').getDataRange().getValues(), map = {};
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][1]) map[rows[i][1].toString().trim().toLowerCase()] = rows[i][0].toString().trim();
  }
  return csv.toString().split(',').map(function(e) {
    var k = e.trim().toLowerCase();
    return map[k] || e.trim();
  }).join(', ');
}

function getTasksForUser_(email) {
  var rows = getSheet_('Tasks').getDataRange().getValues(), tasks = [];
  var key = String(email||'').toLowerCase();
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][2]) continue;
    if (!emailInCsv_(rows[i][2], email)) continue;
    var t = rowToTask_(rows[i]);
    // Employee view: surface THEIR own status as `status`. Manager still gets aggregate via getAllTasks_.
    t.status = t.statusMap[key] || 'Da fare';
    tasks.push(t);
  }
  return tasks.sort(function(a,b){ return a.deadline < b.deadline ? -1 : 1; });
}

function getAllTasks_() {
  var rows = getSheet_('Tasks').getDataRange().getValues(), tasks = [];
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    tasks.push(rowToTask_(rows[i]));
  }
  return tasks.sort(function(a,b){ return a.deadline < b.deadline ? -1 : 1; });
}

// ── WRITE OPERATIONS ──────────────────────────────────────────────────────────

function updateStatus_(taskId, newStatus, callerEmail) {
  var valid = ['Da fare', 'In lavoro', 'Completato'];
  if (valid.indexOf(newStatus) === -1) throw new Error('Stato non valido');
  var sheet = getSheet_('Tasks'), rows = sheet.getDataRange().getValues();
  var key = String(callerEmail||'').toLowerCase();
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0] || rows[i][0].toString() !== taskId.toString()) continue;
    // Per-assignee status: only assignees can change their own slot. A manager who
    // isn't assigned to this task has nothing to update here (the aggregate is derived).
    if (!emailInCsv_(rows[i][2], callerEmail))
      throw new Error('Non sei assegnato a questo task');
    var map = parseStatusMap_(rows[i][8], rows[i][2]);
    map[key] = newStatus;
    sheet.getRange(i + 1, 9).setValue(JSON.stringify(map));
    return;
  }
  throw new Error('Task non trovato');
}

function deleteTask_(taskId) {
  var sheet = getSheet_('Tasks'), rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0] || rows[i][0].toString() !== taskId.toString()) continue;
    sheet.deleteRow(i + 1);
    return;
  }
  throw new Error('Task non trovato');
}

// ── GOOGLE DOC ────────────────────────────────────────────────────────────────

function updateDocBody_(docUrl, company, assignee, assignDate, deadline, brief, driveUrl) {
  var docId = docUrl.replace('https://docs.google.com/document/d/', '').split('/')[0].split('?')[0];
  var doc = DocumentApp.openById(docId);
  doc.setName('Brief - ' + company + ' - ' + assignDate + ' → ' + deadline);
  var body = doc.getBody();
  body.clear();
  body.appendParagraph(company).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('Assegnatari: ' + namesForCsv_(assignee));
  body.appendParagraph('Data assegnazione: ' + assignDate);
  body.appendParagraph('Deadline: ' + deadline);
  body.appendParagraph('');
  body.appendParagraph('Brief:').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(brief || '—');
  doc.saveAndClose();
}

// ── SHEET VALIDATION SETUP (run once) ────────────────────────────────────────

function setupSheetValidation() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var taskSheet = ss.getSheetByName('Tasks');
  var extraRows = 500; // cover future rows

  // Status (col I) is now per-assignee JSON — drop the legacy enum validator
  // so the sheet doesn't reject the new format.
  taskSheet.getRange(2, 9, extraRows, 1).clearDataValidations();

  // NOTE: the assignee dropdown (col C) is intentionally NOT set here.
  // It must be created via the Sheets UI so the native
  // "Allow multiple selections" toggle stays available.

  // Date picker for assignDate (col D=4) and deadline (col E=5)
  var dateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .build();
  taskSheet.getRange(2, 4, extraRows, 1).setDataValidation(dateRule);
  taskSheet.getRange(2, 5, extraRows, 1).setDataValidation(dateRule);
  taskSheet.getRange(2, 4, extraRows, 1).setNumberFormat('yyyy-mm-dd');
  taskSheet.getRange(2, 5, extraRows, 1).setNumberFormat('yyyy-mm-dd');

  Logger.log('Sheet validation setup completato!');
}

// ── INSTALLABLE TRIGGER: Sheet → Doc ──────────────────────────────────────────

// Run createInstallableTrigger() ONCE from the Apps Script editor to activate.
function createInstallableTrigger() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  ScriptApp.getProjectTriggers().forEach(function(t) {
    var fn = t.getHandlerFunction();
    if (fn === 'onTaskSheetEdit' || fn === 'onTaskSheetOpen') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onTaskSheetEdit').forSpreadsheet(ss).onEdit().create();
  Logger.log('Trigger installato (onEdit)!');
}

// Deprecated no-op — kept so any stale onOpen trigger doesn't error.
function onTaskSheetOpen() {}

// ── ROW SYNC (called by the Sheet-bound script after a multi-assign) ──────────

function syncTaskRow_(row) {
  if (!row || row < 2) return;
  var sheet = getSheet_('Tasks');
  var r = sheet.getRange(row, 1, 1, 11).getValues()[0];
  if (!r[0]) return;
  var company = String(r[1]||''), assignee = String(r[2]||'');
  var assignDate = formatDate_(r[3]), deadline = formatDate_(r[4]);
  var brief = String(r[5]||''), driveUrl = String(r[6]||''), docUrl = String(r[7]||'');
  if (docUrl) {
    try { updateDocBody_(docUrl, company, assignee, assignDate, deadline, brief, driveUrl); }
    catch(e) { Logger.log('syncTaskRow_ doc err: ' + e.message); }
  }
}

function onTaskSheetEdit(e) {
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet();
    if (sheet.getName() !== 'Tasks') return;
    var row = e.range.getRow();
    if (row < 2) return;

    var col = e.range.getColumn();
    // Only content columns require Doc sync (1-indexed):
    // 2=company 3=assignee 4=assignDate 5=deadline 6=brief 7=driveUrl
    if ([2,3,4,5,6,7].indexOf(col) === -1) return;

    var rowData = sheet.getRange(row, 1, 1, 11).getValues()[0];
    var taskId          = String(rowData[0]||'');
    var company         = String(rowData[1]||'');
    var assignee        = String(rowData[2]||'');
    var assignDate      = formatDate_(rowData[3]);
    var deadline        = formatDate_(rowData[4]);
    var brief           = String(rowData[5]||'');
    var driveUrl        = String(rowData[6]||'');
    var docUrl          = String(rowData[7]||'');

    if (!taskId || !docUrl) return;
    updateDocBody_(docUrl, company, assignee, assignDate, deadline, brief, driveUrl);
  } catch(err) {
    Logger.log('onTaskSheetEdit error: ' + err.message);
  }
}

// ── DAILY BACKUP ──────────────────────────────────────────────────────────────
// Run installBackupTrigger() ONCE from the Apps Script editor to schedule
// daily Sheet snapshots into BACKUP_FOLDER. Old snapshots beyond
// BACKUP_RETENTION_DAYS are trashed automatically.

function installBackupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'dailyBackup') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyBackup').timeBased().atHour(3).everyDays(1).create();
  Logger.log('Backup giornaliero installato (03:00).');
}

function dailyBackup() {
  var src = DriveApp.getFileById(SHEET_ID);
  var folder = DriveApp.getFolderById(BACKUP_FOLDER);
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'UTC', 'yyyy-MM-dd_HH-mm');
  var name = 'Tasks_backup_' + stamp;
  src.makeCopy(name, folder);

  // Trash backups older than BACKUP_RETENTION_DAYS
  var cutoff = Date.now() - BACKUP_RETENTION_DAYS * 86400000;
  var iter = folder.getFiles();
  while (iter.hasNext()) {
    var f = iter.next();
    if (f.getName().indexOf('Tasks_backup_') !== 0) continue;
    if (f.getDateCreated().getTime() < cutoff) f.setTrashed(true);
  }
}
