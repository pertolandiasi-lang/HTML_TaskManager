const SHEET_ID     = '1aUW4iKrUhUQBNISsY2xISgNHlfOc7e6dDjuqVzHgC68';
const CALENDAR_ID  = 'primary';
const DRIVE_FOLDER = '1qZdcgxj1neycYTj4BmTQhI1SbGklgowc';

// Calendar colorIds: 5=Banana(yellow) 8=Graphite(grey) 10=Basil(green) 11=Tomato(red)
// Deadline always red unless Completato (then green).
// Creation event: red if Da fare, yellow if In lavoro, grey if Completato.
const COLOR_MAP = {
  'Da fare':    { assign: '11', deadline: '11' },
  'In lavoro':  { assign: '5',  deadline: '11' },
  'Completato': { assign: '8',  deadline: '10' },
};

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
  UrlFetchApp.fetch('https://www.googleapis.com/oauth2/v3/tokeninfo?access_token=test');
  SpreadsheetApp.openById(SHEET_ID);
  CalendarApp.getDefaultCalendar();
  var tmp = DocumentApp.create('_auth_test_tmp');
  DriveApp.getFileById(tmp.getId()).setTrashed(true);
  Logger.log('Autorizzazione completata');
}

// ── ROUTER ────────────────────────────────────────────────────────────────────

function dispatch_(p) {
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
      setupSheetValidation_();
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

function addOneDay_(dateStr) {
  var d = new Date(dateStr + 'T12:00:00Z');
  d.setUTCDate(d.getUTCDate() + 1);
  return Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');
}

function rowToTask_(r) {
  return {
    id: String(r[0]), company: String(r[1]||''), assignee: String(r[2]||''),
    assignDate: formatDate_(r[3]), deadline: formatDate_(r[4]),
    brief: String(r[5]||''), driveUrl: String(r[6]||''), docUrl: String(r[7]||''),
    status: String(r[8]||'Da fare'),
    assignEventId: String(r[9]||''), deadlineEventId: String(r[10]||'')
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

function getTasksForUser_(email) {
  var rows = getSheet_('Tasks').getDataRange().getValues(), tasks = [];
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][2]) continue;
    if (rows[i][2].toString().trim().toLowerCase() !== email.toLowerCase()) continue;
    tasks.push(rowToTask_(rows[i]));
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
  var role = getUserRole_(callerEmail);
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0] || rows[i][0].toString() !== taskId.toString()) continue;
    if (role !== 'manager' && rows[i][2].toString().trim().toLowerCase() !== callerEmail.toLowerCase())
      throw new Error('Non autorizzato');
    sheet.getRange(i + 1, 9).setValue(newStatus);
    var colors = COLOR_MAP[newStatus];
    if (colors) {
      var assignId = String(rows[i][9]||''), deadlineId = String(rows[i][10]||'');
      if (assignId)   tryPatchColor_(assignId,   colors.assign);
      if (deadlineId) tryPatchColor_(deadlineId, colors.deadline);
    }
    return;
  }
  throw new Error('Task non trovato');
}

function deleteTask_(taskId) {
  var sheet = getSheet_('Tasks'), rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0] || rows[i][0].toString() !== taskId.toString()) continue;
    var assignId = String(rows[i][9]||''), deadlineId = String(rows[i][10]||'');
    if (assignId)   tryDeleteEvent_(assignId);
    if (deadlineId) tryDeleteEvent_(deadlineId);
    sheet.deleteRow(i + 1);
    return;
  }
  throw new Error('Task non trovato');
}

// ── CALENDAR ──────────────────────────────────────────────────────────────────

function recolorAllEvents_() {
  var rows = getSheet_('Tasks').getDataRange().getValues();
  var n = 0;
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    var status = String(rows[i][8] || 'Da fare');
    var colors = COLOR_MAP[status]; if (!colors) continue;
    var aId = String(rows[i][9] || ''), dId = String(rows[i][10] || '');
    if (aId) tryPatchColor_(aId, colors.assign);
    if (dId) tryPatchColor_(dId, colors.deadline);
    n++;
  }
  return n;
}

function tryPatchColor_(eventId, colorId) {
  try { Calendar.Events.patch({ colorId: colorId }, CALENDAR_ID, eventId); }
  catch(e) { Logger.log('Patch color err: ' + e.message); }
}

function tryDeleteEvent_(eventId) {
  try { Calendar.Events.remove(CALENDAR_ID, eventId); }
  catch(e) { Logger.log('Delete event err: ' + e.message); }
}

function syncEventsForRow_(company, assignee, assignDate, deadline, brief, driveUrl, docUrl, status, assignEventId, deadlineEventId) {
  var colors = COLOR_MAP[status] || COLOR_MAP['Da fare'];
  var desc = 'Assegnatario: '+assignee+'\n\nBrief:\n'+brief+(driveUrl?'\n\nDrive:\n'+driveUrl:'')+(docUrl?'\n\nDoc:\n'+docUrl:'');
  if (assignEventId) {
    try {
      Calendar.Events.patch({
        summary: company,
        start: { date: assignDate },
        end: { date: addOneDay_(assignDate) },
        description: desc,
        colorId: colors.assign
      }, CALENDAR_ID, assignEventId);
    } catch(e) { Logger.log('Sync assign event err: '+e.message); }
  }
  if (deadlineEventId) {
    try {
      Calendar.Events.patch({
        summary: '⚠️ DEADLINE - '+company,
        start: { date: deadline },
        end: { date: addOneDay_(deadline) },
        description: desc,
        colorId: colors.deadline
      }, CALENDAR_ID, deadlineEventId);
    } catch(e) { Logger.log('Sync deadline event err: '+e.message); }
  }
}

// ── GOOGLE DOC ────────────────────────────────────────────────────────────────

function updateDocBody_(docUrl, company, assignee, assignDate, deadline, brief, driveUrl) {
  var docId = docUrl.replace('https://docs.google.com/document/d/', '').split('/')[0].split('?')[0];
  var doc = DocumentApp.openById(docId);
  var body = doc.getBody();
  body.clear();
  body.appendParagraph(company).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('Assegnatario: ' + assignee);
  body.appendParagraph('Data assegnazione: ' + assignDate);
  body.appendParagraph('Deadline: ' + deadline);
  body.appendParagraph('');
  body.appendParagraph('Brief:').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(brief || '—');
  if (driveUrl) {
    body.appendParagraph('');
    body.appendParagraph('Drive: ' + driveUrl);
  }
  doc.saveAndClose();
}

// ── SHEET VALIDATION SETUP (run once) ────────────────────────────────────────

function setupSheetValidation_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var taskSheet = ss.getSheetByName('Tasks');
  var teamSheet = ss.getSheetByName('Team');
  var extraRows = 500; // cover future rows

  // Status dropdown — col I (9)
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Da fare', 'In lavoro', 'Completato'], true)
    .setAllowInvalid(false)
    .build();
  taskSheet.getRange(2, 9, extraRows, 1).setDataValidation(statusRule);

  // Assignee dropdown from Team!B column — col C (3)
  var teamLast = Math.max(teamSheet.getLastRow(), 2);
  var emailRange = teamSheet.getRange('B2:B' + teamLast);
  var emailRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(emailRange, true)
    .setAllowInvalid(false)
    .build();
  taskSheet.getRange(2, 3, extraRows, 1).setDataValidation(emailRule);

  // Date format for assignDate (col D=4) and deadline (col E=5)
  taskSheet.getRange(2, 4, extraRows, 1).setNumberFormat('yyyy-mm-dd');
  taskSheet.getRange(2, 5, extraRows, 1).setNumberFormat('yyyy-mm-dd');

  Logger.log('Sheet validation setup completato!');
}

// ── INSTALLABLE TRIGGER: Sheet → Calendar + Doc ───────────────────────────────

// Run createInstallableTrigger() ONCE from the Apps Script editor to activate.
function createInstallableTrigger() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'onTaskSheetEdit') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onTaskSheetEdit').forSpreadsheet(ss).onEdit().create();
  Logger.log('Trigger installato!');
}

function onTaskSheetEdit(e) {
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet();
    if (sheet.getName() !== 'Tasks') return;
    var row = e.range.getRow();
    if (row < 2) return;

    var col = e.range.getColumn();
    // Columns that require Calendar/Doc sync (1-indexed):
    // 2=company 3=assignee 4=assignDate 5=deadline 6=brief 7=driveUrl 9=status
    if ([2,3,4,5,6,7,9].indexOf(col) === -1) return;

    var rowData = sheet.getRange(row, 1, 1, 11).getValues()[0];
    var taskId          = String(rowData[0]||'');
    var company         = String(rowData[1]||'');
    var assignee        = String(rowData[2]||'');
    var assignDate      = formatDate_(rowData[3]);
    var deadline        = formatDate_(rowData[4]);
    var brief           = String(rowData[5]||'');
    var driveUrl        = String(rowData[6]||'');
    var docUrl          = String(rowData[7]||'');
    var status          = String(rowData[8]||'Da fare');
    var assignEventId   = String(rowData[9]||'');
    var deadlineEventId = String(rowData[10]||'');

    if (!taskId || !assignDate || !deadline) return;

    syncEventsForRow_(company, assignee, assignDate, deadline, brief, driveUrl, docUrl, status, assignEventId, deadlineEventId);

    // Update Google Doc for content-related column changes
    if (docUrl && [2,3,4,5,6,7].indexOf(col) !== -1) {
      updateDocBody_(docUrl, company, assignee, assignDate, deadline, brief, driveUrl);
    }
  } catch(err) {
    Logger.log('onTaskSheetEdit error: ' + err.message);
  }
}
