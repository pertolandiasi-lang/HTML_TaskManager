const SHEET_ID    = '1aUW4iKrUhUQBNISsY2xISgNHlfOc7e6dDjuqVzHgC68';
const CALENDAR_ID = 'primary';
const COLOR_MAP   = {
  'Da fare':    { assign: '10', deadline: '11' },
  'In lavoro':  { assign: '10', deadline: '5'  },
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
  try { return dispatch_(e.parameter); }
  catch(err) { return respond_({ error: err.message }); }
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

function tryPatchColor_(eventId, colorId) {
  try { Calendar.Events.patch({ colorId: colorId }, CALENDAR_ID, eventId); }
  catch(e) { Logger.log('Patch color err: ' + e.message); }
}

function tryDeleteEvent_(eventId) {
  try { Calendar.Events.remove(CALENDAR_ID, eventId); }
  catch(e) { Logger.log('Delete event err: ' + e.message); }
}
