const SHEET_ID     = '1aUW4iKrUhUQBNISsY2xISgNHlfOc7e6dDjuqVzHgC68';
const DRIVE_FOLDER = '1qZdcgxj1neycYTj4BmTQhI1SbGklgowc';
const BACKUP_FOLDER = '1ct1aOAP0Qapa3AlcGKJmWxqp5w7ilPCp'; // Drive folder for daily Sheet backups
const BACKUP_RETENTION_DAYS = 30;

// Calendario editoriale: è un ALTRO foglio, non questo. Non serve un secondo
// progetto Apps Script né un secondo login: il web app gira come USER_DEPLOYING
// e quell'account possiede entrambi i file.
const CAL_SHEET_ID = '1AA97-4g9UyFzkSse9zoed6bNY3HkIzb84o-adm96sVM';

// Etichetta della build. Serve a rispondere alla domanda "il codice che ho
// salvato è davvero quello che sta girando?", che senza un modo di chiederlo
// dall'esterno si può solo dedurre dagli errori.
//
// VA CAMBIATA A OGNI MODIFICA DI QUESTO FILE. Lasciandola ferma la sonda
// risponde "tutto a posto" anche su un deployment vecchio di tre versioni, che
// è peggio di non avere nessuna sonda.
const BUILD = '2026-08-08-liste-unione';

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
  // Sonda senza autenticazione: dice solo quale build sta servendo il
  // deployment. Non tocca dati e non rivela niente che non sia già pubblico,
  // ma permette di verificare da fuori se una nuova versione è andata online
  // invece di dedurlo dagli errori dell'app.
  if (p.action === 'ping') return respond_({ ok: true, build: BUILD });

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

    // Calendario editoriale, solo manager: i piani dei clienti social non
    // c'entrano niente coi task dei dipendenti.
    case 'getCalendario':
      requireRole_(callerEmail, 'manager');
      return respond_({ posts: getCalendarioPosts_(p.from, p.to), liste: getCalendarioListe_() });

    case 'creaPost':
      requireRole_(callerEmail, 'manager');
      return respond_({ riga: creaPost_(JSON.parse(p.post)) });

    case 'aggiornaPost':
      requireRole_(callerEmail, 'manager');
      aggiornaPost_(parseInt(p.riga, 10), p.chiave, JSON.parse(p.post));
      return respond_({ ok: true });

    case 'eliminaPost':
      requireRole_(callerEmail, 'manager');
      eliminaPost_(parseInt(p.riga, 10), p.chiave);
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

// ── CALENDARIO EDITORIALE ────────────────────────────────────────────────────

/**
 * I post del calendario editoriale in un intervallo di date.
 *
 * Legge SOLO il tab DB, che è l'unico posto dove si scrive davvero in quel
 * foglio: Calendario, Matrice, Settimana e Dashboard sono quattro rendering
 * delle stesse righe, fatti a formule perché in un foglio non puoi fare altro.
 * Rifarli qui in HTML sarebbe reimplementare un foglio di calcolo nel browser:
 * il frontend riceve i dati grezzi e disegna le viste che gli servono.
 *
 * L'intervallo non è un lusso: il piano è di ~1.200 righe (30 clienti x 3 mesi),
 * e spedirle tutte a ogni apertura sarebbe mezzo megabyte di JSON per vedere
 * una settimana.
 */
function getCalendarioPosts_(from, to) {
  var sh   = calDB_();
  var last = sh.getLastRow();
  if (last < 2) return [];

  // A..O: le colonne che si compilano. P..T sono calcolate e servono al foglio,
  // non a noi.
  var rows = sh.getRange(2, 1, last - 1, 15).getValues();
  var tz   = Session.getScriptTimeZone();
  var da   = from ? new Date(from + 'T00:00:00') : null;
  var al   = to   ? new Date(to   + 'T23:59:59') : null;

  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    // Una data scritta come testo non è un post: nel foglio sparirebbe dalle
    // viste senza dare errore, e qui la si salta per lo stesso motivo.
    if (!(r[0] instanceof Date)) continue;
    if (da && r[0] < da) continue;
    if (al && r[0] > al) continue;

    out.push({
      riga:     i + 2,                     // riga nel DB, per aprire il foglio sul punto giusto
      data:     Utilities.formatDate(r[0], tz, 'yyyy-MM-dd'),
      ora:      (r[1] instanceof Date) ? Utilities.formatDate(r[1], tz, 'HH:mm') : '',
      cliente:  String(r[2]  || ''),
      canale:   String(r[3]  || ''),
      formato:  String(r[4]  || ''),
      pilastro: String(r[5]  || ''),
      titolo:   String(r[6]  || ''),
      copy:     String(r[7]  || ''),
      hashtag:  String(r[8]  || ''),
      asset:    String(r[9]  || ''),
      stato:    String(r[10] || ''),
      note:     String(r[11] || ''),
      link:     String(r[12] || ''),
      // Impronta della riga com'era quando l'abbiamo letta. Torna indietro a
      // ogni modifica: se nel frattempo qualcuno ha inserito o riordinato righe
      // nel foglio, il numero di riga punta a un altro post e senza questo
      // controllo ci scriveremmo sopra. Vedi verificaRiga_.
      chiave:   chiaveRiga_(r, tz)
    });
  }
  return out;
}

/** L'impronta di una riga del DB: cliente + data + ora. */
function chiaveRiga_(r, tz) {
  return String(r[2] || '') + '|'
       + ((r[0] instanceof Date) ? Utilities.formatDate(r[0], tz, 'yyyy-MM-dd') : '') + '|'
       + ((r[1] instanceof Date) ? Utilities.formatDate(r[1], tz, 'HH:mm') : '');
}

/** Il tab DB del calendario, con gli errori parlanti già applicati. */
function calDB_() {
  var ss;
  try {
    ss = SpreadsheetApp.openById(CAL_SHEET_ID);
  } catch (e) {
    throw new Error(
      'Non riesco ad aprire il file del calendario editoriale. Il web app gira ' +
      'come l\'account che l\'ha distribuito: se il foglio appartiene a un altro ' +
      'account, va condiviso con quello. ID cercato: ' + CAL_SHEET_ID
    );
  }
  var sh = ss.getSheetByName('DB');
  if (!sh) {
    var nomi = ss.getSheets().map(function (s) { return s.getName(); }).join(', ');
    throw new Error('Nel calendario editoriale non c\'è nessun tab chiamato "DB". Trovate: ' + nomi);
  }
  return sh;
}

/**
 * Le tendine, lette dal tab Liste del calendario. Vengono da lì e non da una
 * copia qui dentro: se aggiungi un cliente nel foglio deve comparire nell'app
 * senza che nessuno rideployi niente.
 */
function getCalendarioListe_() {
  var out = { clienti: [], canali: [], formati: [], stati: [], pilastri: [] };

  var L = SpreadsheetApp.openById(CAL_SHEET_ID).getSheetByName('Liste');
  if (L) {
    var colonna = function (lettera) {
      return L.getRange(lettera + '2:' + lettera + '60').getValues()
        .map(function (r) { return String(r[0] || '').trim(); })
        .filter(function (v) { return v !== ''; });
    };
    out.clienti  = colonna('A');
    out.canali   = colonna('C');
    out.formati  = colonna('E');
    out.stati    = colonna('H');
    out.pilastri = colonna('K');
  }

  // Se il tab Liste manca o è vuoto (succede se il foglio è stato ricreato o
  // copiato senza far girare lo script che lo genera) le tendine arriverebbero
  // vuote e non si potrebbe compilare niente. Si ricavano allora dai post che
  // ci sono già, e per i vocabolari fissi si usa quello dello script.
  var sh   = calDB_();
  var last = sh.getLastRow();
  var rows = last > 1 ? sh.getRange(2, 1, last - 1, 15).getValues() : [];
  function distinti(idx) {
    var v = [], visti = {};
    rows.forEach(function (r) {
      var s = String(r[idx] || '').trim();
      if (s && !visti[s]) { visti[s] = 1; v.push(s); }
    });
    return v.sort(function (a, b) { return a.localeCompare(b, 'it'); });
  }
  // Unione, non ripiego. Prendere i valori "solo se la lista è vuota" sembrava
  // ragionevole ma dava tendine monche: i post generati dallo script hanno
  // tutti canale Instagram, quindi bastava una riga nel DB perché il canale
  // mostrasse quell'unica voce e nascondesse tutte le altre.
  function unisci() {
    var v = [], visti = {};
    for (var i = 0; i < arguments.length; i++) {
      (arguments[i] || []).forEach(function (x) {
        var s = String(x || '').trim();
        if (s && !visti[s]) { visti[s] = 1; v.push(s); }
      });
    }
    return v;
  }

  out.clienti  = unisci(out.clienti, distinti(2));
  out.canali   = unisci(out.canali, distinti(3),  ['Instagram','Facebook','TikTok','LinkedIn','YouTube']);
  out.formati  = unisci(out.formati, distinti(4), ['Post','Carosello','Reel','Storia','Video','Articolo']);
  out.pilastri = unisci(out.pilastri, distinti(5),['Prodotto','Dietro le quinte','Educativo','Social proof','Promo','Community']);
  out.stati    = unisci(out.stati, distinti(10),  ['Idea','Da produrre','In revisione','Approvato','Programmato','Pubblicato']);

  return out;
}

/**
 * Le 15 colonne A..O nell'ordine del DB. Le P..T sono formule: non si toccano.
 *
 * `esistente` è la riga com'è adesso nel foglio, e serve per le colonne che
 * l'app non mostra: Views e Interazioni si compilano a mano nel foglio dopo la
 * pubblicazione, e riscriverle vuote a ogni salvataggio cancellerebbe dati che
 * l'utente non si aspetta nemmeno di stare toccando.
 */
function rigaDaPost_(p, esistente) {
  var vecchia = esistente || [];
  var data = p.data ? new Date(p.data + 'T00:00:00') : '';
  var ora  = '';
  if (p.ora) {
    var hm = p.ora.split(':');
    ora = new Date(1899, 11, 30, parseInt(hm[0], 10) || 0, parseInt(hm[1], 10) || 0);
  }
  return [
    data, ora, p.cliente || '', p.canale || '', p.formato || '', p.pilastro || '',
    p.titolo || '', p.copy || '', p.hashtag || '', p.asset || '',
    p.stato || 'Idea', p.note || '', p.link || '',
    vecchia.length > 13 ? vecchia[13] : '',   // Views
    vecchia.length > 14 ? vecchia[14] : ''    // Interazioni
  ];
}

function creaPost_(p) {
  if (!p.data)    throw new Error('Manca la data');
  if (!p.cliente) throw new Error('Manca il cliente');
  var sh = calDB_();
  var riga = sh.getLastRow() + 1;
  sh.getRange(riga, 1, 1, 15).setValues([rigaDaPost_(p)]);
  return riga;
}

/**
 * Controlla che la riga sia ancora quella che il browser credeva di modificare.
 * Senza, basta che qualcuno inserisca una riga nel foglio mentre hai la vista
 * aperta e la modifica finisce sul post sbagliato — in silenzio.
 */
function verificaRiga_(sh, riga, chiave) {
  if (!riga || riga < 2) throw new Error('Riga non valida');
  var r = sh.getRange(riga, 1, 1, 15).getValues()[0];
  if (chiaveRiga_(r, Session.getScriptTimeZone()) !== chiave) {
    throw new Error('Il foglio è cambiato da quando hai aperto la vista: ricarica e riprova.');
  }
  return r;
}

function aggiornaPost_(riga, chiave, p) {
  var sh = calDB_();
  var vecchia = verificaRiga_(sh, riga, chiave);
  sh.getRange(riga, 1, 1, 15).setValues([rigaDaPost_(p, vecchia)]);
}

function eliminaPost_(riga, chiave) {
  var sh = calDB_();
  verificaRiga_(sh, riga, chiave);
  sh.deleteRow(riga);
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
