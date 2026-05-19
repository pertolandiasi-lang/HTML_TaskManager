// ============================================================
// CONFIGURAZIONE — aggiorna questi valori prima del deploy
// ============================================================
const SHEET_ID    = '1aUW4iKrUhUQBNISsY2xISgNHlfOc7e6dDjuqVzHgC68'; // ID dello Sheet (dalla URL)
const CALENDAR_ID = 'primary';

// colorId Calendar per stato
// 5=Banana(giallo), 8=Graphite(grigio), 10=Basil(verde), 11=Tomato(rosso)
const COLOR_MAP = {
  'Da fare':    { assign: '10', deadline: '11' },
  'In lavoro':  { assign: '10', deadline: '5'  },
  'Completato': { assign: '8',  deadline: '10' },
};

// ============================================================
// ENTRY POINT WEBAPP
// ============================================================
function doGet(e) {
  const email = Session.getActiveUser().getEmail();
  const role  = getUserRole_(email);
  const name  = getUserName_(email);

  let teamJson = '[]';
  if (role === 'manager') {
    try {
      const result = getTeam();
      if (result.success) teamJson = JSON.stringify(result.data);
    } catch(err) { Logger.log('doGet team error: ' + err.message); }
  }

  const tpl = HtmlService.createTemplateFromFile('Index');
  tpl.userEmail  = email;
  tpl.userRole   = role;
  tpl.userName   = name;
  tpl.oauthToken = ScriptApp.getOAuthToken();
  tpl.teamJson   = teamJson;

  return tpl.evaluate()
    .setTitle('Task Manager')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
// HELPER PRIVATI
// ============================================================
function getUserRole_(email) {
  try {
    const data = SpreadsheetApp.openById(SHEET_ID)
      .getSheetByName('Team').getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString().trim().toLowerCase() === email.toLowerCase()) {
        return data[i][2].toString().trim().toLowerCase();
      }
    }
  } catch (err) {
    Logger.log('getUserRole_ error: ' + err.message);
  }
  return 'unauthorized';
}

function getUserName_(email) {
  try {
    const data = SpreadsheetApp.openById(SHEET_ID)
      .getSheetByName('Team').getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString().trim().toLowerCase() === email.toLowerCase()) {
        return data[i][0].toString().trim();
      }
    }
  } catch (err) {}
  return email;
}

function addOneDay_(dateStr) {
  const d = new Date(dateStr + 'T12:00:00Z');
  d.setUTCDate(d.getUTCDate() + 1);
  return Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');
}

function formatDate_(val) {
  if (!val) return '';
  if (val instanceof Date) return Utilities.formatDate(val, 'UTC', 'yyyy-MM-dd');
  return val.toString().trim();
}

function tryPatchColor_(eventId, colorId) {
  try {
    Calendar.Events.patch({ colorId: colorId }, CALENDAR_ID, eventId);
  } catch (e) {
    Logger.log('Calendar patch error (eventId=' + eventId + '): ' + e.message);
  }
}

// ============================================================
// API PUBBLICA — chiamata via google.script.run
// ============================================================

function getTeam() {
  try {
    const data = SpreadsheetApp.openById(SHEET_ID)
      .getSheetByName('Team').getDataRange().getValues();
    const members = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][1]) { // ha email
        members.push({
          name:  data[i][0].toString().trim(),
          email: data[i][1].toString().trim(),
          role:  (data[i][2] || '').toString().trim().toLowerCase(),
        });
      }
    }
    return { success: true, data: members };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function createTask(taskData) {
  try {
    if (!taskData.company || !taskData.assignee ||
        !taskData.assignDate || !taskData.deadline || !taskData.brief) {
      return { success: false, error: 'Compila tutti i campi obbligatori.' };
    }
    if (taskData.deadline < taskData.assignDate) {
      return { success: false, error: 'La deadline deve essere uguale o successiva alla data di assegnazione.' };
    }

    const now    = new Date();
    const taskId = now.getTime().toString();

    // 1. Crea Google Doc con il brief
    const docTitle = 'Brief - ' + taskData.company + ' - ' + taskData.assignDate;
    const doc = DocumentApp.create(docTitle);
    doc.getBody().setText(taskData.brief);
    doc.saveAndClose();
    const docUrl = doc.getUrl();

    // 2. Testo descrizione eventi Calendar
    const desc = [
      'Assegnatario: ' + taskData.assignee,
      '',
      'Brief:',
      taskData.brief,
      taskData.driveUrl ? '\nCartella Drive:\n' + taskData.driveUrl : '',
      '\nGoogle Doc Brief:\n' + docUrl,
    ].join('\n');

    // 3. Evento assegnazione (verde, all-day)
    const assignEvent = Calendar.Events.insert({
      summary:     taskData.company,
      description: desc,
      start: { date: taskData.assignDate },
      end:   { date: addOneDay_(taskData.assignDate) },
      colorId: '10',
    }, CALENDAR_ID);

    // 4. Evento deadline (rosso, all-day)
    const deadlineEvent = Calendar.Events.insert({
      summary:     '⚠️ DEADLINE - ' + taskData.company,
      description: desc,
      start: { date: taskData.deadline },
      end:   { date: addOneDay_(taskData.deadline) },
      colorId: '11',
    }, CALENDAR_ID);

    // 5. Scrivi riga sul foglio Tasks
    SpreadsheetApp.openById(SHEET_ID).getSheetByName('Tasks').appendRow([
      taskId,
      taskData.company,
      taskData.assignee,
      taskData.assignDate,
      taskData.deadline,
      taskData.brief,
      taskData.driveUrl || '',
      docUrl,
      'Da fare',
      assignEvent.id,
      deadlineEvent.id,
      now.toISOString(),
    ]);

    return { success: true, taskId: taskId };
  } catch (err) {
    Logger.log('createTask error: ' + err.message);
    return { success: false, error: err.message };
  }
}

function getMyTasks(email) {
  try {
    const data = SpreadsheetApp.openById(SHEET_ID)
      .getSheetByName('Tasks').getDataRange().getValues();
    const tasks = [];

    for (let i = 1; i < data.length; i++) {
      if (!data[i][2]) continue;
      if (data[i][2].toString().trim().toLowerCase() !== email.toLowerCase()) continue;

      tasks.push({
        id:              data[i][0].toString(),
        company:         data[i][1].toString(),
        assignee:        data[i][2].toString(),
        assignDate:      formatDate_(data[i][3]),
        deadline:        formatDate_(data[i][4]),
        brief:           data[i][5].toString(),
        driveUrl:        data[i][6].toString(),
        docUrl:          data[i][7].toString(),
        status:          data[i][8].toString() || 'Da fare',
        assignEventId:   data[i][9].toString(),
        deadlineEventId: data[i][10].toString(),
      });
    }

    tasks.sort((a, b) => new Date(a.deadline) - new Date(b.deadline));
    return { success: true, data: tasks };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function updateTaskStatus(taskId, newStatus) {
  try {
    const valid = ['Da fare', 'In lavoro', 'Completato'];
    if (!valid.includes(newStatus)) {
      return { success: false, error: 'Stato non valido.' };
    }

    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Tasks');
    const data  = sheet.getDataRange().getValues();
    let rowIndex = -1, assignId = '', deadlineId = '';

    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === taskId.toString()) {
        rowIndex   = i + 1;
        assignId   = data[i][9].toString();
        deadlineId = data[i][10].toString();
        break;
      }
    }

    if (rowIndex === -1) return { success: false, error: 'Task non trovato.' };

    // Aggiorna colonna Stato (col 9, 1-indexed)
    sheet.getRange(rowIndex, 9).setValue(newStatus);

    // Aggiorna colori Calendar
    const colors = COLOR_MAP[newStatus];
    if (assignId)   tryPatchColor_(assignId,   colors.assign);
    if (deadlineId) tryPatchColor_(deadlineId, colors.deadline);

    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}
