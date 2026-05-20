/* ────────────────────────────────────────────────────────────────────────
   HTML Task — script DEL FOGLIO (multi-assegnatari)

   Questo NON è il progetto Apps Script principale. Va incollato dentro
   lo script legato al foglio Google:
     1. Apri il foglio Google "Tasks"
     2. Menu  Estensioni ▸ Apps Script
     3. Cancella tutto e incolla questo file, poi Salva (icona dischetto)
     4. Ricarica il foglio Google → comparirà il menu "👥 Assegnatari"
   ──────────────────────────────────────────────────────────────────────── */

var WEBAPP_URL = 'https://script.google.com/macros/s/AKfycbxPO690wTE3ntdbZCS2f5CFETLn5JsqSdxlKpQFmfgmkk73lZdNspNQPK4wHdjCdHzX-Q/exec';
var SYNC_KEY   = 'tf_syncrow_v56';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('👥 Assegnatari')
    .addItem('Assegna persone alla riga selezionata', 'showAssigneeDialog')
    .addToUi();
}

function esc_(s) {
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}

function showAssigneeDialog() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  if (sheet.getName() !== 'Tasks') {
    ui.alert('Apri il foglio "Tasks" e seleziona una riga di task.');
    return;
  }
  var row = sheet.getActiveRange().getRow();
  if (row < 2) {
    ui.alert('Seleziona una riga di task (non l\'intestazione).');
    return;
  }

  var taskName = String(sheet.getRange(row, 2).getValue() || ('Riga ' + row));
  var current = String(sheet.getRange(row, 3).getValue() || '')
    .toLowerCase().split(',').map(function(s){ return s.trim(); });

  var trows = ss.getSheetByName('Team').getDataRange().getValues();
  var team = [];
  for (var i = 1; i < trows.length; i++) {
    if (trows[i][1]) team.push({
      name:  String(trows[i][0] || ''),
      email: String(trows[i][1]).trim(),
      role:  String(trows[i][2] || '').toLowerCase()
    });
  }

  var checkboxes = team.map(function(m) {
    var ck = current.indexOf(m.email.toLowerCase()) >= 0 ? ' checked' : '';
    return '<label><input type="checkbox" value="'+esc_(m.email)+'"'+ck+'>'
      + '<span>'+esc_(m.name)+(m.role==='manager'?' <em>(manager)</em>':'')+'</span></label>';
  }).join('');

  var html = '<!DOCTYPE html><html><head><base target="_top"><style>'
    + 'body{font-family:-apple-system,Roboto,Arial,sans-serif;margin:0;padding:16px;color:#202124;}'
    + '.hd{font-size:12px;color:#5f6368;}.tn{font-size:16px;font-weight:600;margin:2px 0 14px;}'
    + '.list{max-height:300px;overflow-y:auto;border:1px solid #e0e0e0;border-radius:10px;padding:4px;}'
    + 'label{display:flex;align-items:center;gap:10px;padding:9px 8px;border-radius:8px;cursor:pointer;font-size:14px;}'
    + 'label:hover{background:#f1f1f4;}label em{color:#5f6368;font-style:italic;}'
    + 'input[type=checkbox]{width:17px;height:17px;accent-color:#5e17eb;}'
    + '.row{display:flex;gap:8px;margin-top:14px;}'
    + 'button{flex:1;padding:10px;border-radius:8px;border:none;font-size:14px;font-weight:600;cursor:pointer;font-family:inherit;}'
    + '.save{background:#5e17eb;color:#fff;}.cancel{background:#f1f1f4;color:#202124;}button:disabled{opacity:.6;}'
    + '</style></head><body>'
    + '<div class="hd">Assegna persone a:</div>'
    + '<div class="tn">'+esc_(taskName)+'</div>'
    + '<div class="list">'+(checkboxes || '<div style="padding:12px;color:#5f6368;">Team vuoto</div>')+'</div>'
    + '<div class="row"><button class="cancel" onclick="google.script.host.close()">Annulla</button>'
    + '<button class="save" onclick="save()">Salva</button></div>'
    + '<script>function save(){'
    + 'var c=document.querySelectorAll("input[type=checkbox]:checked"),e=[];'
    + 'for(var i=0;i<c.length;i++)e.push(c[i].value);'
    + 'var b=document.querySelector(".save");b.textContent="Salvataggio…";b.disabled=true;'
    + 'google.script.run.withSuccessHandler(function(){google.script.host.close();})'
    + '.withFailureHandler(function(err){b.textContent="Salva";b.disabled=false;alert("Errore: "+err.message);})'
    + '.setRowAssignees('+row+',e.join(", "));}<\/script>'
    + '</body></html>';

  ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(340).setHeight(440), 'Assegnatari');
}

function setRowAssignees(row, csv) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tasks')
    .getRange(row, 3).setValue(csv);
  // Tell the main web app to re-sync this row's Calendar events + Google Doc.
  try {
    UrlFetchApp.fetch(
      WEBAPP_URL + '?action=syncRow&key=' + encodeURIComponent(SYNC_KEY) + '&row=' + row,
      { muteHttpExceptions: true }
    );
  } catch (e) {
    // sheet value is saved regardless; sync will also happen on next normal edit
  }
}
