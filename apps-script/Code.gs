// =============================================================
// Code.gs — Backend del Strong Standard Personal Dashboard
// =============================================================
// Guarda los datos en la Google Sheet a la que está ligado este
// script (Extensiones > Apps Script desde la propia hoja).
//
// API:
//   GET  ?action=load              → { config, data } completo
//   POST { action:'saveDay', person, entry }   → guarda/actualiza un día
//   POST { action:'saveConfig', config }       → guarda la config inicial
//   POST { action:'saveAll', state }           → reemplaza todo (importar)
//
// Despliegue (Implementar > Nueva implementación > App web):
//   - Ejecutar como: Yo
//   - Quién tiene acceso: Cualquier usuario
// =============================================================

var ENTRY_COLS = ['person','date','calTrack','calories','calCompliance','veggies','meals','steps','weight','gym'];

function getSheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    if (name === 'entries') sh.getRange('B:B').setNumberFormat('@'); // fecha como texto
  } else if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
  }
  return sh;
}

function blank_(v) { return (v === null || v === undefined) ? '' : v; }

function ymdCell_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var s = String(v);
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;           // ya está bien
  var d = new Date(s);                                    // texto raro (ej. "Thu Jan 01 2099 ...")
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return s;
}

function loadAll_() {
  var entriesSh = getSheet_('entries', ENTRY_COLS);
  var configSh  = getSheet_('config', ['key','value']);
  var data = { bernardo: {}, manu: {} };

  var rows = entriesSh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    var person = r[0], date = ymdCell_(r[1]);
    if (!person || !date) continue;
    if (!data[person]) data[person] = {};
    data[person][date] = {
      date:          date,
      calTrack:      r[2] === '' ? null : String(r[2]),
      calories:      r[3] === '' ? null : Number(r[3]),
      calCompliance: r[4] === '' ? null : (r[4] === true || r[4] === 'true' || r[4] === 'TRUE'),
      veggies:       r[5] === '' ? 0    : Number(r[5]),
      meals:         r[6] === '' ? null : Number(r[6]),
      steps:         r[7] === '' ? null : Number(r[7]),
      weight:        r[8] === '' ? null : Number(r[8]),
      gym:           r[9] === '' ? null : String(r[9])
    };
  }

  var config = null;
  var crows = configSh.getDataRange().getValues();
  var cmap = {};
  for (var j = 1; j < crows.length; j++) { if (crows[j][0]) cmap[crows[j][0]] = crows[j][1]; }
  if (cmap.startDate) {
    config = {
      startDate:    ymdCell_(cmap.startDate),
      bStartWeight: Number(cmap.bStartWeight),
      bGoalWeight:  Number(cmap.bGoalWeight),
      mStartWeight: Number(cmap.mStartWeight),
      mGoalWeight:  Number(cmap.mGoalWeight)
    };
  }
  return { config: config, data: data };
}

function rowFor_(person, e, date) {
  return [person, e.date || date, blank_(e.calTrack), blank_(e.calories), blank_(e.calCompliance),
          blank_(e.veggies), blank_(e.meals), blank_(e.steps), blank_(e.weight), blank_(e.gym)];
}

function saveDay_(person, entry) {
  var sh = getSheet_('entries', ENTRY_COLS);
  var rows = sh.getDataRange().getValues();
  var rowIdx = -1;
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === person && ymdCell_(rows[i][1]) === entry.date) { rowIdx = i + 1; break; }
  }
  var vals = rowFor_(person, entry, entry.date);
  if (rowIdx === -1) sh.appendRow(vals);
  else sh.getRange(rowIdx, 1, 1, vals.length).setValues([vals]);
}

function saveConfig_(config) {
  var sh = getSheet_('config', ['key','value']);
  sh.clearContents();
  sh.appendRow(['key','value']);
  var pairs = [
    ['startDate',    config.startDate],
    ['bStartWeight', config.bStartWeight],
    ['bGoalWeight',  config.bGoalWeight],
    ['mStartWeight', config.mStartWeight],
    ['mGoalWeight',  config.mGoalWeight]
  ];
  sh.getRange(2, 1, pairs.length, 2).setValues(pairs);
}

function saveAll_(state) {
  var sh = getSheet_('entries', ENTRY_COLS);
  sh.clearContents();
  sh.appendRow(ENTRY_COLS);
  sh.getRange('B:B').setNumberFormat('@');
  var out = [];
  ['bernardo','manu'].forEach(function(person) {
    var pd = (state.data && state.data[person]) || {};
    Object.keys(pd).forEach(function(date) { out.push(rowFor_(person, pd[date], date)); });
  });
  if (out.length) sh.getRange(2, 1, out.length, ENTRY_COLS.length).setValues(out);
  if (state.config) saveConfig_(state.config);
}

function doGet(e) {
  var result;
  try {
    var action = (e.parameter && e.parameter.action) || 'load';
    if (action === 'load') result = { success: true, data: loadAll_() };
    else result = { success: false, error: 'Acción desconocida: ' + action };
  } catch (err) {
    result = { success: false, error: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var result;
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var payload = JSON.parse(e.postData.contents);
    switch (payload.action) {
      case 'saveDay':    saveDay_(payload.person, payload.entry); result = { success: true }; break;
      case 'saveConfig': saveConfig_(payload.config);             result = { success: true }; break;
      case 'saveAll':    saveAll_(payload.state);                 result = { success: true }; break;
      default:           result = { success: false, error: 'Acción desconocida: ' + payload.action };
    }
  } catch (err) {
    result = { success: false, error: err.message };
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}
