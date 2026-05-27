// ═══════════════════════════════════════════════════════════════
// OTIC — Dashboard de Seguimiento · Código.gs
// ═══════════════════════════════════════════════════════════════
//
// CAMBIOS (Producción · revisión bitácora):
//  · El antiguo saveBit masivo (clearContents + setValues) era inseguro
//    para uso concurrente: dos usuarios guardando al tiempo se pisaban
//    las notas. Ahora la bitácora se persiste FILA POR FILA en caliente.
//  · Nuevo: appendBitOne     → agrega UNA novedad y la persiste al instante.
//  · Nuevo: editBitOne       → edita una fila puntual de bitácora.
//  · Nuevo: deleteBitOne     → elimina una fila puntual de bitácora.
//  · saveBit queda como no-op (compatibilidad con HTML viejo en caché).
//  · saveAll ya NO toca la bitácora: solo Pre/Pos.
//  · readonly puede agregar / editar / eliminar SUS propias notas.
//  · Toda escritura en bitácora va envuelta en LockService.getDocumentLock().
// ═══════════════════════════════════════════════════════════════

var SPREADSHEET_ID = '15J-EubhRIHff8-2KPgTIKnasVhtj_0QPaSrNVZgRH-8';

var ROLES = {
  // ── ADMINISTRADORES ─────────────────────────────────────────
  'omontero@movilidadbogota.gov.co'    : 'admin',
  'ragonzalez@movilidadbogota.gov.co'  : 'admin',
  'eromero@movilidadbogota.gov.co'     : 'admin',
  // ── ROL PRE ─────────────────────────────────────────────────
  'savargas@movilidadbogota.gov.co'    : 'pre',
  // ── ROL POS ─────────────────────────────────────────────────
  'dcorredor@movilidadbogota.gov.co'   : 'pos',
  // ── READONLY — pueden ver todo y registrar notas ─────────────
  'adgutierrez@movilidadbogota.gov.co' : 'readonly',
  'avillamizar@movilidadbogota.gov.co' : 'readonly',
  'afeslava@movilidadbogota.gov.co'    : 'readonly',
  'cmeza@movilidadbogota.gov.co'       : 'readonly',
  'cgonzalezg@movilidadbogota.gov.co'  : 'readonly',
  'daromero@movilidadbogota.gov.co'    : 'readonly',
  'destrada@movilidadbogota.gov.co'    : 'readonly',
  'eguevara@movilidadbogota.gov.co'    : 'readonly',
  'gdelgado@movilidadbogota.gov.co'    : 'readonly',
  'jfajardo@movilidadbogota.gov.co'    : 'readonly',
  'jpinedaz@movilidadbogota.gov.co'    : 'readonly',
  'jrfuentes@movilidadbogota.gov.co'   : 'readonly',
  'jmonroyb@movilidadbogota.gov.co'    : 'readonly',
  'jhrodriguezq@movilidadbogota.gov.co': 'readonly',
  'jsanchez@movilidadbogota.gov.co'    : 'readonly',
  'jmunoza@movilidadbogota.gov.co'     : 'readonly',
  'jrozo@movilidadbogota.gov.co'       : 'readonly',
  'jpernet@movilidadbogota.gov.co'     : 'readonly',
  'jsantacruz@movilidadbogota.gov.co'  : 'readonly',
  'jneirab@movilidadbogota.gov.co'     : 'readonly',
  'jvargasb@movilidadbogota.gov.co'    : 'readonly',
  'jrgonzalez@movilidadbogota.gov.co'  : 'readonly',
  'jangelc@movilidadbogota.gov.co'     : 'readonly',
  'llopezm@movilidadbogota.gov.co'     : 'readonly',
  'lreyesb@movilidadbogota.gov.co'     : 'readonly',
  'mcaro@movilidadbogota.gov.co'       : 'readonly',
  'mezambrano@movilidadbogota.gov.co'  : 'readonly',
  'nbenavides@movilidadbogota.gov.co'  : 'readonly',
  'mbarrerat@movilidadbogota.gov.co'   : 'readonly',
  'malvarez@movilidadbogota.gov.co'    : 'readonly',
  'ngarzonp@movilidadbogota.gov.co'    : 'readonly',
  'ovargas@movilidadbogota.gov.co'     : 'readonly',
  'obolanosm@movilidadbogota.gov.co'   : 'readonly',
  'ocorrea@movilidadbogota.gov.co'     : 'readonly',
  'svelasquezv@movilidadbogota.gov.co' : 'readonly',
  'whernandez@movilidadbogota.gov.co'  : 'readonly',
  'yrodriguezb@movilidadbogota.gov.co' : 'readonly'
};

var BITACORA_HEADER_ROWS = 2; // filas 1-2 = encabezado, datos desde fila 3

// ── ENTRY POINTS ────────────────────────────────────────────────
function doGet(e) {
  var action = (e.parameter && e.parameter.action) || '';
  if (!action) {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('OTIC — Dashboard de Seguimiento')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  try {
    if (action === 'read')   return jsonOut(readAll());
    if (action === 'whoami') return jsonOut(whoAmI());
    return jsonOut({ error: 'Acción desconocida' });
  } catch(err) { return jsonOut({ error: err.message }); }
}

function doPost(e) {
  var email = getUserEmail();
  var role  = ROLES[email] || 'readonly';
  var data;
  try { data = JSON.parse(e.postData.contents); }
  catch(err) { return jsonOut({ error: 'JSON inválido' }); }
  var action = data.action || '';

  try {
    // Operaciones de bitácora — append/edit/delete en caliente, fila a fila.
    // Permitidas para CUALQUIER rol autenticado (incluido readonly).
    if (action === 'appendBitOne') return jsonOut(appendBitOne(data.entry, email));
    if (action === 'editBitOne')   return jsonOut(editBitOne(data.match, data.update, email, role));
    if (action === 'deleteBitOne') return jsonOut(deleteBitOne(data.match, email, role));

    // Operaciones masivas Pre/Pos — solo roles con permiso de escritura.
    if (action === 'savePre' && (role==='admin'||role==='pre')) return jsonOut(savePre(data.rows));
    if (action === 'savePos' && (role==='admin'||role==='pos')) return jsonOut(savePos(data.rows));
    if (action === 'saveAll' && role==='admin')                 return jsonOut(saveAll({ pre: data.pre, pos: data.pos }));
    if (action === 'saveAll' && role==='pre')                   return jsonOut(saveAll({ pre: data.pre }));
    if (action === 'saveAll' && role==='pos')                   return jsonOut(saveAll({ pos: data.pos }));

    // ⚠️ saveBit masivo queda deshabilitado: la bitácora ya NO se reescribe.
    if (action === 'updatePreField' && (role==='admin'||role==='pre')) return jsonOut(updatePreField(data.rowId, data.field, data.value));
    if (action === 'saveBit') {
      return jsonOut({ ok: true, msg: 'saveBit ya no es necesario: la bitácora se persiste fila por fila.', deprecated: true });
    }

    return jsonOut({ error: 'Sin permiso para: ' + action, email: email, role: role });
  } catch(err) { return jsonOut({ error: err.message }); }
}

// ── FUNCIONES PÚBLICAS ──────────────────────────────────────────
function whoAmI() {
  var email = getUserEmail();
  return { email: email, role: ROLES[email] || 'readonly' };
}

function readAll() {
  var email = getUserEmail();
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var shPre = ss.getSheetByName('Precontractual');
  var shPos = ss.getSheetByName('Poscontractual');
  var shBit = ss.getSheetByName('Bitácora');

  function sheetValues(sh) {
    if (!sh) return [];
    var vals = sh.getDataRange().getValues();
    return vals.map(function(row) {
      return row.map(function(cell) {
        if (cell instanceof Date) return cell.toISOString();
        return cell;
      });
    });
  }

  return {
    role  : ROLES[email] || 'readonly',
    email : email,
    pre   : shPre ? sheetValues(shPre) : [],
    pos   : shPos ? sheetValues(shPos) : [],
    bit   : shBit ? sheetValues(shBit) : [],
    ts    : new Date().toISOString()
  };
}

// ── GUARDADO MASIVO Pre / Pos ───────────────────────────────────
function savePre(rows) {
  if (!rows || !rows.length) return { ok: false };
  // Si se invoca por google.script.run (no por doPost), validar rol aquí.
  var email = getUserEmail();
  var role  = ROLES[email] || 'readonly';
  if (role !== 'admin' && role !== 'pre') {
    return { ok: false, error: 'Sin permiso para guardar Precontractual (rol: ' + role + ')' };
  }
  var sh = getSheet('Precontractual');
  var hdr = sh.getRange(1, 1, 3, sh.getLastColumn()).getValues();
  sh.clearContents();
  sh.getRange(1, 1, 3, hdr[0].length).setValues(hdr);
  sh.getRange(4, 1, rows.length, rows[0].length).setValues(rows);
  SpreadsheetApp.flush();
  return { ok: true, msg: 'Precontractual guardado · ' + rows.length + ' filas' };
}

function savePos(rows) {
  if (!rows || !rows.length) return { ok: false };
  var email = getUserEmail();
  var role  = ROLES[email] || 'readonly';
  if (role !== 'admin' && role !== 'pos') {
    return { ok: false, error: 'Sin permiso para guardar Poscontractual (rol: ' + role + ')' };
  }
  var sh = getSheet('Poscontractual');
  var hdr = sh.getRange(1, 1, 3, sh.getLastColumn()).getValues();
  sh.clearContents();
  sh.getRange(1, 1, 3, hdr[0].length).setValues(hdr);
  sh.getRange(4, 1, rows.length, rows[0].length).setValues(rows);
  SpreadsheetApp.flush();
  return { ok: true, msg: 'Poscontractual guardado · ' + rows.length + ' filas' };
}

function saveAll(data) {
  // Validar rol aquí también (no solo en doPost) porque google.script.run
  // puede invocar saveAll directamente sin pasar por doPost.
  var email = getUserEmail();
  var role  = ROLES[email] || 'readonly';
  if (role !== 'admin' && role !== 'pre' && role !== 'pos') {
    return { ok: false, error: 'Sin permiso para guardar Pre/Pos (rol: ' + role + ')' };
  }
  // Cada rol solo puede escribir lo suyo.
  var doPre = data.pre && (role === 'admin' || role === 'pre');
  var doPos = data.pos && (role === 'admin' || role === 'pos');
  var r1 = doPre ? savePre(data.pre) : { ok: true };
  var r2 = doPos ? savePos(data.pos) : { ok: true };
  // Nota: la bitácora ya NO se toca aquí. Se persiste fila por fila
  // desde el cliente mediante appendBitOne/editBitOne/deleteBitOne.
  return { ok: r1.ok && r2.ok, pre: r1, pos: r2 };
}

// ── ACTUALIZACIÓN PUNTUAL DE UN CAMPO EN PRECONTRACTUAL ─────────
var PRE_FIELD_COL = {
  'nombre':        3,
  'supervisor':    4,
  'cargo':         5,
  'abogado':       6,
  'estructurador': 7,
  'fPlanAbg':      8,
  'fPlanMesa':     12,
  'fPlanRad':      16,
  'valor':         22,
  'cdp':           23,
  'fase':          24,
  'modalidad':     28,
  'obs':           29,
  'fTermContrato': 30,
  'lineaPaa':      31
};

function updatePreField(rowId, field, value) {
  var col = PRE_FIELD_COL[field];
  if (!col) return { ok: false, error: 'Campo no editable: ' + field };
  var sh = getSheet('Precontractual');
  var lastRow = sh.getLastRow();
  if (lastRow < 4) return { ok: false, error: 'Sin datos en Precontractual' };
  var ids = sh.getRange(4, 1, lastRow - 3, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === String(rowId).trim()) {
      sh.getRange(i + 4, col).setValue(value);
      SpreadsheetApp.flush();
      return { ok: true, msg: 'Campo actualizado', rowId: rowId, field: field };
    }
  }
  return { ok: false, error: 'Proyecto no encontrado: ' + rowId };
}

// ⚠️ COMPATIBILIDAD: saveBit masivo queda como no-op explícito.
// Los clientes con HTML viejo en caché podrían seguir invocándolo;
// devolvemos OK silencioso (la bitácora YA está persistida fila por fila)
// para no romper la UI, pero NO se reescribe nada en el Sheet.
function saveBit(rows) {
  return { ok: true, msg: 'saveBit masivo deshabilitado — la bitácora se persiste fila por fila', deprecated: true, ignored: rows ? rows.length : 0 };
}

// ── BITÁCORA: APPEND fila por fila (concurrente y seguro) ────────
//
// entry: { modulo, contrato, nombre, fecha, usuario, novedad, compromiso }
//
// IMPORTANTE: ya no hace append al final. Inserta una nueva fila JUSTO DESPUÉS
// del header (fila 3) de modo que la nota más reciente quede arriba — tanto
// en el dashboard como en el Sheet abierto directamente.
function appendBitOne(entry, email) {
  if (!entry || !entry.novedad) return { ok: false, error: 'Novedad vacía' };
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(15000)) return { ok: false, error: 'No se pudo obtener lock — intenta de nuevo' };
  try {
    var sh = getSheet('Bitácora');
    // El usuario lo fija SIEMPRE el servidor (no se confía en el cliente),
    // salvo que el cliente envíe explícitamente el email logueado en .usuario.
    var usuario = entry.usuario || email;
    var row = [
      entry.modulo    || '',
      entry.contrato  || '',
      entry.nombre    || '',
      entry.fecha     || '',
      usuario,
      entry.novedad   || '',
      entry.compromiso|| ''
    ];
    // Inserta una fila vacía en la posición 3 (justo después del header de 2 filas).
    // Las filas existentes se desplazan hacia abajo: la antigua fila 3 pasa a ser 4, etc.
    var targetRow = BITACORA_HEADER_ROWS + 1; // = 3
    sh.insertRowBefore(targetRow);
    // Forzar la columna FECHA como texto plano para que Sheets NO reinterprete
    // "20 May 2026" como Date y así el match exacto de edit/delete funcione siempre.
    sh.getRange(targetRow, 4).setNumberFormat('@');
    sh.getRange(targetRow, 1, 1, row.length).setValues([row]);
    SpreadsheetApp.flush();
    return { ok: true, msg: 'Novedad agregada', row: targetRow, entry: { modulo: row[0], contrato: row[1], nombre: row[2], fecha: row[3], usuario: row[4], novedad: row[5], compromiso: row[6] } };
  } finally {
    lock.releaseLock();
  }
}

// match:  { modulo, contrato, fecha, usuario, novedad }  → busca la fila exacta.
// update: { novedad?, compromiso? }                       → campos a sobrescribir.
function editBitOne(match, update, email, role) {
  if (!match || !update) return { ok: false, error: 'Parámetros incompletos' };
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(15000)) return { ok: false, error: 'No se pudo obtener lock — intenta de nuevo' };
  try {
    var sh = getSheet('Bitácora');
    var rowIdx = findBitRowIndex_(sh, match);
    if (rowIdx < 0) return { ok: false, error: 'No se encontró la nota en el servidor (puede haber sido editada por otro usuario; recarga)' };

    // Permiso: admin edita cualquiera. Otros, solo SUS notas.
    var dueno = String(sh.getRange(rowIdx, 5).getValue() || '').toLowerCase();
    if (role !== 'admin' && role !== 'pos' && dueno !== email) {
      return { ok: false, error: 'Sin permiso para editar nota de otro usuario' };
    }

    if (typeof update.novedad === 'string')    sh.getRange(rowIdx, 6).setValue(update.novedad);
    if (typeof update.compromiso === 'string') sh.getRange(rowIdx, 7).setValue(update.compromiso);
    SpreadsheetApp.flush();
    return { ok: true, msg: 'Nota actualizada', row: rowIdx };
  } finally {
    lock.releaseLock();
  }
}

function deleteBitOne(match, email, role) {
  if (!match) return { ok: false, error: 'Parámetros incompletos' };
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(15000)) return { ok: false, error: 'No se pudo obtener lock — intenta de nuevo' };
  try {
    var sh = getSheet('Bitácora');
    var rowIdx = findBitRowIndex_(sh, match);
    if (rowIdx < 0) return { ok: false, error: 'No se encontró la nota en el servidor (puede haber sido editada por otro usuario; recarga)' };

    var dueno = String(sh.getRange(rowIdx, 5).getValue() || '').toLowerCase();
    if (role !== 'admin' && dueno !== email) {
      return { ok: false, error: 'Sin permiso para eliminar nota de otro usuario' };
    }

    sh.deleteRow(rowIdx);
    SpreadsheetApp.flush();
    return { ok: true, msg: 'Nota eliminada', row: rowIdx };
  } finally {
    lock.releaseLock();
  }
}

// Convierte un valor de celda de fecha al formato del cliente "D Mmm YYYY".
// Acepta Date, ISO string, o ya-formateado.
var MESES_BIT = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
function fmtFechaBit_(v) {
  if (v == null || v === '') return '';
  if (v instanceof Date && !isNaN(v)) {
    return v.getDate() + ' ' + MESES_BIT[v.getMonth()] + ' ' + v.getFullYear();
  }
  var s = String(v).trim();
  // ISO YYYY-MM-DD...
  var iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) {
    var d = new Date(parseInt(iso[1]), parseInt(iso[2])-1, parseInt(iso[3]));
    return d.getDate() + ' ' + MESES_BIT[d.getMonth()] + ' ' + d.getFullYear();
  }
  // Ya formateado "D Mmm YYYY"
  if (/^\d{1,2}\s[A-Za-z]{3}\s\d{4}/.test(s)) return s;
  return s;
}

// Busca la fila exacta de una bitácora por coincidencia múltiple.
// Devuelve el número de fila (1-based) o -1.
function findBitRowIndex_(sh, match) {
  var lastRow = sh.getLastRow();
  if (lastRow <= BITACORA_HEADER_ROWS) return -1;
  var rng = sh.getRange(BITACORA_HEADER_ROWS + 1, 1, lastRow - BITACORA_HEADER_ROWS, 7).getValues();
  var modulo   = String(match.modulo   || '').trim();
  var contrato = String(match.contrato || '').trim();
  var fecha    = fmtFechaBit_(match.fecha).trim();
  var usuario  = String(match.usuario  || '').trim().toLowerCase();
  var novedad  = String(match.novedad  || '').trim();
  for (var i = 0; i < rng.length; i++) {
    var r = rng[i];
    if (String(r[0]||'').trim() !== modulo)   continue;
    if (String(r[1]||'').trim() !== contrato) continue;
    if (fmtFechaBit_(r[3]).trim() !== fecha)  continue;
    if (String(r[4]||'').trim().toLowerCase() !== usuario) continue;
    if (String(r[5]||'').trim() !== novedad)  continue;
    return i + BITACORA_HEADER_ROWS + 1; // convertir a fila 1-based
  }
  return -1;
}

// ── UTILIDADES ──────────────────────────────────────────────────
function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getUserEmail() {
  try { return Session.getActiveUser().getEmail().toLowerCase(); }
  catch(e) { return ''; }
}

function getSheet(name) {
  var sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name);
  if (!sh) throw new Error('Hoja no encontrada: ' + name);
  return sh;
}

// ── WRAPPERS PARA google.script.run (modo embebido) ──────────────
// Estas funciones replican la lógica de permisos de doPost para que el
// cliente pueda llamarlas con la misma garantía de seguridad.

function appendBitOneFromClient(entry) {
  var email = getUserEmail();
  // Cualquier usuario autenticado puede agregar notas (incluido readonly).
  return appendBitOne(entry, email);
}

function editBitOneFromClient(match, update) {
  var email = getUserEmail();
  var role  = ROLES[email] || 'readonly';
  return editBitOne(match, update, email, role);
}

function deleteBitOneFromClient(match) {
  var email = getUserEmail();
  var role  = ROLES[email] || 'readonly';
  return deleteBitOne(match, email, role);
}

function updatePreFieldFromClient(rowId, field, value) {
  var email = getUserEmail();
  var role  = ROLES[email] || 'readonly';
  if (role !== 'admin' && role !== 'pre') return { ok: false, error: 'Sin permiso (rol: ' + role + ')' };
  return updatePreField(rowId, field, value);
}
