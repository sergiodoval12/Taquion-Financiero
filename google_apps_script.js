// ============================================================
// TAQUION CF — Google Apps Script API
// ============================================================
// Pegar este código en: Extensiones → Apps Script del Google Sheet
// Deploy: Nueva implementación → Web App → "Cualquiera con el enlace"
// ============================================================

const SHEET_MOV = 'BASE DE DATOS MOVIMIENTOS';
const SHEET_BD_TQN = 'Deuda prestamos bancarios TQN';
const SHEET_BD_LMS = 'Deuda prest. banc LMS';

// GET: Read data
function doGet(e) {
  const action = e.parameter.action || 'all';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let result;

  try {
    if (action === 'movimientos') {
      result = getMovimientos(ss);
    } else if (action === 'deuda_tqn') {
      result = getDeudaTQN(ss);
    } else if (action === 'deuda_lms') {
      result = getDeudaLMS(ss);
    } else if (action === 'all') {
      result = {
        mov: getMovimientos(ss),
        bd: getDeudaTQN(ss),
        bdl: getDeudaLMS(ss),
        meta: { lastSync: new Date().toISOString(), source: 'Google Sheets' }
      };
    } else if (action === 'ping') {
      result = { ok: true, ts: new Date().toISOString() };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// POST: Write changes
function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const body = JSON.parse(e.postData.contents);
  const action = body.action;

  try {
    let result;

    if (action === 'update_movimiento') {
      // Update a single movement row
      result = updateMovimiento(ss, body.rowIndex, body.field, body.value);
    } else if (action === 'add_movimientos') {
      // Add new movement rows (from sandbox implementations)
      result = addMovimientos(ss, body.rows);
    } else if (action === 'delete_movimiento') {
      // Mark a movement as deleted (set value to 0)
      result = deleteMovimiento(ss, body.rowIndex);
    } else if (action === 'batch_update') {
      // Batch update multiple changes at once
      result = batchUpdate(ss, body.changes);
    } else {
      result = { error: 'Unknown action: ' + action };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---- READ FUNCTIONS ----

function getMovimientos(ss) {
  const ws = ss.getSheetByName(SHEET_MOV);
  if (!ws) return [];
  const data = ws.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim().toLowerCase());

  // Map column names to our internal format
  const colMap = {};
  headers.forEach((h, i) => {
    if (h.includes('fecha')) colMap.f = i;
    else if (h.includes('estado') || h === 'eo') colMap.eo = i;
    else if (h.includes('empresa') || h === 'emp') colMap.emp = i;
    else if (h.includes('bancarizado') || h === 'bn') colMap.bn = i;
    else if (h.includes('categoría') || h.includes('categoria') || h === 'cat') colMap.cat = i;
    else if (h.includes('tipo') || h === 't') colMap.t = i;
    else if (h.includes('marco') || h === 'm') colMap.m = i;
    else if (h.includes('detalle') || h === 'd') colMap.d = i;
    else if (h.includes('item') || h === 'i') colMap.i = i;
    else if (h.includes('entidad') || h === 'en') colMap.en = i;
    else if (h.includes('movimiento_orig') || h.includes('monto')) colMap.v = i;
    else if (h.includes('movimiento') && !h.includes('orig')) colMap.v_div = i;
  });

  const rows = [];
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    // Skip empty rows
    if (!row[colMap.f] && !row[colMap.v] && !row[colMap.v_div]) continue;

    let fecha = '';
    if (row[colMap.f] instanceof Date) {
      fecha = row[colMap.f].toISOString().slice(0, 10);
    } else if (row[colMap.f]) {
      fecha = String(row[colMap.f]).slice(0, 10);
    }

    let valor = 0;
    if (colMap.v !== undefined && row[colMap.v]) {
      valor = Number(row[colMap.v]) || 0;
    } else if (colMap.v_div !== undefined && row[colMap.v_div]) {
      valor = (Number(row[colMap.v_div]) || 0) * 1000; // Convert from thousands
    }

    rows.push({
      _row: r + 1, // 1-indexed row number in sheet (for updates)
      f: fecha,
      eo: String(row[colMap.eo] || '').trim() === 'R' ? 'R' : 'P',
      emp: String(row[colMap.emp] || '').trim(),
      bn: String(row[colMap.bn] || '').trim(),
      cat: String(row[colMap.cat] || '').trim(),
      t: String(row[colMap.t] || '').trim(),
      m: String(row[colMap.m] || '').trim(),
      d: String(row[colMap.d] || '').trim(),
      i: String(row[colMap.i] || '').trim(),
      en: String(row[colMap.en] || '').trim(),
      v: valor
    });
  }
  return rows;
}

function getDeudaTQN(ss) {
  const ws = ss.getSheetByName(SHEET_BD_TQN);
  if (!ws) return [];
  const data = ws.getDataRange().getValues();
  // Parse the debt schedule - structure may vary
  const result = [];
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    let mes = '';
    if (row[0] instanceof Date) {
      mes = row[0].toISOString().slice(0, 7);
    } else if (row[0]) {
      const s = String(row[0]);
      if (s.match(/^\d{4}-\d{2}/)) mes = s.slice(0, 7);
    }
    if (!mes) continue;
    result.push({
      mes: mes,
      cap: Number(row[1]) || 0,
      int: Number(row[2]) || 0
    });
  }
  return result;
}

function getDeudaLMS(ss) {
  const ws = ss.getSheetByName(SHEET_BD_LMS);
  if (!ws) return [];
  const data = ws.getDataRange().getValues();
  const result = [];
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    let mes = '';
    if (row[0] instanceof Date) {
      mes = row[0].toISOString().slice(0, 7);
    } else if (row[0]) {
      mes = String(row[0]).slice(0, 7);
    }
    if (!mes) continue;
    result.push({
      mes: mes,
      total: Number(row[1]) || 0
    });
  }
  return result;
}

// ---- WRITE FUNCTIONS ----

function updateMovimiento(ss, rowIndex, field, value) {
  const ws = ss.getSheetByName(SHEET_MOV);
  if (!ws) return { error: 'Sheet not found' };

  const headers = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0];
  const colIdx = headers.findIndex(h => {
    const hl = String(h).trim().toLowerCase();
    if (field === 'v') return hl.includes('movimiento_orig') || hl.includes('monto');
    if (field === 'f') return hl.includes('fecha');
    return false;
  });

  if (colIdx < 0) return { error: 'Column not found for field: ' + field };

  ws.getRange(rowIndex, colIdx + 1).setValue(value);

  // If updating Movimiento_Orig, also update Movimiento (÷1000)
  if (field === 'v') {
    const movCol = headers.findIndex(h => {
      const hl = String(h).trim().toLowerCase();
      return hl === 'movimiento' || (hl.includes('movimiento') && !hl.includes('orig'));
    });
    if (movCol >= 0) {
      ws.getRange(rowIndex, movCol + 1).setValue(value / 1000);
    }
  }

  return { ok: true, row: rowIndex, field, value };
}

function addMovimientos(ss, rows) {
  const ws = ss.getSheetByName(SHEET_MOV);
  if (!ws) return { error: 'Sheet not found' };

  const headers = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0];
  let lastRow = ws.getLastRow();
  let added = 0;

  rows.forEach(mov => {
    lastRow++;
    headers.forEach((h, colIdx) => {
      const hl = String(h).trim().toLowerCase();
      let val = '';
      if (hl.includes('fecha')) val = mov.f || '';
      else if (hl.includes('estado') || hl === 'eo') val = mov.eo || 'P';
      else if (hl.includes('empresa') || hl === 'emp') val = mov.emp || 'TQN';
      else if (hl.includes('bancarizado') || hl === 'bn') val = mov.bn || 'B';
      else if (hl.includes('categoría') || hl.includes('categoria') || hl === 'cat') val = mov.cat || '';
      else if (hl.includes('tipo') || hl === 't') val = mov.t || '';
      else if (hl.includes('marco') || hl === 'm') val = mov.m || 'BAU';
      else if (hl.includes('detalle') || hl === 'd') val = mov.d || '';
      else if (hl.includes('item') || hl === 'i') val = mov.i || '';
      else if (hl.includes('entidad') || hl === 'en') val = mov.en || '';
      else if (hl.includes('movimiento_orig') || hl.includes('monto')) val = mov.v || 0;
      else if (hl.includes('movimiento') && !hl.includes('orig')) val = (mov.v || 0) / 1000;

      if (val !== '') ws.getRange(lastRow, colIdx + 1).setValue(val);
    });
    added++;
  });

  return { ok: true, added, lastRow };
}

function deleteMovimiento(ss, rowIndex) {
  const ws = ss.getSheetByName(SHEET_MOV);
  if (!ws) return { error: 'Sheet not found' };

  const headers = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0];
  const valCol = headers.findIndex(h => {
    const hl = String(h).trim().toLowerCase();
    return hl.includes('movimiento_orig') || hl.includes('monto');
  });
  if (valCol >= 0) ws.getRange(rowIndex, valCol + 1).setValue(0);

  return { ok: true, row: rowIndex };
}

function batchUpdate(ss, changes) {
  let applied = 0;
  changes.forEach(c => {
    if (c.type === 'update') {
      updateMovimiento(ss, c.rowIndex, c.field, c.value);
      applied++;
    } else if (c.type === 'add') {
      addMovimientos(ss, [c.row]);
      applied++;
    } else if (c.type === 'delete') {
      deleteMovimiento(ss, c.rowIndex);
      applied++;
    }
  });
  return { ok: true, applied };
}
