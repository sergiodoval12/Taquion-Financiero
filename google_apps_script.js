// ============================================================
// TAQUION CF — Google Apps Script API v2
// ============================================================
// Pegar este código en: Extensiones → Apps Script del Google Sheet
// Deploy: Nueva implementación → Web App → "Cualquiera con el enlace"
// ============================================================

// Sheet names - must match exactly
const SHEET_MOV = 'BD Movimientos';
const SHEET_GC = 'Deuda GC';
const SHEET_BD = 'Deuda Bancaria';
const SHEET_NOM = 'Nómina';

// ---- GET (Read) ----
function doGet(e) {
  const action = e.parameter.action || 'all';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let result;

  try {
    if (action === 'movimientos') {
      result = getMovimientos(ss);
    } else if (action === 'deuda_gc') {
      result = getDeudaGC(ss);
    } else if (action === 'deuda_bancaria') {
      result = getDeudaBancaria(ss);
    } else if (action === 'nomina') {
      result = getNomina(ss);
    } else if (action === 'all') {
      result = {
        mov: getMovimientos(ss),
        gc: getDeudaGC(ss),
        bd: getDeudaBancaria(ss),
        nom: getNomina(ss),
        meta: { lastSync: new Date().toISOString(), source: 'Google Sheets' }
      };
    } else if (action === 'ping') {
      result = { ok: true, ts: new Date().toISOString() };
    } else if (action === 'debug') {
      const sheets = ss.getSheets().map(s => {
        const name = s.getName();
        const rows = s.getLastRow();
        const cols = s.getLastColumn();
        let headers = [];
        if (rows > 0 && cols > 0) {
          headers = s.getRange(1, 1, 1, Math.min(cols, 15)).getValues()[0].map(h => String(h).trim());
        }
        return { name, rows, cols, headers };
      });
      result = { ok: true, sheets, ts: new Date().toISOString() };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message, stack: err.stack }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---- POST (Write) ----
function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const body = JSON.parse(e.postData.contents);
  const action = body.action;

  try {
    let result;

    if (action === 'update_movimiento') {
      result = updateMovimiento(ss, body.rowIndex, body.field, body.value);
    } else if (action === 'add_movimientos') {
      result = addMovimientos(ss, body.rows);
    } else if (action === 'delete_movimiento') {
      result = deleteMovimiento(ss, body.rowIndex);
    } else if (action === 'batch_update') {
      result = batchUpdate(ss, body.changes);
    } else if (action === 'seed') {
      const results = {};
      if (body.mov) results.mov = seedMovimientos(ss, body.mov);
      if (body.gc) results.gc = { ok: true, note: 'GC sheet is managed manually' };
      if (body.bd) results.bd = { ok: true, note: 'BD sheet is managed manually' };
      result = { ok: true, seed: results };
    } else if (action === 'update_gc_estado') {
      result = updateGCEstado(ss, body.cuotaNum, body.estado);
    } else if (action === 'add_nomina') {
      result = addNomina(ss, body.rows);
    } else if (action === 'update_nomina') {
      result = updateNominaRow(ss, body.id, body.fields);
    } else if (action === 'seed_nomina') {
      result = seedNomina(ss, body.rows);
    } else {
      result = { error: 'Unknown action: ' + action };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message, stack: err.stack }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
// READ FUNCTIONS
// ============================================================

function getMovimientos(ss) {
  const ws = ss.getSheetByName(SHEET_MOV);
  if (!ws) return [];
  const lastRow = ws.getLastRow();
  const lastCol = ws.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  const data = ws.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = data[0].map(h => String(h).trim().toLowerCase());

  // Map columns by header name
  const colMap = {};
  headers.forEach((h, i) => {
    if (h === 'fecha') colMap.f = i;
    else if (h === 'estado') colMap.eo = i;
    else if (h === 'empresa') colMap.emp = i;
    else if (h === 'b/n' || h === 'bancarizado') colMap.bn = i;
    else if (h.startsWith('categor')) colMap.cat = i;
    else if (h === 'tipo') colMap.t = i;
    else if (h === 'marco') colMap.m = i;
    else if (h === 'detalle') colMap.d = i;
    else if (h === 'item') colMap.i = i;
    else if (h.startsWith('entidad') || h.startsWith('cliente')) colMap.en = i;
    else if (h.includes('monto') && h.includes('orig') || h === 'movimiento_orig') colMap.v = i;
    else if (h === 'monto (k)' || h === 'movimiento') colMap.v_div = i;
    else if (h.startsWith('forma')) colMap.fp = i;
    else if (h === 'moneda') colMap.moneda = i;
    else if (h === 'monto_usd' || h === 'monto usd') colMap.vUSD = i;
    else if (h === 'tc' || h === 'tipo_cambio' || h === 'tipo cambio') colMap.tc = i;
    else if (h === 'centro_costo' || h === 'centro de costo' || h === 'proyecto') colMap.cc = i;
  });

  const rows = [];
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    // Skip empty rows
    if (!row[colMap.f] && !row[colMap.v] && !row[colMap.v_div]) continue;

    // Parse fecha
    let fecha = '';
    if (row[colMap.f] instanceof Date) {
      fecha = Utilities.formatDate(row[colMap.f], Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else if (row[colMap.f]) {
      fecha = String(row[colMap.f]).slice(0, 10);
    }

    // Parse estado: "REAL" → "R", "PROYECTADO" → "P"
    let estado = String(row[colMap.eo] || '').trim().toUpperCase();
    if (estado === 'REAL') estado = 'R';
    else if (estado === 'PROYECTADO' || estado === 'PROY') estado = 'P';
    else if (estado !== 'R' && estado !== 'P') estado = 'P';

    // Parse value: prefer Monto Original, fallback to Monto (K) * 1000
    let valor = 0;
    if (colMap.v !== undefined && row[colMap.v] !== '' && row[colMap.v] !== null) {
      valor = Number(row[colMap.v]) || 0;
    } else if (colMap.v_div !== undefined && row[colMap.v_div] !== '' && row[colMap.v_div] !== null) {
      valor = (Number(row[colMap.v_div]) || 0) * 1000;
    }

    rows.push({
      _row: r + 1, // 1-indexed sheet row (for updates)
      f: fecha,
      eo: estado,
      emp: String(row[colMap.emp] || '').trim(),
      bn: String(row[colMap.bn] || '').trim(),
      cat: String(row[colMap.cat] || '').trim(),
      t: String(row[colMap.t] || '').trim(),
      m: String(row[colMap.m] || '').trim(),
      d: String(row[colMap.d] || '').trim(),
      i: String(row[colMap.i] || '').trim(),
      en: String(row[colMap.en] || '').trim(),
      v: valor,
      fp: colMap.fp !== undefined ? String(row[colMap.fp] || '').trim() : '',
      monOrig: colMap.moneda !== undefined ? (String(row[colMap.moneda] || '').trim().toUpperCase() || 'ARS') : 'ARS',
      vUSD: colMap.vUSD !== undefined ? (Number(row[colMap.vUSD]) || undefined) : undefined,
      tcUsado: colMap.tc !== undefined ? (Number(row[colMap.tc]) || undefined) : undefined,
      cc: colMap.cc !== undefined ? String(row[colMap.cc] || '').trim() : ''
    });
  }
  return rows;
}

function getDeudaGC(ss) {
  const ws = ss.getSheetByName(SHEET_GC);
  if (!ws) return { schedule: [], orig: {}, resumen: {} };

  const lastRow = ws.getLastRow();
  const lastCol = ws.getLastColumn();
  if (lastRow < 2) return { schedule: [], orig: {}, resumen: {} };

  const data = ws.getRange(1, 1, lastRow, Math.min(lastCol, 10)).getValues();

  // Parse original composition (rows 3-11 area)
  const orig = {};
  const resumen = {};
  const schedule = [];

  for (let r = 0; r < data.length; r++) {
    const cellA = String(data[r][0] || '').trim();
    const cellB = data[r][1];

    // Original composition
    if (cellA.includes('Credito GMC') && cellA.includes('Capital')) orig.credito_gmc_capital_usd = Number(cellB) || 0;
    else if (cellA.includes('Credito GMC') && cellA.includes('Interes')) orig.credito_gmc_intereses = Number(cellB) || 0;
    else if (cellA.includes('Inversora GJ')) orig.credito_inv_gj_usd = Number(cellB) || 0;
    else if (cellA.includes('Adelantos')) orig.adelantos_no_doc_usd = Number(cellB) || 0;
    else if (cellA.includes('TOTAL DEUDA ORIGINAL')) orig.total_deuda_usd = Number(cellB) || 0;
    else if (cellA.includes('Credito Cedido')) orig.credito_cedido_usd = Math.abs(Number(cellB) || 0);
    else if (cellA.includes('Pago Anticipado') && !cellA.includes('Cronograma')) orig.pago_anticipado_usd = Math.abs(Number(cellB) || 0);
    else if (cellA.includes('DEUDA REESTRUCTURADA')) orig.deuda_reestructurada_usd = Number(cellB) || 0;

    // Resumen
    else if (cellA.includes('Total Reestructurado')) resumen.total_usd = Number(cellB) || 0;
    else if (cellA.includes('Pagado')) {
      const match = cellA.match(/(\d+)\s*cuota/);
      resumen.pagado_usd = Number(cellB) || 0;
      if (match) resumen.cuotas_pagadas = parseInt(match[1]);
    }
    else if (cellA.includes('Pendiente')) {
      const match = cellA.match(/(\d+)\s*cuota/);
      resumen.pendiente_usd = Number(cellB) || 0;
      if (match) resumen.cuotas_pendientes = parseInt(match[1]);
    }

    // Schedule rows: detect by # column being a number
    const numCol = data[r][0];
    if (typeof numCol === 'number' && numCol >= 1 && numCol <= 50) {
      // This is a schedule row: #, Fecha, USD, ARS, Destino, Tipo, Estado
      let fecha = '';
      if (data[r][1] instanceof Date) {
        fecha = Utilities.formatDate(data[r][1], Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        fecha = String(data[r][1] || '').slice(0, 10);
      }

      const estado = String(data[r][6] || '').trim().toUpperCase();

      schedule.push({
        num: numCol,
        fecha: fecha,
        usd: Number(data[r][2]) || 0,
        ars_est: Number(data[r][3]) || 0,
        dest: String(data[r][4] || '').trim(),
        tipo: String(data[r][5] || '').trim(),
        pagada: estado === 'PAGADA'
      });
    }
  }

  resumen.cuotas_total = schedule.length;
  if (!resumen.cuotas_pagadas) resumen.cuotas_pagadas = schedule.filter(c => c.pagada).length;
  if (!resumen.cuotas_pendientes) resumen.cuotas_pendientes = schedule.filter(c => !c.pagada).length;

  return { schedule, orig, resumen };
}

function getDeudaBancaria(ss) {
  const ws = ss.getSheetByName(SHEET_BD);
  if (!ws) return { tqn: [], lms: [] };

  const lastRow = ws.getLastRow();
  const lastCol = ws.getLastColumn();
  if (lastRow < 2) return { tqn: [], lms: [] };

  const data = ws.getRange(1, 1, lastRow, Math.min(lastCol, 12)).getValues();

  const tqn = [];
  const lms = [];
  let currentSection = null;

  for (let r = 0; r < data.length; r++) {
    const cellA = String(data[r][0] || '').trim();

    // Detect section headers
    if (cellA.includes('TQN')) { currentSection = 'tqn'; continue; }
    if (cellA.includes('LMS') || cellA.includes('LUMOS')) { currentSection = 'lms'; continue; }
    if (cellA === 'Mes' || cellA === 'TOTAL' || cellA === '') continue;

    // Parse data rows: Mes, Monto Cuota, ..., Estado (col K = index 10)
    let mes = '';
    if (data[r][0] instanceof Date) {
      mes = Utilities.formatDate(data[r][0], Session.getScriptTimeZone(), 'yyyy-MM');
    } else if (cellA.match(/^\d{4}-\d{2}/)) {
      mes = cellA.slice(0, 7);
    } else {
      continue;
    }

    const monto = Number(data[r][1]) || 0;
    const estadoCol = data[r][10] !== undefined ? String(data[r][10] || '').trim().toUpperCase() : '';
    const pagada = estadoCol === 'PAGADA';

    const entry = { mes, monto, pagada };

    if (currentSection === 'lms') {
      lms.push(entry);
    } else {
      tqn.push(entry); // Default to TQN
    }
  }

  return { tqn, lms };
}

// ============================================================
// WRITE FUNCTIONS
// ============================================================

function getMovHeaders(ws) {
  const headers = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => {
    const hl = String(h).trim().toLowerCase();
    if (hl === 'fecha') map.f = i;
    else if (hl === 'estado') map.eo = i;
    else if (hl === 'empresa') map.emp = i;
    else if (hl === 'b/n' || hl === 'bancarizado') map.bn = i;
    else if (hl.startsWith('categor')) map.cat = i;
    else if (hl === 'tipo') map.t = i;
    else if (hl === 'marco') map.m = i;
    else if (hl === 'detalle') map.d = i;
    else if (hl === 'item') map.i = i;
    else if (hl.startsWith('entidad') || hl.startsWith('cliente')) map.en = i;
    else if ((hl.includes('monto') && hl.includes('orig')) || hl === 'movimiento_orig') map.v = i;
    else if (hl === 'monto (k)' || hl === 'movimiento') map.v_div = i;
    else if (hl.startsWith('forma')) map.fp = i;
    else if (hl === 'moneda') map.moneda = i;
    else if (hl === 'monto_usd' || hl === 'monto usd') map.vUSD = i;
    else if (hl === 'tc' || hl === 'tipo_cambio' || hl === 'tipo cambio') map.tc = i;
    else if (hl === 'centro_costo' || hl === 'centro de costo' || hl === 'proyecto') map.cc = i;
  });
  return { headers, map, count: headers.length };
}

function updateMovimiento(ss, rowIndex, field, value) {
  const ws = ss.getSheetByName(SHEET_MOV);
  if (!ws) return { error: 'Sheet not found: ' + SHEET_MOV };

  const { headers, map } = getMovHeaders(ws);

  if (field === 'v' || field === 'value') {
    // Update Monto Original
    if (map.v !== undefined) ws.getRange(rowIndex, map.v + 1).setValue(value);
    // Also update Monto (K)
    if (map.v_div !== undefined) ws.getRange(rowIndex, map.v_div + 1).setValue(value / 1000);
    return { ok: true, row: rowIndex, field: 'v', value };
  }

  if (field === 'f' || field === 'fecha') {
    if (map.f !== undefined) ws.getRange(rowIndex, map.f + 1).setValue(value);
    return { ok: true, row: rowIndex, field: 'f', value };
  }

  if (field === 'eo' || field === 'estado') {
    const estadoStr = value === 'R' ? 'REAL' : 'PROYECTADO';
    if (map.eo !== undefined) ws.getRange(rowIndex, map.eo + 1).setValue(estadoStr);
    return { ok: true, row: rowIndex, field: 'eo', value };
  }

  // Support all other movement fields
  const directFields = { emp: 'emp', t: 't', i: 'i', en: 'en', d: 'd', cat: 'cat', m: 'm', bn: 'bn', fp: 'fp', moneda: 'moneda', vUSD: 'vUSD', tc: 'tc', cc: 'cc' };
  if (directFields[field] !== undefined) {
    const colKey = directFields[field];
    if (map[colKey] !== undefined) {
      ws.getRange(rowIndex, map[colKey] + 1).setValue(value);
      return { ok: true, row: rowIndex, field: field, value };
    }
    return { error: 'Column not found for field: ' + field };
  }

  return { error: 'Unknown field: ' + field };
}

function addMovimientos(ss, rows) {
  const ws = ss.getSheetByName(SHEET_MOV);
  if (!ws) return { error: 'Sheet not found: ' + SHEET_MOV };

  const { map, count } = getMovHeaders(ws);
  let lastRow = ws.getLastRow();
  let added = 0;

  // Build all rows as array for batch write
  const newData = [];

  rows.forEach(mov => {
    const rowArr = new Array(count).fill('');
    if (map.f !== undefined) rowArr[map.f] = mov.f || '';
    if (map.eo !== undefined) rowArr[map.eo] = (mov.eo === 'R') ? 'REAL' : 'PROYECTADO';
    if (map.emp !== undefined) rowArr[map.emp] = mov.emp || 'TQN';
    if (map.bn !== undefined) rowArr[map.bn] = mov.bn || 'B';
    if (map.cat !== undefined) rowArr[map.cat] = mov.cat || '';
    if (map.t !== undefined) rowArr[map.t] = mov.t || '';
    if (map.m !== undefined) rowArr[map.m] = mov.m || 'BAU';
    if (map.d !== undefined) rowArr[map.d] = mov.d || '';
    if (map.i !== undefined) rowArr[map.i] = mov.i || '';
    if (map.en !== undefined) rowArr[map.en] = mov.en || '';
    if (map.v !== undefined) rowArr[map.v] = Number(mov.v) || 0;
    if (map.v_div !== undefined) rowArr[map.v_div] = (Number(mov.v) || 0) / 1000;
    if (map.fp !== undefined) rowArr[map.fp] = mov.fp || '';
    if (map.moneda !== undefined) rowArr[map.moneda] = mov.monOrig || 'ARS';
    if (map.vUSD !== undefined) rowArr[map.vUSD] = mov.vUSD || '';
    if (map.tc !== undefined) rowArr[map.tc] = mov.tcUsado || '';
    if (map.cc !== undefined) rowArr[map.cc] = mov.cc || '';

    newData.push(rowArr);
    added++;
  });

  if (newData.length > 0) {
    ws.getRange(lastRow + 1, 1, newData.length, count).setValues(newData);
  }

  return { ok: true, added, lastRow: lastRow + newData.length };
}

function deleteMovimiento(ss, rowIndex) {
  const ws = ss.getSheetByName(SHEET_MOV);
  if (!ws) return { error: 'Sheet not found' };

  const { map } = getMovHeaders(ws);
  if (map.v !== undefined) ws.getRange(rowIndex, map.v + 1).setValue(0);
  if (map.v_div !== undefined) ws.getRange(rowIndex, map.v_div + 1).setValue(0);

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

function updateGCEstado(ss, cuotaNum, estado) {
  const ws = ss.getSheetByName(SHEET_GC);
  if (!ws) return { error: 'Sheet Deuda GC not found' };

  const lastRow = ws.getLastRow();
  const data = ws.getRange(1, 1, lastRow, 7).getValues();

  for (let r = 0; r < data.length; r++) {
    if (data[r][0] === cuotaNum) {
      ws.getRange(r + 1, 7).setValue(estado); // Column G = Estado
      return { ok: true, cuota: cuotaNum, estado };
    }
  }
  return { error: 'Cuota not found: ' + cuotaNum };
}

// ============================================================
// SEED (one-time bulk load)
// ============================================================

function seedMovimientos(ss, rows) {
  let ws = ss.getSheetByName(SHEET_MOV);
  if (!ws) {
    ws = ss.insertSheet(SHEET_MOV);
  } else {
    ws.clear();
  }

  const headers = ['Fecha', 'Estado', 'Empresa', 'B/N', 'Categoría', 'Tipo', 'Marco', 'Detalle', 'Item', 'Entidad / Cliente', 'Monto (K)', 'Monto Original', 'Forma Pago'];
  const data = [headers];

  rows.forEach(m => {
    const v = Number(m.v) || 0;
    data.push([
      m.f || '',
      (m.eo === 'R') ? 'REAL' : 'PROYECTADO',
      m.emp || '',
      m.bn || '',
      m.cat || '',
      m.t || '',
      m.m || '',
      m.d || '',
      m.i || '',
      m.en || '',
      v / 1000,
      v,
      m.fp || ''
    ]);
  });

  ws.getRange(1, 1, data.length, headers.length).setValues(data);

  // Format header
  const headerRange = ws.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#1f2937');
  headerRange.setFontColor('#ffffff');
  ws.setFrozenRows(1);
  for (let c = 1; c <= headers.length; c++) ws.autoResizeColumn(c);

  return { ok: true, rows: data.length - 1, sheet: SHEET_MOV };
}

// ============================================================
// NÓMINA FUNCTIONS
// ============================================================

const NOM_HEADERS = ['ID','Nombre','Apellido','CUIT','Líder','Equipo','UN (Área)','Clasificación','Tipo','Modalidad','Empresa','Cargo','Seniority','Sueldo Bruto','Factura Monotributo','Benef Tarjeta','Benef Conect','Benef Mono','Benef OS','Sueldo Neto Total','Fecha Ingreso','Estado','Fecha Baja','Notas'];

function getNomHeaders(ws) {
  const headers = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => {
    const hl = String(h).trim().toLowerCase();
    if (hl === 'id') map.id = i;
    else if (hl === 'nombre') map.nombre = i;
    else if (hl === 'apellido') map.apellido = i;
    else if (hl === 'cuit') map.cuit = i;
    else if (hl.includes('der') || hl === 'líder' || hl === 'lider') map.lider = i;
    else if (hl === 'equipo') map.equipo = i;
    else if (hl.includes('un') || hl.includes('rea') || hl === 'un (área)') map.area = i;
    else if (hl.includes('clasif')) map.clasif = i;
    else if (hl === 'tipo') map.tipo = i;
    else if (hl.includes('modal')) map.modalidad = i;
    else if (hl === 'empresa') map.empresa = i;
    else if (hl === 'cargo') map.cargo = i;
    else if (hl.includes('senior')) map.seniority = i;
    else if (hl.includes('sueldo bruto') || hl === 'sueldo bruto') map.sueldoBruto = i;
    else if (hl.includes('factura') || hl.includes('monotributo factura')) map.facturaMono = i;
    else if (hl.includes('tarjeta')) map.benTarjeta = i;
    else if (hl.includes('conect')) map.benConect = i;
    else if (hl === 'benef mono' || hl.includes('beneficio monotributo') || hl.includes('benef mono')) map.benMono = i;
    else if (hl.includes('benef os') || hl.includes('beneficio os') || hl === 'benef os') map.benOS = i;
    else if (hl.includes('neto')) map.sueldoNeto = i;
    else if (hl.includes('ingreso')) map.fechaIngreso = i;
    else if (hl === 'estado') map.estado = i;
    else if (hl.includes('baja')) map.fechaBaja = i;
    else if (hl.includes('nota')) map.notas = i;
  });
  return { headers, map, count: headers.length };
}

function getNomina(ss) {
  const ws = ss.getSheetByName(SHEET_NOM);
  if (!ws) return [];
  const lastRow = ws.getLastRow();
  if (lastRow < 2) return [];

  const data = ws.getRange(1, 1, lastRow, ws.getLastColumn()).getValues();
  const { map } = getNomHeaders(ws);

  const rows = [];
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const id = map.id !== undefined ? (Number(row[map.id]) || r) : r;
    const nombre = map.nombre !== undefined ? String(row[map.nombre] || '').trim() : '';
    if (!nombre) continue;

    let fechaIng = '';
    if (map.fechaIngreso !== undefined && row[map.fechaIngreso]) {
      if (row[map.fechaIngreso] instanceof Date) {
        fechaIng = Utilities.formatDate(row[map.fechaIngreso], Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        fechaIng = String(row[map.fechaIngreso]).slice(0, 10);
      }
    }
    let fechaBaja = '';
    if (map.fechaBaja !== undefined && row[map.fechaBaja]) {
      if (row[map.fechaBaja] instanceof Date) {
        fechaBaja = Utilities.formatDate(row[map.fechaBaja], Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        fechaBaja = String(row[map.fechaBaja]).slice(0, 10);
      }
    }

    rows.push({
      _row: r + 1,
      id: id,
      nombre: nombre,
      apellido: map.apellido !== undefined ? String(row[map.apellido] || '').trim() : '',
      cuit: map.cuit !== undefined ? String(row[map.cuit] || '').trim() : '',
      lider: map.lider !== undefined ? String(row[map.lider] || '').trim() : '',
      equipo: map.equipo !== undefined ? String(row[map.equipo] || '').trim() : '',
      area: map.area !== undefined ? String(row[map.area] || '').trim() : '',
      clasif: map.clasif !== undefined ? String(row[map.clasif] || '').trim() : '',
      tipo: map.tipo !== undefined ? String(row[map.tipo] || '').trim() : '',
      modalidad: map.modalidad !== undefined ? String(row[map.modalidad] || '').trim() : '',
      empresa: map.empresa !== undefined ? String(row[map.empresa] || '').trim() : '',
      cargo: map.cargo !== undefined ? String(row[map.cargo] || '').trim() : '',
      seniority: map.seniority !== undefined ? String(row[map.seniority] || '').trim() : '',
      sueldoBruto: map.sueldoBruto !== undefined ? (Number(row[map.sueldoBruto]) || 0) : 0,
      facturaMono: map.facturaMono !== undefined ? (Number(row[map.facturaMono]) || 0) : 0,
      benTarjeta: map.benTarjeta !== undefined ? (Number(row[map.benTarjeta]) || 0) : 0,
      benConect: map.benConect !== undefined ? (Number(row[map.benConect]) || 0) : 0,
      benMono: map.benMono !== undefined ? (Number(row[map.benMono]) || 0) : 0,
      benOS: map.benOS !== undefined ? (Number(row[map.benOS]) || 0) : 0,
      sueldoNeto: map.sueldoNeto !== undefined ? (Number(row[map.sueldoNeto]) || 0) : 0,
      fechaIngreso: fechaIng,
      estado: map.estado !== undefined ? String(row[map.estado] || 'Activo').trim() : 'Activo',
      fechaBaja: fechaBaja,
      notas: map.notas !== undefined ? String(row[map.notas] || '').trim() : ''
    });
  }
  return rows;
}

function addNomina(ss, rows) {
  let ws = ss.getSheetByName(SHEET_NOM);
  if (!ws) {
    ws = ss.insertSheet(SHEET_NOM);
    ws.getRange(1, 1, 1, NOM_HEADERS.length).setValues([NOM_HEADERS]);
    ws.getRange(1, 1, 1, NOM_HEADERS.length).setFontWeight('bold').setBackground('#1f2937').setFontColor('#ffffff');
    ws.setFrozenRows(1);
  }

  const count = NOM_HEADERS.length;
  let lastRow = ws.getLastRow();
  const newData = [];

  rows.forEach(e => {
    const arr = new Array(count).fill('');
    arr[0] = e.id || Date.now();
    arr[1] = e.nombre || '';
    arr[2] = e.apellido || '';
    arr[3] = e.cuit || '';
    arr[4] = e.lider || '';
    arr[5] = e.equipo || '';
    arr[6] = e.area || '';
    arr[7] = e.clasif || '';
    arr[8] = e.tipo || '';
    arr[9] = e.modalidad || '';
    arr[10] = e.empresa || '';
    arr[11] = e.cargo || '';
    arr[12] = e.seniority || '';
    arr[13] = Number(e.sueldoBruto) || 0;
    arr[14] = Number(e.facturaMono) || 0;
    arr[15] = Number(e.benTarjeta) || 0;
    arr[16] = Number(e.benConect) || 0;
    arr[17] = Number(e.benMono) || 0;
    arr[18] = Number(e.benOS) || 0;
    arr[19] = Number(e.sueldoNeto) || 0;
    arr[20] = e.fechaIngreso || '';
    arr[21] = e.estado || 'Activo';
    arr[22] = e.fechaBaja || '';
    arr[23] = e.notas || '';
    newData.push(arr);
  });

  if (newData.length > 0) {
    ws.getRange(lastRow + 1, 1, newData.length, count).setValues(newData);
  }
  return { ok: true, added: newData.length };
}

function updateNominaRow(ss, id, fields) {
  const ws = ss.getSheetByName(SHEET_NOM);
  if (!ws) return { error: 'Sheet Nómina not found' };

  const lastRow = ws.getLastRow();
  const { map } = getNomHeaders(ws);
  if (map.id === undefined) return { error: 'ID column not found' };

  const ids = ws.getRange(2, map.id + 1, lastRow - 1, 1).getValues();
  let targetRow = -1;
  for (let r = 0; r < ids.length; r++) {
    if (Number(ids[r][0]) === Number(id)) { targetRow = r + 2; break; }
  }
  if (targetRow === -1) return { error: 'Employee ID not found: ' + id };

  const fieldToCol = {
    nombre: 'nombre', apellido: 'apellido', cuit: 'cuit', lider: 'lider',
    equipo: 'equipo', area: 'area', clasif: 'clasif', tipo: 'tipo',
    modalidad: 'modalidad', empresa: 'empresa', cargo: 'cargo', seniority: 'seniority',
    sueldoBruto: 'sueldoBruto', facturaMono: 'facturaMono', benTarjeta: 'benTarjeta',
    benConect: 'benConect', benMono: 'benMono', benOS: 'benOS', sueldoNeto: 'sueldoNeto',
    fechaIngreso: 'fechaIngreso', estado: 'estado', fechaBaja: 'fechaBaja', notas: 'notas'
  };

  let updated = 0;
  for (const [key, val] of Object.entries(fields)) {
    const colKey = fieldToCol[key];
    if (colKey && map[colKey] !== undefined) {
      ws.getRange(targetRow, map[colKey] + 1).setValue(val);
      updated++;
    }
  }
  return { ok: true, id, updated, row: targetRow };
}

function seedNomina(ss, rows) {
  let ws = ss.getSheetByName(SHEET_NOM);
  if (!ws) {
    ws = ss.insertSheet(SHEET_NOM);
  } else {
    ws.clear();
  }

  const data = [NOM_HEADERS];
  rows.forEach(e => {
    data.push([
      e.id || Date.now(), e.nombre||'', e.apellido||'', e.cuit||'', e.lider||'',
      e.equipo||'', e.area||'', e.clasif||'', e.tipo||'', e.modalidad||'',
      e.empresa||'', e.cargo||'', e.seniority||'',
      Number(e.sueldoBruto)||0, Number(e.facturaMono)||0, Number(e.benTarjeta)||0,
      Number(e.benConect)||0, Number(e.benMono)||0, Number(e.benOS)||0,
      Number(e.sueldoNeto)||0, e.fechaIngreso||'', e.estado||'Activo',
      e.fechaBaja||'', e.notas||''
    ]);
  });

  ws.getRange(1, 1, data.length, NOM_HEADERS.length).setValues(data);
  const headerRange = ws.getRange(1, 1, 1, NOM_HEADERS.length);
  headerRange.setFontWeight('bold').setBackground('#1f2937').setFontColor('#ffffff');
  ws.setFrozenRows(1);
  for (let c = 1; c <= NOM_HEADERS.length; c++) ws.autoResizeColumn(c);

  return { ok: true, rows: data.length - 1, sheet: SHEET_NOM };
}
