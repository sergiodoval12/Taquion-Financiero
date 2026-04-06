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
const SHEET_CC = 'Catálogo CC';

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
    } else if (action === 'catalogo_cc') {
      result = getCatalogoCC(ss);
    } else if (action === 'all') {
      result = {
        mov: getMovimientos(ss),
        gc: getDeudaGC(ss),
        bd: getDeudaBancaria(ss),
        nom: getNomina(ss),
        catCC: getCatalogoCC(ss),
        meta: { lastSync: new Date().toISOString(), source: 'Google Sheets' }
      };
    } else if (action === 'ping') {
      result = { ok: true, ts: new Date().toISOString() };
    } else if (action === 'ensure_structure') {
      result = ensureStructure(ss);
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
    } else if (action === 'update_catalogo_cc') {
      result = updateCatalogoCC(ss, body.rowIndex, body.field, body.value);
    } else if (action === 'seed_catalogo_cc') {
      result = seedCatalogoCC(ss, body.rows);
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
  // FIX timezone: usar la zona horaria de la SHEET, no del SCRIPT.
  // Si no usamos esto, los Date "10/04 00:00 BA" se formatean en el TZ del script
  // (ej. America/Los_Angeles) y caen al d\u00eda anterior (off-by-one).
  const tz = ss.getSpreadsheetTimeZone() || 'America/Argentina/Buenos_Aires';

  // Map columns by header name
  // IMPORTANT: order matters — more specific matches must come BEFORE generic ones
  // Schema:
  //   en  ← "Entidad" or "Entidad / Cliente" (a quién le pagás / proveedor)
  //   cli ← "Cliente" or "Cliente Final" (para quién es el gasto)
  //   proy← "Proyecto" (proyecto del cliente)
  //   cc  ← "Centro de Costo" (uno de los 11)
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
    // Entidad must match BEFORE generic 'cliente' check
    else if (h === 'entidad' || h === 'entidad / cliente' || h === 'entidad/cliente' || h === 'proveedor') colMap.en = i;
    // Cliente Final / Cliente — destinatario del valor (no proveedor)
    else if (h === 'cliente' || h === 'cliente final' || h === 'cliente_final') colMap.cli = i;
    else if (h === 'proyecto' || h === 'proyecto / iniciativa') colMap.proy = i;
    else if (h.includes('monto') && h.includes('orig') || h === 'movimiento_orig') colMap.v = i;
    else if (h === 'monto (k)' || h === 'movimiento') colMap.v_div = i;
    else if (h.startsWith('forma')) colMap.fp = i;
    else if (h === 'moneda') colMap.moneda = i;
    else if (h === 'tc' || h === 'tipo_cambio' || h === 'tipo cambio') colMap.tc = i;
    else if (h === 'centro_costo' || h === 'centro de costo' || h === 'centro de costos') colMap.cc = i;
    else if (h === 'factura' || h === 'fac' || h === 'estado factura' || h === 'factura ok') colMap.factura = i;
  });

  const rows = [];
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    // Skip empty rows
    if (!row[colMap.f] && !row[colMap.v] && !row[colMap.v_div]) continue;

    // Parse fecha
    let fecha = '';
    if (row[colMap.f] instanceof Date) {
      fecha = Utilities.formatDate(row[colMap.f], tz, 'yyyy-MM-dd');
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
      tcUsado: colMap.tc !== undefined ? (Number(row[colMap.tc]) || undefined) : undefined,
      cc: colMap.cc !== undefined ? String(row[colMap.cc] || '').trim() : '',
      cli: colMap.cli !== undefined ? String(row[colMap.cli] || '').trim() : '',
      proy: colMap.proy !== undefined ? String(row[colMap.proy] || '').trim() : '',
      factura: colMap.factura !== undefined ? String(row[colMap.factura] || '').trim() : ''
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
  const tz = ss.getSpreadsheetTimeZone() || 'America/Argentina/Buenos_Aires';

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
        fecha = Utilities.formatDate(data[r][1], tz, 'yyyy-MM-dd');
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
  // New structure (2026+): two stacked tables in "Deuda Bancaria" sheet
  //   Table 1: CUOTAS PRESTAMOS BANCARIOS TAQUION  -> total cuota = capital + interés
  //   Table 2: SOLO CAPITAL PRESTAMOS BANCARIOS TAQUION -> solo capital
  // Each table: row0 title | row1 entidad headers | row2 CAPITAL TOMADO | row3 referencia | rowN schedule... | rowTotal
  // Returns: { loans: [...], monthly: [...], bd: [...], bdl: [] }
  //   loans: per-loan detail with full schedule
  //   monthly: aggregated {mes, cuota, capital, interes} for all loans
  //   bd: legacy array {mes, cap, int, total} for backward compat with renderCashflow etc.
  //   bdl: empty array (LMS no longer tracked separately in this sheet)
  const empty = { loans: [], monthly: [], bd: [], bdl: [], tqn: [], lms: [] };
  const ws = ss.getSheetByName(SHEET_BD);
  if (!ws) return empty;

  const lastRow = ws.getLastRow();
  const lastCol = ws.getLastColumn();
  if (lastRow < 5) return empty;

  const nCols = Math.max(lastCol, 7);
  const data = ws.getRange(1, 1, lastRow, nCols).getValues();

  // Locate table anchors
  let t1Start = -1, t2Start = -1;
  for (let r = 0; r < data.length; r++) {
    const cellA = String(data[r][0] || '').toUpperCase();
    if (t1Start === -1 && cellA.indexOf('CUOTAS PRESTAMOS') >= 0) t1Start = r;
    if (cellA.indexOf('SOLO CAPITAL') >= 0) { t2Start = r; break; }
  }
  if (t1Start === -1) return empty;

  function parseMesCell(v) {
    if (v instanceof Date) {
      // The source sheet was built with DD as year offset (e.g. day=26 -> 2026)
      const d = v.getDate();
      const m = v.getMonth() + 1;
      const year = (d >= 20 && d <= 40) ? 2000 + d : v.getFullYear();
      return year + '-' + (m < 10 ? '0' + m : m);
    }
    if (typeof v === 'number') {
      // Excel serial date
      if (v > 30000 && v < 80000) {
        const epoch = new Date(Date.UTC(1899, 11, 30));
        const dt = new Date(epoch.getTime() + v * 86400000);
        const m = dt.getUTCMonth() + 1;
        return dt.getUTCFullYear() + '-' + (m < 10 ? '0' + m : m);
      }
      return '';
    }
    const s = String(v || '').trim();
    const mx = s.match(/(\d{4})[-\/](\d{1,2})/);
    if (mx) return mx[1] + '-' + (mx[2].length < 2 ? '0' + mx[2] : mx[2]);
    // dd/mm/yyyy
    const mx2 = s.match(/(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/);
    if (mx2) return mx2[3] + '-' + (mx2[2].length < 2 ? '0' + mx2[2] : mx2[2]);
    // "Abril 2026" / "abr 2026"
    const meses = {ene:'01',feb:'02',mar:'03',abr:'04',may:'05',jun:'06',jul:'07',ago:'08',sep:'09',oct:'10',nov:'11',dic:'12'};
    const mx3 = s.toLowerCase().match(/(ene|feb|mar|abr|may|jun|jul|ago|sep|oct|nov|dic)[a-z]*\s+(\d{4})/);
    if (mx3) return mx3[2] + '-' + meses[mx3[1]];
    return '';
  }

  function parseNum(v) {
    if (typeof v === 'number') return v;
    if (v == null || v === '') return 0;
    const s = String(v).replace(/[^\d,.\-]/g, '').replace(/\./g, '').replace(',', '.');
    const n = parseFloat(s);
    return isNaN(n) ? 0 : n;
  }

  function parseTable(startRow, maxCol) {
    const entRow = data[startRow + 1] || [];
    const capRow = data[startRow + 2] || [];
    const refRow = data[startRow + 3] || [];
    const loans = [];
    const upper = Math.min(maxCol, nCols - 1);
    for (let c = 1; c <= upper; c++) {
      const ent = String(entRow[c] || '').trim();
      if (!ent) continue;
      loans.push({
        col: c,
        entidad: ent,
        capital_tomado: parseNum(capRow[c]),
        referencia: String(refRow[c] || '').trim(),
        tna: 0
      });
    }
    const schedule = [];
    let blanks = 0;
    for (let r = startRow + 4; r < data.length; r++) {
      const first = data[r][0];
      // Stop only on the next table's title rows
      if (typeof first === 'string') {
        const fu = first.toUpperCase();
        if (fu.indexOf('TOTAL') >= 0) break;
        if (fu.indexOf('CUOTAS PRESTAMOS') >= 0) break;
        if (fu.indexOf('SOLO CAPITAL') >= 0) break;
      }
      if (first == null || first === '') {
        blanks++;
        // Allow up to 4 consecutive blank rows inside the schedule before breaking
        if (blanks >= 4 && schedule.length > 0) break;
        continue;
      }
      blanks = 0;
      const mes = parseMesCell(first);
      if (!mes) continue;
      const row = { mes: mes, vals: {} };
      for (let i = 0; i < loans.length; i++) {
        row.vals[loans[i].col] = parseNum(data[r][loans[i].col]);
      }
      schedule.push(row);
    }
    return { loans: loans, schedule: schedule };
  }

  const t1 = parseTable(t1Start, 5);
  const t2 = t2Start >= 0 ? parseTable(t2Start, 5) : { loans: [], schedule: [] };

  // Pull TNA from T2 title row (cells 2..5 contain "TNA 45%" etc)
  if (t2Start >= 0) {
    const tnaRow = data[t2Start] || [];
    t1.loans.forEach(function(loan) {
      for (let c = 1; c < nCols; c++) {
        const raw = String(tnaRow[c] || '');
        const m = raw.match(/TNA\s*([\d.,]+)/i);
        if (m && c === loan.col) { loan.tna = parseFloat(m[1].replace(',', '.')) / 100; break; }
      }
    });
    // Enrich referencia with more descriptive text from T2 (row t2Start+1)
    const t2refRow = data[t2Start + 1] || [];
    t1.loans.forEach(function(loan) {
      const r2 = String(t2refRow[loan.col] || '').trim();
      if (r2 && r2.length > (loan.referencia || '').length) loan.referencia = r2;
    });
  }

  // Combine schedules per loan and aggregate monthly
  const monthlyMap = {};
  t1.loans.forEach(function(loan) {
    loan.schedule = [];
    let totCuota = 0, totCap = 0, totInt = 0;
    t1.schedule.forEach(function(row) {
      const cuota = row.vals[loan.col] || 0;
      const t2row = t2.schedule.find(function(x) { return x.mes === row.mes; });
      const capital = t2row ? (t2row.vals[loan.col] || 0) : 0;
      const interes = Math.max(0, cuota - capital);
      if (cuota > 0 || capital > 0) {
        loan.schedule.push({ mes: row.mes, cuota: cuota, capital: capital, interes: interes });
        totCuota += cuota; totCap += capital; totInt += interes;
        if (!monthlyMap[row.mes]) monthlyMap[row.mes] = { mes: row.mes, cuota: 0, capital: 0, interes: 0 };
        monthlyMap[row.mes].cuota += cuota;
        monthlyMap[row.mes].capital += capital;
        monthlyMap[row.mes].interes += interes;
      }
    });
    loan.total_cuota = totCuota;
    loan.total_capital = totCap;
    loan.total_interes = totInt;
  });

  const monthly = Object.keys(monthlyMap).sort().map(function(k) { return monthlyMap[k]; });

  // Legacy `bd` array: {mes, cap, int, total} for existing cashflow / KPI code
  const bd = monthly.map(function(m) {
    return { mes: m.mes, cap: m.capital, int: m.interes, total: m.cuota };
  });

  return { loans: t1.loans, monthly: monthly, bd: bd, bdl: [], tqn: bd, lms: [] };
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
    // Entidad / proveedor
    else if (hl === 'entidad' || hl === 'entidad / cliente' || hl === 'entidad/cliente' || hl === 'proveedor') map.en = i;
    // Cliente final (destinatario del valor)
    else if (hl === 'cliente' || hl === 'cliente final' || hl === 'cliente_final') map.cli = i;
    else if (hl === 'proyecto' || hl === 'proyecto / iniciativa') map.proy = i;
    else if ((hl.includes('monto') && hl.includes('orig')) || hl === 'movimiento_orig') map.v = i;
    else if (hl === 'monto (k)' || hl === 'movimiento') map.v_div = i;
    else if (hl.startsWith('forma')) map.fp = i;
    else if (hl === 'moneda') map.moneda = i;
    else if (hl === 'tc' || hl === 'tipo_cambio' || hl === 'tipo cambio') map.tc = i;
    else if (hl === 'centro_costo' || hl === 'centro de costo' || hl === 'centro de costos') map.cc = i;
    else if (hl === 'factura' || hl === 'fac' || hl === 'estado factura' || hl === 'factura ok') map.factura = i;
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
  const directFields = { emp: 'emp', t: 't', i: 'i', en: 'en', d: 'd', cat: 'cat', m: 'm', bn: 'bn', fp: 'fp', moneda: 'moneda', tc: 'tc', cc: 'cc', cli: 'cli', proy: 'proy', factura: 'factura' };
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
    if (map.tc !== undefined) rowArr[map.tc] = mov.tcUsado || '';
    if (map.cc !== undefined) rowArr[map.cc] = mov.cc || '';
    if (map.cli !== undefined) rowArr[map.cli] = mov.cli || '';
    if (map.proy !== undefined) rowArr[map.proy] = mov.proy || '';
    if (map.factura !== undefined) rowArr[map.factura] = mov.factura || '';

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

// Full column list for BD Movimientos
// IMPORTANT: 'Entidad' = a quién le pagás (proveedor); 'Cliente' = para quién es el gasto; 'Proyecto' = proyecto del cliente
// Note: 'Monto USD' fue removido — Monto Original + Moneda son la fuente de verdad,
// la conversión a pesos/dólares se hace en la app según la vista activa.
const MOV_ALL_HEADERS = ['Fecha','Estado','Empresa','B/N','Categoría','Tipo','Marco','Detalle','Item','Entidad','Monto (K)','Monto Original','Moneda','TC','Forma Pago','Centro de Costo','Cliente','Proyecto','Factura'];

function seedMovimientos(ss, rows) {
  let ws = ss.getSheetByName(SHEET_MOV);
  if (!ws) {
    ws = ss.insertSheet(SHEET_MOV);
  } else {
    ws.clear();
  }

  const data = [MOV_ALL_HEADERS];

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
      m.monOrig || 'ARS',
      m.tcUsado || '',
      m.fp || '',
      m.cc || '',
      m.cli || '',
      m.proy || '',
      m.factura || ''
    ]);
  });

  ws.getRange(1, 1, data.length, MOV_ALL_HEADERS.length).setValues(data);

  const headerRange = ws.getRange(1, 1, 1, MOV_ALL_HEADERS.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#1f2937');
  headerRange.setFontColor('#ffffff');
  ws.setFrozenRows(1);
  for (let c = 1; c <= MOV_ALL_HEADERS.length; c++) ws.autoResizeColumn(c);

  return { ok: true, rows: data.length - 1, sheet: SHEET_MOV };
}

// ============================================================
// ENSURE STRUCTURE — adds missing columns/sheets without losing data
// ============================================================

function ensureStructure(ss) {
  const report = { movColumns: [], nomCreated: false };

  // 1) Ensure BD Movimientos has all columns
  let wsMov = ss.getSheetByName(SHEET_MOV);
  if (wsMov) {
    const lastCol = wsMov.getLastColumn();
    const existingHeaders = lastCol > 0 ? wsMov.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim()) : [];
    const existingLower = existingHeaders.map(h => h.toLowerCase());

    MOV_ALL_HEADERS.forEach(header => {
      // Check if this header already exists (case-insensitive, with common aliases)
      const hLower = header.toLowerCase();
      const exists = existingLower.some(eh =>
        eh === hLower ||
        (hLower === 'forma pago' && eh.startsWith('forma')) ||
        (hLower === 'moneda' && eh === 'moneda') ||
        (hLower === 'tc' && (eh === 'tc' || eh === 'tipo_cambio' || eh === 'tipo cambio')) ||
        (hLower === 'centro de costo' && (eh === 'centro_costo' || eh === 'centro de costo' || eh === 'centro de costos')) ||
        // Backwards compat: existing 'Entidad / Cliente' header satisfies the new 'Entidad' column
        (hLower === 'entidad' && (eh === 'entidad' || eh === 'entidad / cliente' || eh === 'entidad/cliente' || eh === 'proveedor')) ||
        (hLower === 'cliente' && (eh === 'cliente' || eh === 'cliente final' || eh === 'cliente_final')) ||
        (hLower === 'proyecto' && eh === 'proyecto') ||
        (hLower === 'factura' && (eh === 'factura' || eh === 'fac' || eh === 'estado factura' || eh === 'factura ok'))
      );

      if (!exists) {
        const newCol = wsMov.getLastColumn() + 1;
        wsMov.getRange(1, newCol).setValue(header);
        wsMov.getRange(1, newCol).setFontWeight('bold').setBackground('#1f2937').setFontColor('#ffffff');
        report.movColumns.push(header);
      }
    });
  } else {
    // Create BD Movimientos with all headers
    wsMov = ss.insertSheet(SHEET_MOV);
    wsMov.getRange(1, 1, 1, MOV_ALL_HEADERS.length).setValues([MOV_ALL_HEADERS]);
    wsMov.getRange(1, 1, 1, MOV_ALL_HEADERS.length).setFontWeight('bold').setBackground('#1f2937').setFontColor('#ffffff');
    wsMov.setFrozenRows(1);
    report.movColumns = MOV_ALL_HEADERS;
  }

  // 2) Ensure Nómina sheet exists
  let wsNom = ss.getSheetByName(SHEET_NOM);
  if (!wsNom) {
    wsNom = ss.insertSheet(SHEET_NOM);
    wsNom.getRange(1, 1, 1, NOM_HEADERS.length).setValues([NOM_HEADERS]);
    wsNom.getRange(1, 1, 1, NOM_HEADERS.length).setFontWeight('bold').setBackground('#1f2937').setFontColor('#ffffff');
    wsNom.setFrozenRows(1);
    for (let c = 1; c <= NOM_HEADERS.length; c++) wsNom.autoResizeColumn(c);
    report.nomCreated = true;
  }

  // 3) Ensure Catálogo CC exists with the 11 cost centers + 9 industrias precargadas
  let wsCC = ss.getSheetByName(SHEET_CC);
  if (!wsCC) {
    wsCC = ss.insertSheet(SHEET_CC);
    wsCC.getRange(1, 1, 1, CC_HEADERS.length).setValues([CC_HEADERS]);
    wsCC.getRange(1, 1, 1, CC_HEADERS.length).setFontWeight('bold').setBackground('#1f2937').setFontColor('#ffffff');
    wsCC.setFrozenRows(1);
    // Precarga del borrador
    if (CC_SEED && CC_SEED.length > 0) {
      wsCC.getRange(2, 1, CC_SEED.length, CC_HEADERS.length).setValues(CC_SEED);
    }
    for (let c = 1; c <= CC_HEADERS.length; c++) wsCC.autoResizeColumn(c);
    report.ccCreated = true;
    report.ccRows = CC_SEED.length;
  } else {
    // El catálogo ya existe — chequeamos que estén las Industrias.
    // Si faltan, las agregamos sin pisar lo existente.
    const lastRow = wsCC.getLastRow();
    const existingCCs = lastRow >= 2
      ? wsCC.getRange(2, 1, lastRow - 1, 1).getValues().map(r => String(r[0] || '').trim().toLowerCase())
      : [];
    const missingRows = CC_SEED.filter(row => {
      const name = String(row[0] || '').trim().toLowerCase();
      return name && existingCCs.indexOf(name) === -1;
    });
    if (missingRows.length > 0) {
      const startRow = wsCC.getLastRow() + 1;
      wsCC.getRange(startRow, 1, missingRows.length, CC_HEADERS.length).setValues(missingRows);
      report.ccAppended = missingRows.map(r => r[0]);
    }
  }

  // 4) Ensure Deuda GC exists (just check)
  report.deudaGC = !!ss.getSheetByName(SHEET_GC);
  report.deudaBD = !!ss.getSheetByName(SHEET_BD);

  return { ok: true, report, ts: new Date().toISOString() };
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
  const tz = ss.getSpreadsheetTimeZone() || 'America/Argentina/Buenos_Aires';

  const rows = [];
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const id = map.id !== undefined ? (Number(row[map.id]) || r) : r;
    const nombre = map.nombre !== undefined ? String(row[map.nombre] || '').trim() : '';
    if (!nombre) continue;

    let fechaIng = '';
    if (map.fechaIngreso !== undefined && row[map.fechaIngreso]) {
      if (row[map.fechaIngreso] instanceof Date) {
        fechaIng = Utilities.formatDate(row[map.fechaIngreso], tz, 'yyyy-MM-dd');
      } else {
        fechaIng = String(row[map.fechaIngreso]).slice(0, 10);
      }
    }
    let fechaBaja = '';
    if (map.fechaBaja !== undefined && row[map.fechaBaja]) {
      if (row[map.fechaBaja] instanceof Date) {
        fechaBaja = Utilities.formatDate(row[map.fechaBaja], tz, 'yyyy-MM-dd');
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

// ============================================================
// CATÁLOGO CC — Manual de Centros de Costo
// ============================================================
// Sirve como fuente de verdad y "manual de entendimiento" del modelo de costos.
// Vive como pestaña aparte en el mismo Sheet, editable a mano por el CFO.
// ============================================================

const CC_HEADERS = ['Centro de Costo','Tipo','Descripción','Qué incluye','Qué NO incluye','Reglas de asignación','Ejemplos'];

// Borrador inicial — el CFO lo edita y refina con el tiempo
const CC_SEED = [
  ['Insights', 'Vendible',
   'Behavioral Science: investigación cuali y cuanti, paneles, herramientas de research y data',
   'Estudios cualitativos y cuantitativos, paneles online, herramientas tipo Digimind/Onclusive, sentiment analysis, brand tracking',
   'Producción audiovisual (Inspire), pauta paga (Ignite), fees comerciales (Cuentas)',
   'Si el output al cliente es un insight, un dato o un análisis',
   'Digimind, Onclusive, paneles cuanti, brand trackers'],
  ['Inspire', 'Vendible',
   'Lumos / Diseño / Audiovisual: producción de piezas gráficas, audiovisuales y branding',
   'Diseño gráfico, video, edición, animación, branding visual, fotografía, post-producción, ilustración',
   'Pauta paga (Ignite), research previo a la pieza (Insights)',
   'Si el deliverable al cliente es una pieza visual o audiovisual',
   'Tute Nogueira, Elena Ternogol, Julia Distefano, freelances de diseño/video'],
  ['Ignite', 'Vendible',
   'Growth / Paymedia / Prensa / Activaciones: todo lo que termina en un medio externo o evento',
   'Pauta digital y tradicional, prensa, activaciones BTL, eventos, sponsoreos, performance media',
   'Diseño de la pieza (Inspire), research previo (Insights)',
   'Si el dinero termina en un medio, en prensa o en una activación',
   'DUAL, Google Ads, Meta Ads, prensa, productoras de eventos'],
  ['Cuentas', 'Vendible',
   'AM / Negocio: gestión comercial de cuentas vivas',
   'Account managers, atención al cliente, esfuerzo de retención y crecimiento de cuenta existente',
   'Esfuerzo de venta nueva (Comercial), delivery del proyecto (otras unidades)',
   'Si es trabajo de gestión sobre una cuenta ya existente',
   'Sol Brinatti, Maria Azul Alvarez, Julian Cordoba Pivotto'],
  ['Comercial', 'Vendible',
   'MKT / Negocios: marketing propio + new business',
   'Marketing propio de Taquion, materiales comerciales, eventos comerciales propios, esfuerzo de venta nueva, identidad de marca',
   'Cuentas vivas (Cuentas), pauta de cliente (Ignite)',
   'Si es para conseguir un cliente nuevo o promocionar Taquion al mercado',
   'Diego Kupferberg, MKT propio, eventos de venta, Lorena Rinaldini'],
  ['Top Management', 'Soporte',
   'Socios / CEO / COO / CFO / CCO: el C-Suite y dirección estratégica',
   'Socios, C-Level (CEO, COO, CFO, CCO), dirección estratégica, board fees, sueldos socios',
   'Reportes a managers (van a su área respectiva)',
   'Si es un rol C-Level o la propia dirección de la compañía',
   'Sergio Doval, socios, dirección, sueldos C-Suite'],
  ['Tecnología', 'Soporte',
   'Tech: stack y herramientas tecnológicas internas',
   'SaaS, infraestructura cloud, devops, herramientas internas, hardware corporativo, dominios, hosting',
   'Software para entrega al cliente (va al CC del cliente)',
   'Si es una herramienta de uso interno transversal',
   'AWS, Notion, Slack, GitHub, Google Workspace, hardware'],
  ['Administración', 'Soporte',
   'Legales / Contabilidad / Impuestos: servicios profesionales administrativos',
   'Estudios contables, legales, abogados, gestores, impuestos no bancarios, AFIP, IIBB, Ganancias, escribanos',
   'Impuestos sobre saldos bancarios (Financiero)',
   'Si es un servicio profesional administrativo o un impuesto general',
   'Estudio EMA, DLA Piper, Yanina Ferrari, AFIP, ARBA'],
  ['Capital Humano', 'Soporte',
   'RRHH / Beneficios / Cargas sociales transversales',
   'Cargas sociales, beneficios al empleado (prepaga, OSDE, Swiss Medical), capacitación, payroll outsourcing, búsquedas',
   'Sueldos individuales (van al CC del empleado vía Nómina)',
   'Si es un servicio o beneficio transversal del equipo',
   'Swiss Medical, OSDE, Pagos Digitales, capacitaciones'],
  ['Estructura Operativa', 'Soporte',
   'Alquiler / servicios / maestranza / logística / Viáticos: costo fijo de operar',
   'Alquiler oficina, expensas, luz, gas, internet, teléfono, limpieza, maestranza, mensajería, viáticos, taxis',
   'Alquiler de equipo para un cliente puntual (va al CC del cliente)',
   'Si es un costo fijo de mantener la oficina abierta o moverse',
   'Bavio, Sucre, Edenor, Metrogas, Telecom oficina, Andreani, viáticos del equipo'],
  ['Financiero', 'Soporte',
   'Bancos / préstamos / impuestos bancarios / Deuda Privada',
   'Intereses bancarios, comisiones, impuestos sobre saldos, mantenimiento de cuenta, préstamos bancarios, Deuda Privada (ex-socio Guido Comparada, mutuos)',
   'Servicios contables (Administración)',
   'Si es un costo o ingreso vinculado a un instrumento financiero o deuda',
   'Galicia, Macro, Santander, BBVA, Ciudad, Bind, Mills, Guido Comparada'],
  // INDUSTRIAS — se usan en la columna Centro de Costo cuando el movimiento es un INGRESO
  // (sirven para taggear de dónde viene la facturación por vertical de negocio).
  ['Public Affairs', 'Industria',
   'Ingresos vinculados a Public Affairs / Asuntos Públicos',
   'Facturación de proyectos de public affairs, lobby, gobierno, asuntos regulatorios',
   'Consumo masivo (aunque sea gobierno), entretenimiento',
   'Tag de INGRESO. Usar cuando el cliente facturado está contratando por su agenda institucional',
   'Ministerios, cámaras, reguladores'],
  ['Turismo', 'Industria',
   'Ingresos vinculados a la vertical Turismo',
   'Facturación de hoteles, aerolíneas, agencias, entes de promoción turística',
   'Entretenimiento (Entertaiment y Deportes), retail de souvenirs (Consumo Masivo)',
   'Tag de INGRESO. Cliente cuyo negocio principal es el turismo',
   'Hoteles, aerolíneas, oficinas de turismo'],
  ['Banca & Fintech', 'Industria',
   'Ingresos vinculados a Banca tradicional y Fintechs',
   'Bancos, billeteras virtuales, PSPs, exchanges, lending, seguros vinculados a crédito',
   'Seguros puros (Seguros), Tecnología pura (Tecnología & AI)',
   'Tag de INGRESO. Cliente cuyo negocio está regulado por BCRA o CNV',
   'Galicia, Mercado Pago, Ualá, Lemon'],
  ['Seguros', 'Industria',
   'Ingresos vinculados a la industria aseguradora',
   'Compañías de seguros patrimoniales, vida, ART, salud',
   'Bancaseguros (Banca & Fintech), prepagas puras de salud (Consumo Masivo / Seguros según caso)',
   'Tag de INGRESO. Cliente asegurador',
   'Sancor Seguros, La Caja, Galeno, Swiss Medical (según caso)'],
  ['Consumo Masivo', 'Industria',
   'Ingresos vinculados a CPG / Retail / FMCG',
   'Alimentos, bebidas, limpieza, cuidado personal, retail masivo, ecommerce de consumo',
   'Lujo (podría ir Entretenimiento), fintech aunque venda a consumidor (Banca & Fintech)',
   'Tag de INGRESO. Cliente cuyo negocio es producto de consumo masivo',
   'Unilever, Coca-Cola, Mondelez, Arcor'],
  ['Tecnología & AI', 'Industria',
   'Ingresos vinculados a empresas tech y AI',
   'Software companies, SaaS, AI, hardware, plataformas digitales, startups tech',
   'Fintechs (Banca & Fintech), e-commerce de consumo (Consumo Masivo)',
   'Tag de INGRESO. Cliente cuyo producto core es tecnología',
   'Globant, Mercado Libre (discutir), startups SaaS'],
  ['Energía, Minería y Agua', 'Industria',
   'Ingresos vinculados a Energía, Oil & Gas, Minería, Agua, Utilities',
   'Petroleras, mineras, energéticas, renovables, distribuidoras eléctricas / gas / agua',
   'Construcción/real estate puro (Urbanismo Real State)',
   'Tag de INGRESO. Cliente del sector primario / utilities',
   'YPF, Pan American Energy, Vista, Edenor, Metrogas'],
  ['Entretenimiento y Deportes', 'Industria',
   'Ingresos vinculados a Entretenimiento, Medios, Deportes y Cultura',
   'Productoras, medios, clubes deportivos, federaciones, streaming, gaming, cultura',
   'Turismo puro (Turismo), marcas de consumo que sponsorean (Consumo Masivo)',
   'Tag de INGRESO. Cliente del mundo del entretenimiento, deportes o medios',
   'Clubes, productoras, medios, federaciones deportivas'],
  ['Urbanismo Real State', 'Industria',
   'Ingresos vinculados a Urbanismo, Real Estate, Construcción',
   'Desarrolladoras inmobiliarias, constructoras, municipios (proyecto urbano), brokers',
   'Utilities (Energía, Minería y Agua), retail en locales (Consumo Masivo)',
   'Tag de INGRESO. Cliente cuyo negocio es desarrollar ciudad / real estate',
   'Desarrolladoras, constructoras, brokers inmobiliarios']
];

// Lista plana de industrias (para validaciones y selectores en la app)
const INDUSTRIAS_LIST = [
  'Public Affairs',
  'Turismo',
  'Banca & Fintech',
  'Seguros',
  'Consumo Masivo',
  'Tecnología & AI',
  'Energía, Minería y Agua',
  'Entretenimiento y Deportes',
  'Urbanismo Real State'
];

function getCatalogoCC(ss) {
  const ws = ss.getSheetByName(SHEET_CC);
  if (!ws) return [];
  const lastRow = ws.getLastRow();
  if (lastRow < 2) return [];
  const data = ws.getRange(2, 1, lastRow - 1, CC_HEADERS.length).getValues();
  return data.map((row, idx) => ({
    _row: idx + 2,
    cc: String(row[0] || '').trim(),
    tipo: String(row[1] || '').trim(),
    desc: String(row[2] || '').trim(),
    incluye: String(row[3] || '').trim(),
    noIncluye: String(row[4] || '').trim(),
    reglas: String(row[5] || '').trim(),
    ejemplos: String(row[6] || '').trim()
  })).filter(r => r.cc);
}

function updateCatalogoCC(ss, rowIndex, field, value) {
  const ws = ss.getSheetByName(SHEET_CC);
  if (!ws) return { error: 'Sheet ' + SHEET_CC + ' not found. Run ensure_structure first.' };
  const fieldMap = { cc: 1, tipo: 2, desc: 3, incluye: 4, noIncluye: 5, reglas: 6, ejemplos: 7 };
  const col = fieldMap[field];
  if (!col) return { error: 'Unknown field: ' + field };
  ws.getRange(rowIndex, col).setValue(value);
  return { ok: true, row: rowIndex, field, value };
}

function seedCatalogoCC(ss, rows) {
  let ws = ss.getSheetByName(SHEET_CC);
  if (!ws) {
    ws = ss.insertSheet(SHEET_CC);
  } else {
    ws.clear();
  }
  ws.getRange(1, 1, 1, CC_HEADERS.length).setValues([CC_HEADERS]);
  ws.getRange(1, 1, 1, CC_HEADERS.length).setFontWeight('bold').setBackground('#1f2937').setFontColor('#ffffff');
  ws.setFrozenRows(1);
  const data = (rows && rows.length > 0 ? rows : CC_SEED);
  if (data.length > 0) {
    ws.getRange(2, 1, data.length, CC_HEADERS.length).setValues(data);
  }
  for (let c = 1; c <= CC_HEADERS.length; c++) ws.autoResizeColumn(c);
  return { ok: true, rows: data.length, sheet: SHEET_CC };
}
