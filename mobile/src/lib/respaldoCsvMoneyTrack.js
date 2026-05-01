/**
 * Respaldo único en .csv: varias “hojas” separadas por filas marcadoras.
 * La app reconstruye el mismo objeto `data` que el JSON (luego normalizeState).
 */

export const CSV_EXPORT_MAGIC = 'MoneyTrack-CSV-Export;v1';

/** Tablas donde cada fila es un JSON en la columna `json` (objetos anidados / forma variable). */
const TABLAS_FILA_JSON = new Set([
  'tarjetasCredito',
  'extractosTarjetasHistorial',
  'pagosProgramados',
  'intencionesCompra',
  'listaSuperCompraItems',
]);

/** Secciones con filas clave/valor planas (no columna json). */
const TABLAS_PLANAS = new Set([
  'bancosDetalle',
  'plataformasDetalle',
  'gastos',
  'ingresos',
  'categorias',
  'metas',
  'contribucionesMetas',
  'bolsillos',
  'avisosGastosMovimiento',
]);

const ESCALARES_CSV = [
  'moneda',
  'limiteTarjetaCredito',
  'presupuestoMensual',
  'presupuestoDesdeFecha',
  'saldoInicialNota',
  'asistenteUmbral48h',
  'listaSuperCategoriaPreferida',
];

function joinCsvRow(cells) {
  return cells.map((c) => {
    const s = c == null ? '' : String(c);
    const limpio = s.replace(/\r\n/g, '\n').replace(/\r/g, '\n').replace(/\n/g, ' ');
    if (/[",]/.test(limpio)) return `"${limpio.replace(/"/g, '""')}"`;
    return limpio;
  }).join(',');
}

function parseCsvRow(line) {
  const out = [];
  let cur = '';
  let inQ = false;
  for (let i = 0; i < line.length; i += 1) {
    const c = line[i];
    if (inQ) {
      if (c === '"') {
        if (line[i + 1] === '"') {
          cur += '"';
          i += 1;
        } else {
          inQ = false;
        }
      } else {
        cur += c;
      }
    } else if (c === '"') {
      inQ = true;
    } else if (c === ',') {
      out.push(cur);
      cur = '';
    } else {
      cur += c;
    }
  }
  out.push(cur);
  return out;
}

function unionKeys(rows) {
  const s = new Set();
  for (const r of rows) {
    if (r && typeof r === 'object' && !Array.isArray(r)) {
      Object.keys(r).forEach((k) => s.add(k));
    }
  }
  return [...s].sort();
}

function emitSection(name, header, rows) {
  const lines = [joinCsvRow(['__SECTION__', name]), joinCsvRow(header), ...rows.map((r) => joinCsvRow(r))];
  return lines.join('\r\n');
}

/**
 * @param {object} data - mismo shape que `data` del JSON (ordenarDataParaExportacion)
 */
export function serializarDataACsv(data, exportedAt, onboardingCompletado) {
  const bloques = [
    CSV_EXPORT_MAGIC,
    joinCsvRow(['__META__', exportedAt, onboardingCompletado ? '1' : '0']),
  ];

  const d = data && typeof data === 'object' ? data : {};

  const filasScalar = [];
  for (const k of ESCALARES_CSV) {
    const v = d[k];
    if (v === undefined || v === null || v === '') continue;
    filasScalar.push([k, String(v)]);
  }
  if (filasScalar.length) {
    bloques.push(emitSection('scalars', ['key', 'value'], filasScalar));
  }

  const saldos = d.saldosCuentas && typeof d.saldosCuentas === 'object' ? d.saldosCuentas : {};
  const filasSaldo = Object.keys(saldos).map((cuenta) => [cuenta, String(saldos[cuenta] ?? '')]);
  if (filasSaldo.length) {
    bloques.push(emitSection('saldosCuentas', ['cuenta', 'valor'], filasSaldo));
  }

  for (const nombre of TABLAS_PLANAS) {
    const arr = Array.isArray(d[nombre]) ? d[nombre] : [];
    if (!arr.length) continue;
    if (TABLAS_FILA_JSON.has(nombre)) continue;
    const keys = unionKeys(arr);
    if (!keys.length) continue;
    const header = keys;
    const rows = arr.map((obj) => header.map((k) => (obj && obj[k] != null ? String(obj[k]) : '')));
    bloques.push(emitSection(nombre, header, rows));
  }

  for (const nombre of TABLAS_FILA_JSON) {
    const arr = Array.isArray(d[nombre]) ? d[nombre] : [];
    if (!arr.length) continue;
    const rows = arr.map((obj) => [JSON.stringify(obj ?? {})]);
    bloques.push(emitSection(nombre, ['json'], rows));
  }

  const rec = Array.isArray(d.recordatoriosPagoRegistrado) ? d.recordatoriosPagoRegistrado : [];
  if (rec.length) {
    bloques.push(emitSection('recordatoriosPagoRegistrado', ['clave'], rec.map((x) => [String(x)])));
  }

  const extra = Array.isArray(d.listaSuperArticulosExtra) ? d.listaSuperArticulosExtra : [];
  if (extra.length) {
    bloques.push(emitSection('listaSuperArticulosExtra', ['item'], extra.map((x) => [String(x)])));
  }

  return bloques.join('\r\n');
}

function dataVaciaImportacion() {
  return {
    moneda: '',
    saldosCuentas: {},
    bancosDetalle: [],
    limiteTarjetaCredito: 0,
    presupuestoMensual: 0,
    presupuestoDesdeFecha: '',
    ingresos: [],
    gastos: [],
    categorias: [],
    metas: [],
    contribucionesMetas: [],
    pagosProgramados: [],
    saldoInicialNota: '',
    plataformasDetalle: [],
    tarjetasCredito: [],
    extractosTarjetasHistorial: [],
    bolsillos: [],
    recordatoriosPagoRegistrado: [],
    intencionesCompra: [],
    asistenteUmbral48h: 50,
    listaSuperCategoriaPreferida: '',
    listaSuperArticulosExtra: [],
    listaSuperCompraItems: [],
    avisosGastosMovimiento: [],
  };
}

function csvBool(x) {
  if (x === true || x === false) return x;
  const s = String(x ?? '').trim().toLowerCase();
  return s === '1' || s === 'true' || s === 'sí' || s === 'si';
}

function csvNum(x) {
  if (x === '' || x == null) return undefined;
  if (typeof x === 'number' && Number.isFinite(x)) return x;
  const v = parseFloat(String(x).replace(',', '.'));
  return Number.isFinite(v) ? v : undefined;
}

function csvInt(x, def = 1) {
  if (x === '' || x == null) return def;
  const n = parseInt(String(x), 10);
  return Number.isFinite(n) ? n : def;
}

/** Tras leer CSV todo llega como string; devuelve tipos razonables para normalizeState. */
function coaccionarTiposTrasCsv(data) {
  if (!data || typeof data !== 'object') return data;

  if (Array.isArray(data.gastos)) {
    data.gastos = data.gastos.map((g) => {
      if (!g || typeof g !== 'object') return g;
      const cant = csvNum(g.cantidad);
      const cuo = g.cuotas !== undefined && g.cuotas !== '' ? csvInt(g.cuotas, 1) : undefined;
      const cuoM = csvNum(g.cuotaMensual);
      return {
        ...g,
        ...(cant !== undefined ? { cantidad: cant } : {}),
        ...(cuo !== undefined ? { cuotas: cuo } : {}),
        ...(cuoM !== undefined ? { cuotaMensual: cuoM } : {}),
        ...(g.esAbonoDeudaTarjeta !== undefined ? { esAbonoDeudaTarjeta: csvBool(g.esAbonoDeudaTarjeta) } : {}),
        ...(g.esTransferenciaBolsillo !== undefined ? { esTransferenciaBolsillo: csvBool(g.esTransferenciaBolsillo) } : {}),
        ...(g.notaListadoTicketCompleto !== undefined
          ? { notaListadoTicketCompleto: csvBool(g.notaListadoTicketCompleto) }
          : {}),
      };
    });
  }

  if (Array.isArray(data.ingresos)) {
    data.ingresos = data.ingresos.map((row) => {
      if (!row || typeof row !== 'object') return row;
      const cant = csvNum(row.cantidad);
      return {
        ...row,
        ...(cant !== undefined ? { cantidad: cant } : {}),
        ...(row.esRetiroBolsillo !== undefined ? { esRetiroBolsillo: csvBool(row.esRetiroBolsillo) } : {}),
      };
    });
  }

  if (Array.isArray(data.contribucionesMetas)) {
    data.contribucionesMetas = data.contribucionesMetas.map((row) => {
      if (!row || typeof row !== 'object') return row;
      const cant = csvNum(row.cantidad);
      return { ...row, ...(cant !== undefined ? { cantidad: cant } : {}) };
    });
  }

  if (Array.isArray(data.metas)) {
    data.metas = data.metas.map((row) => {
      if (!row || typeof row !== 'object') return row;
      const obj = csvNum(row.objetivo);
      return { ...row, ...(obj !== undefined ? { objetivo: obj } : {}) };
    });
  }

  if (Array.isArray(data.bancosDetalle)) {
    data.bancosDetalle = data.bancosDetalle.map((row) => {
      if (!row || typeof row !== 'object') return row;
      const sal = csvNum(row.saldo);
      return { ...row, ...(sal !== undefined ? { saldo: sal } : {}) };
    });
  }

  if (Array.isArray(data.plataformasDetalle)) {
    data.plataformasDetalle = data.plataformasDetalle.map((row) => {
      if (!row || typeof row !== 'object') return row;
      const sal = csvNum(row.saldo);
      return { ...row, ...(sal !== undefined ? { saldo: sal } : {}) };
    });
  }

  if (Array.isArray(data.bolsillos)) {
    data.bolsillos = data.bolsillos.map((row) => {
      if (!row || typeof row !== 'object') return row;
      const sal = csvNum(row.saldo);
      const obj = csvNum(row.objetivo);
      return {
        ...row,
        ...(sal !== undefined ? { saldo: sal } : {}),
        ...(obj !== undefined ? { objetivo: obj } : {}),
      };
    });
  }

  if (Array.isArray(data.categorias)) {
    data.categorias = data.categorias.map((row) => {
      if (!row || typeof row !== 'object') return row;
      const lim = csvNum(row.limite);
      return { ...row, ...(lim !== undefined ? { limite: lim } : {}) };
    });
  }

  if (Array.isArray(data.avisosGastosMovimiento)) {
    data.avisosGastosMovimiento = data.avisosGastosMovimiento.map((row) => {
      if (!row || typeof row !== 'object') return row;
      const ts = row.ts !== undefined && row.ts !== '' ? csvInt(row.ts, 0) : undefined;
      return { ...row, ...(ts !== undefined && ts > 0 ? { ts } : {}) };
    });
  }

  return data;
}

/**
 * @returns {{ ok: true, onboardingCompletado: boolean, data: object, exportedAt: string|null } | { ok: false, error: string }}
 */
export function parsearRespaldoCsv(texto) {
  if (!texto || typeof texto !== 'string') {
    return { ok: false, error: 'Archivo vacío o no legible.' };
  }
  const limpio = texto.replace(/^\uFEFF/, '').replace(/\r\n/g, '\n').replace(/\r/g, '\n').trim();
  const lines = limpio.split('\n').filter((ln) => ln.length > 0);
  if (!lines.length) {
    return { ok: false, error: 'El archivo CSV está vacío.' };
  }
  if (!lines[0].startsWith(CSV_EXPORT_MAGIC)) {
    return { ok: false, error: 'No es un respaldo CSV de MoneyTrack (falta cabecera MoneyTrack-CSV-Export).' };
  }

  let exportedAt = null;
  let onboardingCompletado = false;
  let i = 1;
  if (i < lines.length) {
    const metaCells = parseCsvRow(lines[i]);
    if (metaCells[0] === '__META__' && metaCells.length >= 3) {
      exportedAt = metaCells[1] || null;
      onboardingCompletado = metaCells[2] === '1' || metaCells[2] === 'true';
      i += 1;
    }
  }

  const data = dataVaciaImportacion();

  while (i < lines.length) {
    const line = lines[i];
    const headCells = parseCsvRow(line);
    if (headCells[0] !== '__SECTION__') {
      i += 1;
      continue;
    }
    const metaCells = headCells;
    const sectionName = String(metaCells[1] || '').trim();
    i += 1;
    if (i >= lines.length) break;
    const header = parseCsvRow(lines[i]).map((h) => h.trim());
    i += 1;

    const body = [];
    while (i < lines.length && parseCsvRow(lines[i])[0] !== '__SECTION__') {
      body.push(parseCsvRow(lines[i]));
      i += 1;
    }

    if (sectionName === 'scalars') {
      const ik = header.indexOf('key');
      const iv = header.indexOf('value');
      if (ik >= 0 && iv >= 0) {
        for (const row of body) {
          const k = row[ik];
          const v = row[iv];
          if (!k) continue;
          if (['limiteTarjetaCredito', 'presupuestoMensual', 'asistenteUmbral48h'].includes(k)) {
            data[k] = parseFloat(v) || 0;
          } else if (k === 'moneda' || k === 'presupuestoDesdeFecha' || k === 'saldoInicialNota' || k === 'listaSuperCategoriaPreferida') {
            data[k] = v != null ? String(v) : '';
          }
        }
      }
      continue;
    }

    if (sectionName === 'saldosCuentas') {
      const ic = header.indexOf('cuenta');
      const iv = header.indexOf('valor');
      if (ic >= 0 && iv >= 0) {
        const o = { ...data.saldosCuentas };
        for (const row of body) {
          const cuenta = row[ic];
          if (cuenta) o[cuenta] = parseFloat(row[iv]) || 0;
        }
        data.saldosCuentas = o;
      }
      continue;
    }

    if (sectionName === 'recordatoriosPagoRegistrado') {
      const ic = header.indexOf('clave');
      if (ic >= 0) {
        data.recordatoriosPagoRegistrado = body.map((row) => String(row[ic] || '').trim()).filter(Boolean);
      }
      continue;
    }

    if (sectionName === 'listaSuperArticulosExtra') {
      const ii = header.indexOf('item');
      if (ii >= 0) {
        data.listaSuperArticulosExtra = body.map((row) => String(row[ii] || '').trim()).filter(Boolean);
      }
      continue;
    }

    if (TABLAS_FILA_JSON.has(sectionName) && header.length === 1 && header[0] === 'json') {
      const arr = [];
      for (const row of body) {
        const raw = row[0] || '{}';
        try {
          arr.push(JSON.parse(raw));
        } catch {
          return { ok: false, error: `Fila JSON inválida en sección «${sectionName}».` };
        }
      }
      data[sectionName] = arr;
      continue;
    }

    if (TABLAS_PLANAS.has(sectionName) && Array.isArray(data[sectionName])) {
      const arr = [];
      for (const row of body) {
        const obj = {};
        header.forEach((hk, idx) => {
          if (!hk) return;
          obj[hk] = row[idx] !== undefined && row[idx] !== '' ? row[idx] : undefined;
        });
        Object.keys(obj).forEach((k) => {
          if (obj[k] === undefined) delete obj[k];
        });
        arr.push(obj);
      }
      data[sectionName] = arr;
    }
  }

  coaccionarTiposTrasCsv(data);

  return {
    ok: true,
    onboardingCompletado,
    data,
    exportedAt,
  };
}
