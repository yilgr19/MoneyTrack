/**
 * Extracción de datos típicos de ticket / recibo a partir del texto OCR.
 * Diseñado para español y formatos latinos habituales.
 */

function parseAmount(str) {
  if (!str) return null;
  let s = String(str).trim().replace(/\s/g, '').replace(/^[\$¢€]/, '');
  if (!s) return null;
  /** Formato COP / miles con punto decimal coma: 49.900,00 o 25.990 */
  const comaDecimal = /^(\d{1,3}(\.\d{3})*)(,\d{1,4})$/.exec(s);
  if (comaDecimal) {
    const n = comaDecimal[1].replace(/\./g, '') + comaDecimal[0].slice(comaDecimal[1].length).replace(',', '.');
    const v = parseFloat(n.replace(/[^\d.]/g, ''));
    return Number.isFinite(v) ? v : null;
  }
  /** 25,990.00 o 1,234.56 */
  if (/^\d{1,3}(,\d{3})*\.\d{2}$/.test(s)) {
    const v = parseFloat(s.replace(/,/g, ''));
    return Number.isFinite(v) ? v : null;
  }
  /** Solo coma decimal típico EU: 25,90 */
  if (/^\d+[.,]\d{1,4}$/.test(s)) {
    const v = parseFloat(s.replace(/\./g, '').replace(',', '.'));
    return Number.isFinite(v) ? v : null;
  }
  const v = parseFloat(s.replace(/[^\d.-]/g, '').replace(',', '.'));
  return Number.isFinite(v) && v >= 0 ? v : null;
}

function pickTotalFromLines(lines) {
  /** Priorizar líneas con TOTAL palabra clave, de abajo hacia arriba */
  const patterns = [
    /(?:TOTAL|SUBTOTAL|TOTAL\s+A\s+PAGAR|IMPORTE\s+TOTAL|VALOR\s+TOTAL|AMOUNT|TOTAL\s+DUE|TOTAL:)\s*[:.]?\s*\$?\s*([\d.,\s]+)/iu,
    /(?:PAGAR|A\s+PAGAR)\s+[:\.]?\s*\$?\s*([\d.,\s]+)/iu,
  ];
  for (let i = lines.length - 1; i >= 0; i--) {
    const line = lines[i];
    if (/^[\d.]+\s+kg/i.test(line) && lines.length > 20) continue;
    for (const re of patterns) {
      const m = line.match(re);
      if (m) {
        const val = parseAmount(m[1] || line.replace(/^.*TOTAL/iu, ''));
        if (val != null && val > 0) return val;
      }
    }
  }
  /** Último número que parece monto alto en zona inferior del texto */
  const tail = lines.slice(Math.max(0, lines.length - 25)).join(' ');
  const nums = [...tail.matchAll(/(?:\$?\s*)([\d]{1,3}(?:[.,]\d{3})*(?:[.,]\d{2}))/g)].map((x) =>
    parseAmount(x[1])
  );
  const valid = nums.filter((n) => n != null && n > 1);
  return valid.length ? Math.max(...valid) : null;
}

function guessDate(text) {
  const t = String(text);
  const iso = /\b(20\d{2}-\d{2}-\d{2})\b/.exec(t);
  if (iso) {
    const d = new Date(iso[1] + 'T12:00:00');
    if (!Number.isNaN(d.getTime())) return d;
  }
  const dmy =
    /\b(\d{1,2})[\/](\d{1,2})[\/](\d{2,4})\b/.exec(t) ||
    /\b(\d{1,2})-(\d{1,2})-(\d{2,4})\b/.exec(t);
  if (dmy) {
    const a = dmy[1];
    const b = dmy[2];
    const y = dmy[3];
    const numY = parseInt(String(y), 10);
    let year = numY < 100 ? 2000 + numY : numY;
    const d = parseInt(String(a), 10);
    const m = parseInt(String(b), 10) - 1;
    if (year > new Date().getFullYear() + 1 || year < 1990) return null;
    const dt = new Date(year, m, Math.min(d, 28));
    if (!Number.isNaN(dt.getTime())) return dt;
  }
  return null;
}

const SKIP_MERCHANT =
  /^(FACTURA|TICKET|RECIBO|NIT|P\.?\s*NIT|TEL|MESA|CAJERO|CAJA|CEDULA|DOCUMENTO|SUCURSAL|#\s*\d|MES\d|FECHA|HORA|IVA|S\.A\.?$)/i;

function guessEstablishment(lines) {
  const clean = lines.map((s) => s.replace(/\s+/g, ' ').trim()).filter((s) => s.length >= 3);
  for (let i = 0; i < Math.min(8, clean.length); i++) {
    const l = clean[i];
    if (SKIP_MERCHANT.test(l)) continue;
    if (/^[A-ZÁÉÍÓÚÑ\s]{8,}$/i.test(l) && l.length <= 54) return l.slice(0, 60);
    if (/^[A-Za-zÀ-ÿ0-9 &.'-]{6,}$/.test(l) && !/\d{5,}/.test(l)) return l.slice(0, 60);
  }
  return clean[0] ? clean[0].slice(0, 60) : null;
}

/**
 * @returns {{ monto: number|null, establecimiento: string|null, fecha: Date|null }}
 */
export function parseDatosTicketDesdeTexto(texto) {
  const raw = String(texto || '').replace(/\r/g, '\n');
  const lines = raw
    .split('\n')
    .map((l) => l.trim())
    .filter((l) => l.length > 0);

  const monto = pickTotalFromLines(lines);
  const fecha = guessDate(raw);
  const establecimiento = guessEstablishment(lines);

  return { monto: monto != null && monto > 0 ? monto : null, establecimiento, fecha };
}

/** Reconocimiento OCR (Tesseract WASM). Primera llamada puede descargar idiomas (requiere red). */
export async function extraerTextoDeImagen(uri) {
  try {
    const { createWorker } = await import('tesseract.js');
    const worker = await createWorker(['spa', 'eng']);
    try {
      const { data } = await worker.recognize(uri);
      return typeof data?.text === 'string' ? data.text.trim() : '';
    } finally {
      await worker.terminate();
    }
  } catch (e) {
    if (typeof __DEV__ !== 'undefined' && __DEV__) {
      console.warn('[OCR]', e);
    }
    return '';
  }
}
