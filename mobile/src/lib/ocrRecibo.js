/**
 * Extracción de datos típicos de ticket / recibo a partir del texto OCR.
 * Diseñado para español y formatos latinos habituales.
 */

import { Platform } from 'react-native';
import { readAsStringAsync, EncodingType } from 'expo-file-system/legacy';
import { getOptionalExpoTextExtractor } from './optionalExpoTextExtractor';

/** Tesseract en navegador puede usar fetch(uri); en iOS/Android fetch(file://) suele fallar — usamos data URI vía base64. */
function mimeDesdeUri(uri) {
  const u = String(uri || '').toLowerCase();
  if (u.includes('.png')) return 'image/png';
  if (u.includes('.webp')) return 'image/webp';
  return 'image/jpeg';
}

/**
 * Acepta `file://`/`content://` o ya un `data:image/...;base64,...` (recomendado desde la cámara con `base64: true`).
 */
async function prepararImagenParaTesseract(entrada) {
  if (!entrada) return entrada;
  const s = String(entrada);
  if (s.startsWith('data:image')) return s;
  if (Platform.OS === 'web') return entrada;
  try {
    const b64 = await readAsStringAsync(entrada, { encoding: EncodingType.Base64 });
    return `data:${mimeDesdeUri(entrada)};base64,${b64}`;
  } catch (e) {
    if (typeof __DEV__ !== 'undefined' && __DEV__) {
      console.warn('[OCR] readAsStringAsync', e);
    }
    return entrada;
  }
}

function parseAmount(str) {
  if (!str) return null;
  let s = String(str).trim().replace(/\s/g, '').replace(/^[\$¢€]/, '');
  if (!s) return null;
  /** COP / varios LATAM: miles con punto, sin centavos en ticket (8.100, 27.890, 1.234.567) */
  if (/^-?\d{1,3}(?:\.\d{3})+$/.test(s)) {
    const v = parseFloat(s.replace(/\./g, '').replace(/^-/, '') || '0');
    if (Number.isFinite(v) && v >= 0) return s.startsWith('-') ? null : v;
  }
  /** Formato COP / miles con punto decimal coma: 49.900,00 o 25.990 */
  const comaDecimal = /^(\d{1,3}(\.\d{3})*)(,\d{1,4})$/.exec(s);
  if (comaDecimal) {
    const n = comaDecimal[1].replace(/\./g, '') + comaDecimal[0].slice(comaDecimal[1].length).replace(',', '.');
    const v = parseFloat(n.replace(/[^\d.]/g, ''));
    return Number.isFinite(v) ? v : null;
  }
  /** 25,990.00 o 1,234.56 (coma miles, punto decimal — típico en facturas electrónicas / US) */
  if (/^\d{1,3}(,\d{3})*\.\d{2}$/.test(s)) {
    const v = parseFloat(s.replace(/,/g, ''));
    return Number.isFinite(v) ? v : null;
  }
  /** $80.000 en texto US / software contable: miles con coma y sin decimales (80,000 → 80000). Antes del paso EU 25,90 evitaba leer 80,000 como 80 */
  if (/^-?\d{1,3}(?:,\d{3})+$/.test(s)) {
    const v = parseFloat(s.replace(/,/g, '').replace(/^-/, '') || '0');
    if (Number.isFinite(v) && v >= 0) return s.startsWith('-') ? null : v;
  }
  /** Solo coma decimal típico EU: 25,90 */
  if (/^\d+[.,]\d{1,4}$/.test(s)) {
    const v = parseFloat(s.replace(/\./g, '').replace(',', '.'));
    return Number.isFinite(v) ? v : null;
  }
  const v = parseFloat(s.replace(/[^\d.-]/g, '').replace(',', '.'));
  return Number.isFinite(v) && v >= 0 ? v : null;
}

/**
 * TOTAL de cierre (\bTOTAL\b no coincide con SUBTOTAL porque no hay límite de palabra antes de la T).
 * Incluye etiquetas típicas de caja / súper LATAM (Éxito, étc.).
 */
const RE_TOTAL_SIN_SUB = [
  /\bTOTAL\b\s*[:.]?\s*\$?\s*([\d.,\s]+)/iu,
  /(?:IMPORTE\s+TOTAL|VALOR\s+TOTAL|VALOR\s+A\s+PAGAR|TOTAL\s+A\s+PAGAR|TOTAL\s+COP|NETO\s+A\s+PAGAR|NETO\s+PAGADO|GRAN\s+TOTAL|TOTAL\s+DESCRIPCION|TOTAL\s+FACTURA|TOTAL\s+DOCUMENTO)\s*[:.\s]*\$?\s*([\d.,\s]+)/iu,
  /(?:AMOUNT|TOTAL\s+DUE)\s*[:.\s]*\$?\s*([\d.,\s]+)/iu,
  /** Factura electrónica (CO y similares): línea con total en negrita / software contable */
  /\b(?:NETO\s+PAGAR|PAGO\s+TOTAL|IMPORTE\s+A\s+PAGAR)\s*[:.\s]*\$?\s*([\d.,\s]+)/iu,
];
const RE_SUBTOTAL_O_PAGAR = [
  /(?:SUBTOTAL)\s*[:.\s]*\$?\s*([\d.,\s]+)/iu,
  /(?:PAGAR|A\s+PAGAR)\s+[:.\s]*\$?\s*([\d.,\s]+)/iu,
];

function pickTotalFromLines(lines) {
  /** Quitar artefactos de impresión / OCR que rompen \bTOTAL\b (negritas, etc.) */
  const normalizeLine = (line) =>
    String(line)
      .replace(/[*•·─–—]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();

  const tryLine = (line, patterns) => {
    const norm = normalizeLine(line);
    for (const re of patterns) {
      const m = norm.match(re);
      if (m) {
        const val = parseAmount(m[1]);
        if (val != null && val > 0) return val;
      }
    }
    return null;
  };

  for (let i = lines.length - 1; i >= 0; i--) {
    const line = lines[i];
    if (/^[\d.]+\s+kg/i.test(line) && lines.length > 20) continue;
    const v = tryLine(line, RE_TOTAL_SIN_SUB);
    if (v != null) return v;
  }
  for (let i = lines.length - 1; i >= 0; i--) {
    const line = lines[i];
    if (/^[\d.]+\s+kg/i.test(line) && lines.length > 20) continue;
    const v = tryLine(line, RE_SUBTOTAL_O_PAGAR);
    if (v != null) return v;
  }

  /** Zona inferior: COP (miles con punto), US (80,000), y $ explícito en facturas */
  const tail = lines.slice(Math.max(0, lines.length - 35)).join(' ');
  const copStyle = [...tail.matchAll(/\b\d{1,3}(?:\.\d{3})+\b/g)].map((x) => parseAmount(x[0]));
  const usInteger = [...tail.matchAll(/\b\d{1,3}(?:,\d{3})+\b/g)].map((x) => parseAmount(x[0]));
  const dol = [...tail.matchAll(/\$\s*([\d]{1,3}(?:[.,]\d{3})+(?:\.\d{2})?)/g)].map((x) => parseAmount(x[1]));
  const otros = [...tail.matchAll(/(?:\$?\s*)([\d]{1,3}(?:[.,]\d{3})*(?:[.,]\d{2}))/g)].map((x) =>
    parseAmount(x[1])
  );
  const valid = [...copStyle, ...usInteger, ...dol, ...otros].filter((n) => n != null && n > 1);
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
    const dt = new Date(year, m, d);
    if (!Number.isNaN(dt.getTime())) return dt;
  }
  return null;
}

const SKIP_MERCHANT =
  /^(FACTURA|TICKET|RECIBO|NIT|P\.?\s*NIT|TEL|MESA|CAJERO|CAJA|CEDULA|DOCUMENTO|SUCURSAL|#\s*\d|MES\d|FECHA|HORA|IVA|S\.A\.?$)/i;

function guessEstablishment(lines) {
  const clean = lines.map((s) => s.replace(/\s+/g, ' ').trim()).filter((s) => s.length >= 3);
  for (let i = 0; i < Math.min(12, clean.length); i++) {
    const l = clean[i];
    if (/éxito|\bexito\b/i.test(String(l))) {
      const t = String(l).slice(0, 60).trim();
      if (t.length >= 4) return t;
    }
  }
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
  const raw = String(texto || '')
    .replace(/\r/g, '\n')
    .trim();
  const lines = raw
    .split('\n')
    .map((l) => l.trim())
    .filter((l) => l.length > 0);

  const monto = pickTotalFromLines(lines);
  const fecha = guessDate(raw);
  const establecimiento = guessEstablishment(lines);

  return { monto: monto != null && monto > 0 ? monto : null, establecimiento, fecha };
}

/**
 * Intenta OCR nativo (ML Kit / Vision); Tesseract.js queda solo como respaldo — en React Native el WASM
 * de Tesseract suele fallar o devolver texto vacío (no es problema de “luz”).
 *
 * @param {string | { uri?: string; base64?: string }} source Ruta `file://…` y/o base64 de la captura.
 */
export async function extraerTextoDeImagen(source) {
  const uri = typeof source === 'string' ? source : source?.uri;
  const base64 =
    typeof source === 'object' && source != null && typeof source.base64 === 'string' ? source.base64 : null;

  if (Platform.OS !== 'web' && uri && (uri.startsWith('file') || uri.startsWith('content'))) {
    try {
      const ext = getOptionalExpoTextExtractor();
      /** `isSupported` puede no venir en el objeto según el interop del bundler; solo respetamos false explícito. */
      if (ext && typeof ext.extractTextFromImage === 'function' && ext.isSupported !== false) {
        const lines = await ext.extractTextFromImage(uri);
        const text = Array.isArray(lines) ? lines.join('\n').trim() : '';
        if (text.length > 0) return text;
      }
    } catch (e) {
      if (typeof __DEV__ !== 'undefined' && __DEV__) {
        console.warn('[OCR nativo ML Kit/Vision]', e?.message || e);
      }
    }
  }

  /**
   * Tesseract.js exige Web `Worker`; en Hermes/Android/iOS **no existe** (`ReferenceError: Property 'Worker' doesn't exist`).
   * Solo cargarlo en web evita ese error en consola y un bundle inútil al escanear.
   */
  if (Platform.OS !== 'web') {
    return '';
  }

  /** Web: respaldo por si no hay OCR nativo. */
  const entrada =
    base64 && base64.length > 0
      ? `data:image/jpeg;base64,${base64}`
      : uri
        ? await prepararImagenParaTesseract(uri)
        : '';
  if (!entrada) return '';

  try {
    const tesseract = await import('tesseract.js');
    const mod = tesseract.default ?? tesseract;
    const createWorker = mod.createWorker ?? tesseract.createWorker;
    const PSM = mod.PSM ?? tesseract.PSM;
    const worker = await createWorker(['spa', 'eng']);
    try {
      if (PSM?.SINGLE_BLOCK != null) {
        try {
          await worker.setParameters({ tessedit_pageseg_mode: PSM.SINGLE_BLOCK });
        } catch (_) {}
      }
      const { data } = await worker.recognize(entrada);
      const out = typeof data?.text === 'string' ? data.text.trim() : '';
      if (!out && typeof __DEV__ !== 'undefined' && __DEV__) {
        console.warn('[OCR Tesseract WASM] texto vacío; en RN use development build + expo-text-extractor.');
      }
      return out;
    } finally {
      await worker.terminate();
    }
  } catch (e) {
    if (typeof __DEV__ !== 'undefined' && __DEV__) {
      console.warn('[OCR Tesseract]', e);
    }
    return '';
  }
}
