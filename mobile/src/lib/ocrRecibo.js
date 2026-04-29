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

  /** TOTAL con OCR ruidoso (espacios, 0/O): misma línea contiene la palabra y un importe */
  for (let i = lines.length - 1; i >= 0; i--) {
    const norm = normalizeLine(lines[i]);
    if (/^[\d.]+\s+kg/i.test(norm) && lines.length > 20) continue;
    if (/SUB\s*TOTAL|SUB-TOTAL/i.test(norm) && !/T\s*O\s*T\s*A\s*L|TOT\s*AL/i.test(norm)) continue;
    if (!/T\s*O\s*T\s*A\s*L|T\s*0\s*T\s*A\s*L|TOT\s*AL|TO\s*TA\s*L|TOTAL/i.test(norm)) continue;
    const nums = [
      ...norm.matchAll(
        /(\d{1,3}(?:\.\d{3})+(?:,\d{2})?|\d{1,3}(?:,\d{3})+\.\d{2}|\d+,\d{2}|\d{4,})/g
      ),
    ];
    let best = null;
    for (const mm of nums) {
      const val = parseAmount(mm[1]);
      if (val != null && val > 0) best = Math.max(best || 0, val);
    }
    if (best != null && best > 0) return best;
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

/** Normaliza comillas y símbolos raros del OCR para detectar ruido de forma estable. */
function normalizarTextoOCR(s) {
  return String(s || '')
    .replace(/\u00A0/g, ' ')
    .replace(/[\u00BB\u203A\u2039\u00AB\u201C\u201D\u2018\u2019]/g, "'")
    .replace(/\r/g, '\n');
}

/**
 * El OCR nativo confunde M/W, L/Z, etc. Corregimos patrones muy frecuentes en tickets CO
 * (no sustituye un diccionario completo).
 */
function aplicarCorreccionesOCRComunes(texto) {
  let t = String(texto || '');
  const pares = [
    [/\bWercado\b/gi, 'Mercado'],
    [/\bWercand[oó]\b/gi, 'Mercado'],
    [/\bMerculado\b/gi, 'Mercado'],
    [/\bLapatoca\b/gi, 'Zapatoa'],
    [/\bLapataca\b/gi, 'Zapatoa'],
    [/\bZapatuca\b/gi, 'Zapatoa'],
    [/\bZapataca\b/gi, 'Zapatoa'],
    [/\bM3RCADO\b/g, 'MERCADO'],
    [/\bW\s+CILANTRO\b/g, 'V CILANTRO'],
    [/\bW\s+PAPA\b/g, 'V PAPA'],
    [/\bW\s+PLATANO\b/g, 'V PLATANO'],
    [/\bW\s+CEBOLLA\b/g, 'V CEBOLLA'],
    [/\bW\s+TOMATE\b/g, 'V TOMATE'],
  ];
  for (const [re, rep] of pares) t = t.replace(re, rep);
  return t;
}

/** Texto que no parece razón social ni cabecera de ticket (OCR / pantalla / PDF ajeno). */
function lineaPareceRuidoOCR(l) {
  const s = normalizarTextoOCR(l).trim();
  if (s.length < 2) return true;
  if (/electiva|contribuci/i.test(s)) return true;
  if (/[»«›‹]/.test(s)) return true;
  if (/\bEj\s*\d+\s*:/i.test(s) || /\bEj1\b/i.test(s)) return true;
  if (/\d{1,2}\/[a-záéíóúñ]{3,}/i.test(s)) return true;
  if (/\bjaura\b/i.test(s) || /\bfact\b/i.test(s)) return true;
  if (/dis\s*-\s*jaura/i.test(s)) return true;
  if (/\.(pdf|jpg|jpeg|png)\b/i.test(s)) return true;
  const sym = (s.match(/[/»:«°*{}[\]\\|]{1,}/g) || []).length;
  if (sym >= 1 && /\d+\/[a-záéíóúñ]/i.test(s)) return true;
  if (sym >= 2 && s.length < 55) return true;
  return false;
}

function construirFechaDMY(d, m, yRaw) {
  const numY = parseInt(String(yRaw), 10);
  const year = numY < 100 ? 2000 + numY : numY;
  if (year > new Date().getFullYear() + 1 || year < 1990) return null;
  const dt = new Date(year, m - 1, d);
  return !Number.isNaN(dt.getTime()) ? dt : null;
}

/**
 * Línea tipo "Hora: … 12:12:38 PM" con ruido OCR antes de la hora (ej. "899").
 * @returns {{ h: number, mi: number, s: number } | null}
 */
function parseHoraTicketLine(line) {
  const s = String(line || '').trim();
  if (!/^(?:HORA|TIME)\b/i.test(s)) return null;
  const re =
    /^(?:HORA|TIME)\s*[:.]?\s*.*?\b(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM|A\.?M\.?|P\.?M\.?)?/i.exec(s);
  if (!re) return null;
  let h = parseInt(re[1], 10);
  const mi = parseInt(re[2], 10);
  const sec = re[3] ? parseInt(re[3], 10) : 0;
  const ap = re[4] ? String(re[4]).replace(/\./g, '').toUpperCase() : '';
  if (ap.startsWith('P') && h < 12) h += 12;
  if (ap.startsWith('A') && h === 12) h = 0;
  if (h < 0 || h > 23 || mi < 0 || mi > 59) return null;
  return { h, mi, s: sec };
}

function guessDate(text) {
  const t = String(text);
  const lineas = t.split(/\r?\n/).map((l) => l.trim());

  /** 1) Etiqueta explícita Fecha: / Date: (evita tomar la primera dd/mm/aaaa del ruido OCR). */
  let fechaSolo = null;
  let horaSolo = null;
  for (const line of lineas) {
    const mF = /^(?:FECHA|DATE)\s*[:.]?\s*(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/i.exec(line);
    if (mF) {
      const dt = construirFechaDMY(parseInt(mF[1], 10), parseInt(mF[2], 10), mF[3]);
      if (dt) fechaSolo = dt;
    }
    const parsedH = parseHoraTicketLine(line);
    if (parsedH) {
      horaSolo = parsedH;
    }
  }
  if (fechaSolo && !Number.isNaN(fechaSolo.getTime())) {
    if (horaSolo) {
      return new Date(
        fechaSolo.getFullYear(),
        fechaSolo.getMonth(),
        fechaSolo.getDate(),
        horaSolo.h,
        horaSolo.mi,
        horaSolo.s
      );
    }
    return fechaSolo;
  }

  /** 2) ISO */
  const isoT = /\b(20\d{2}-\d{2}-\d{2})[ T](\d{2}:\d{2}(?::\d{2})?)\b/.exec(t);
  if (isoT) {
    const d = new Date(`${isoT[1]}T${isoT[2]}`);
    if (!Number.isNaN(d.getTime())) return d;
  }
  const iso = /\b(20\d{2}-\d{2}-\d{2})\b/.exec(t);
  if (iso) {
    const d = new Date(iso[1] + 'T12:00:00');
    if (!Number.isNaN(d.getTime())) return d;
  }

  /** 3) dd/mm/aaaa + hora en el mismo bloque */
  const dmyh =
    /\b(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})\s+(\d{1,2}:\d{2}(?::\d{2})?)\b/.exec(t) ||
    /\b(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})\s*\|\s*(\d{1,2}:\d{2})\b/.exec(t);
  if (dmyh) {
    const day = parseInt(dmyh[1], 10);
    const month = parseInt(dmyh[2], 10);
    const y = dmyh[3];
    const numY = parseInt(String(y), 10);
    const year = numY < 100 ? 2000 + numY : numY;
    const timePart = dmyh[4];
    const [hh, mm, ss] = timePart.split(':').map((x) => parseInt(x, 10));
    if (year > new Date().getFullYear() + 1 || year < 1990) return null;
    const dt = new Date(year, month - 1, day, hh || 0, mm || 0, ss || 0);
    if (!Number.isNaN(dt.getTime())) return dt;
  }

  /** 4) Primera fecha plausible por línea (prioriza líneas con “fecha” y las primeras del ticket). */
  let bestDt = null;
  let bestRank = -1;
  const maxYear = new Date().getFullYear() + 1;
  for (let i = 0; i < Math.min(45, lineas.length); i++) {
    const line = lineas[i];
    const m = /\b(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})\b/.exec(line);
    if (!m) continue;
    const dt = construirFechaDMY(parseInt(m[1], 10), parseInt(m[2], 10), m[3]);
    if (!dt || dt.getFullYear() > maxYear) continue;
    let rank = 80 - i;
    if (/fecha|emis|impres|transac|ticket/i.test(line)) rank += 120;
    /** Línea casi solo con dd/mm/aaaa (tickets sin prefijo "Fecha:" bien leído). */
    if (/^\d{1,2}\/\d{1,2}\/\d{2,4}\s*$/.test(line) || /^\d{1,2}\/\d{1,2}\/\d{2,4}\s+\d{1,2}:\d{2}/.test(line))
      rank += 55;
    if (lineaPareceRuidoOCR(line)) rank -= 100;
    if (rank > bestRank) {
      bestRank = rank;
      bestDt = dt;
    }
  }
  /** Evita fechas aisladas de bajo rango (basura tipo 18/12/2003); años viejos exigen rango alto (p. ej. línea Fecha:). */
  if (bestDt) {
    if (bestRank < 100) bestDt = null;
    else if (bestDt.getFullYear() < 2010 && bestRank < 135) bestDt = null;
  }
  if (bestDt) return bestDt;

  return null;
}

const SKIP_MERCHANT =
  /^(FACTURA|TICKET|RECIBO|NIT|P\.?\s*NIT|TEL|MESA|CAJERO|CAJA|CEDULA|DOCUMENTO|SUCURSAL|#\s*\d|MES\d|FECHA|HORA|IVA|S\.A\.?$)/i;

/** Columnas de tabla de ítems (DESCRIPCION, CANT, PRECIO…). */
const RE_HEADER_COL = /^(DESCRIP|PRODUCTO|ITEM|ART[IÍ]CULO|CANT|CANT\.|REF|COD\.?|PRECIO|VALOR|UNID)/i;

/** Encabezados de sección de ticket (no son productos). */
const RE_SECCION_TICKET =
  /^(CEREALES|FRUTAS|VERDURAS|FRUTAS\s+Y\s+VERDURAS|LACTEOS|LÁCTEOS|CARNES|ASEO|BEBIDAS|SNACKS|PANADER|CONFIT|CONGELAD|LIMPIE|HOGAR|VARIOS|DROGUER|TEXTIL|GRANOS|POLLO|RESTAURANTE|LIMPIEZA|ABARROTES|LICORES)\b/i;

/** Evita confundir descripción de producto (sin precio en la misma línea) con nombre de tienda. */
function lineaPareceDescripcionProductoSinPrecio(l) {
  const s = String(l || '').trim();
  if (s.length < 14 || s.length > 72) return false;
  if (/plaza|centro|éxito|exito|carulla|alkosto|olimpica|d1|ara|jumbo|tienda|sucursal/i.test(s)) return false;
  const w = s.split(/\s+/).length;
  if (w < 3) return false;
  return s === s.toLowerCase();
}

function lineaEsFechaTicket(line) {
  const l = String(line || '').trim();
  if (l.length > 36) return false;
  return /\b\d{1,2}\/\d{1,2}\/\d{2,4}\b/.test(l) || /\b\d{1,2}-\d{1,2}-\d{2,4}\b/.test(l);
}

/** Precio al final de línea (COP u otros) — típico de ítems, no de cabecera de tienda. */
function lineaTienePrecioAlFinal(l) {
  const s = String(l || '').trim();
  return (
    /\s\d{1,3}(?:\.\d{3})+(?:,\d{2})?\s*$/.test(s) ||
    /\s\d+,\d{2}\s*$/.test(s) ||
    /\s\d{1,3}(?:,\d{3})+\.\d{2}\s*$/.test(s)
  );
}

/** Primera línea que ya no es cabecera (fecha, sección o producto con precio). */
function indiceFinCabeceraTicket(lines) {
  for (let i = 0; i < Math.min(16, lines.length); i++) {
    const l = lines[i].replace(/\s+/g, ' ').trim();
    if (!l) continue;
    if (lineaEsFechaTicket(l)) return i;
    if (RE_SECCION_TICKET.test(l)) return i;
    if (lineaTienePrecioAlFinal(l) && !/^TOTAL\b/i.test(l)) return i;
    if (i >= 3 && lineaPareceDescripcionProductoSinPrecio(l)) return i;
  }
  return Math.min(4, lines.length);
}

function tituloLugarTicket(s) {
  const frags = String(s || '')
    .split(/\s*-\s*/)
    .map((f) => f.trim())
    .filter(Boolean);
  const locale = 'es-419';
  return frags
    .map((frag) =>
      frag
        .split(/\s+/)
        .map((w) => {
          if (!w) return w;
          if (/^éxito$/i.test(w)) return 'Éxito';
          if (/^exito$/i.test(w)) return 'Éxito';
          const lo = w.toLocaleLowerCase(locale);
          return lo.charAt(0).toLocaleUpperCase(locale) + lo.slice(1);
        })
        .join(' ')
    )
    .join(' - ')
    .slice(0, 80);
}

function puntuacionNombreComercio(l) {
  if (lineaPareceRuidoOCR(l)) return -100;
  const s = String(l).trim();
  if (s.length < 4 || s.length > 78) return -50;
  let score = 0;
  if (/\b(S\.A\.S\.?|S\.A\.|LTDA|LTDA\.|E\.U\.|S\.L\.)\b/i.test(s)) score += 8;
  if (/\b(MERCADO|SUPERMERC|HIPERMER|DROGUER|ALKOSTO|ÉXITO|EXITO|CARULLA|OLIMPICA|ARA|JUMBO)\b/i.test(s))
    score += 5;
  const letters = (s.match(/[A-Za-zÁÉÍÓÚÑáéíóúñ]/g) || []).length;
  const upper = (s.match(/[A-ZÁÉÍÓÚÑ]/g) || []).length;
  if (letters > 5 && upper / Math.max(letters, 1) > 0.45) score += 4;
  if (/\d{7,}/.test(s)) score -= 4;
  if (/^\d/.test(s) && !/^(CR|CAL|CL|KR|CRA|AV|DG)\b/i.test(s)) score -= 3;
  if (/\bfact\b/i.test(s) && !/\b(S\.A\.|LTDA|FACTURA)\b/i.test(s)) score -= 6;
  return score;
}

function lineaPareceDireccion(l) {
  const s = String(l || '').trim();
  return /^CR\s|^CALLE\s|^CL\s|^KR\s|^CRA\s|^AV\.?\s|^DG\s|^CARRERA/i.test(s) && /\d/.test(s);
}

function lineaPareceSucursalDespuesDeRazonSocial(l, principal) {
  if (lineaPareceRuidoOCR(l)) return false;
  if (lineaEsFechaTicket(l) || /^fecha\s*:/i.test(l)) return false;
  if (lineaTienePrecioAlFinal(l)) return false;
  if (SKIP_MERCHANT.test(l)) return false;
  if (/nit\s*[.:]?\s*\d/i.test(l)) return false;
  if (lineaPareceDireccion(l)) return false;
  if (RE_SECCION_TICKET.test(l) || RE_HEADER_COL.test(l)) return false;
  const s = String(l).trim();
  if (s.length < 4 || s.length > 52) return false;
  const letters = (s.match(/[A-Za-zÁÉÍÓÚÑáéíóúñ]/g) || []).length;
  if (!letters) return false;
  const upper = (s.match(/[A-ZÁÉÍÓÚÑ]/g) || []).length;
  if (upper / letters < 0.38) return false;
  if (s.split(/\s+/).length > 8) return false;
  return puntuacionNombreComercio(s) <= puntuacionNombreComercio(principal);
}

function guessEstablishment(lines) {
  const clean = lines.map((s) => s.replace(/\s+/g, ' ').trim()).filter((s) => s.length > 0);

  let bestLine = null;
  let bestIdx = -1;
  let bestScore = -9999;
  for (let i = 0; i < Math.min(30, clean.length); i++) {
    const l = clean[i];
    if (SKIP_MERCHANT.test(l)) continue;
    if (lineaEsFechaTicket(l)) continue;
    if (/^fecha\s*:/i.test(l)) continue;
    if (lineaTienePrecioAlFinal(l)) continue;
    if (/nit\s*[.:]?\s*\d/i.test(l)) continue;
    if (RE_SECCION_TICKET.test(l)) continue;
    if (RE_HEADER_COL.test(l)) continue;
    if (lineaPareceDireccion(l)) continue;
    if (lineaPareceRuidoOCR(l)) continue;
    const sc = puntuacionNombreComercio(l);
    if (sc > bestScore) {
      bestScore = sc;
      bestLine = l;
      bestIdx = i;
    }
  }

  const ensamblar = (principal, idx) => {
    const parts = [principal];
    const next = clean[idx + 1];
    if (next && lineaPareceSucursalDespuesDeRazonSocial(next, principal)) {
      parts.push(next);
    }
    return tituloLugarTicket(parts.join(' - '));
  };

  if (bestLine != null && bestScore >= 4) {
    return ensamblar(bestLine, bestIdx);
  }

  const end = indiceFinCabeceraTicket(lines);
  const parts = [];
  for (let i = 0; i < end; i++) {
    const l = lines[i].replace(/\s+/g, ' ').trim();
    if (l.length < 2) continue;
    if (lineaPareceRuidoOCR(l)) continue;
    if (SKIP_MERCHANT.test(l)) continue;
    if (/^[\d\s.:]+$/i.test(l)) continue;
    if (/nit\s*[.:]?\s*\d/i.test(l)) continue;
    if (lineaTienePrecioAlFinal(l)) continue;
    if (RE_SECCION_TICKET.test(l)) continue;
    if (lineaPareceDireccion(l)) continue;
    parts.push(l);
  }
  if (parts.length > 0) {
    return tituloLugarTicket(parts.slice(0, 3).join(' - '));
  }

  if (bestLine != null && bestScore >= 1) {
    return ensamblar(bestLine, bestIdx);
  }

  const clean2 = lines.map((s) => s.replace(/\s+/g, ' ').trim()).filter((s) => s.length >= 3);
  for (let i = 0; i < Math.min(12, clean2.length); i++) {
    const l = clean2[i];
    if (/éxito|\bexito\b/i.test(String(l))) {
      const t = String(l).slice(0, 60).trim();
      if (t.length >= 4 && !lineaPareceRuidoOCR(t)) return tituloLugarTicket(t);
    }
  }
  for (let i = 0; i < Math.min(8, clean2.length); i++) {
    const l = clean2[i];
    if (lineaPareceRuidoOCR(l)) continue;
    if (SKIP_MERCHANT.test(l)) continue;
    if (lineaTienePrecioAlFinal(l)) continue;
    if (/^[A-ZÁÉÍÓÚÑ\s]{8,}$/i.test(l) && l.length <= 54) return tituloLugarTicket(l.slice(0, 60));
    if (/^[A-Za-zÀ-ÿ0-9 &.'-]{6,}$/.test(l) && !/\d{5,}/.test(l) && l.split(/\s+/).length <= 4) {
      return tituloLugarTicket(l.slice(0, 60));
    }
  }
  const firstOk = clean2.find((l) => !lineaPareceRuidoOCR(l));
  return firstOk ? tituloLugarTicket(firstOk.slice(0, 60)) : null;
}

const RE_FOOTER_LINE =
  /^(SUBTOTAL|SUB-TOTAL|IVA|IMPUESTO|IMPUESTOS|CAMBIO|VUELTA|GRACIAS|PAGUE|RECIBIDO|FORMAS?\s+DE\s+PAGO|DESCUENTO|PROPINA|SERVICIO|FACTURA\s*E|NIT|C\.?C\.?|TEL|CAJERO|CAJA\s*:|MEDIOS?\s+DE\s+PAGO|VALOR\s+RECIBIDO)/i;
/** TOTAL de cierre (no SUBTOTAL intermedio): cortamos lista de ítems aquí. */
const RE_TOTAL_CIERRE = /^TOTAL\b/i;

const RE_REF_GRAMAJE_DESC = /\bX\d{2,5}\s*G\b/i;

function descripcionTieneReferenciaGramos(desc) {
  return RE_REF_GRAMAJE_DESC.test(String(desc || ''));
}

function tieneLetrasProducto(s) {
  return /[A-Za-zÁÉÍÓÚÑáéíóúñ]{2,}/.test(String(s || ''));
}

function parseNumColTicket(n) {
  const v = parseAmount(String(n || '').replace(/\s/g, ''));
  return v;
}

/**
 * Una línea para la nota: nombre del producto + cantidad (CANT) con tipo de medida.
 * Decimal con coma/punto → peso kg a granel; entero → pieza; X---g en nombre → empaque por unidad.
 */
/** No es producto: base gravable, IVA, tarifas, subtotales (discriminación de impuestos). */
function descPareceResumenOFiscal(s) {
  const t = String(s || '').trim();
  const u = t.toUpperCase().replace(/\s+/g, ' ');
  if (u.length < 2) return true;
  if (/^(BASE|TARIFA|SUBTOTAL|IVA|INC|IPC|CONSUMO|EXENTO|EXENTA|DESCUENTO|PROPINA|IMPUESTO|DISCRIMINAC|GRAVABLE|GRAVAD|NO\s+AFECT|RTE\s*FTE|RETENC|ICA)\b/.test(u))
    return true;
  if (u === 'BASE' || /^BASE\s+%/i.test(t)) return true;
  if (/\bBASE\s+GRAV/i.test(u)) return true;
  if (/\bIVA\s*\d/.test(u) || /\bTARIFA\s*0?\s*%/i.test(t)) return true;
  return false;
}

/** Mínimo importe de línea para aceptar fila tipo tabla (evita ruido; 1 COP permite ítems muy baratos). */
const MIN_TOTAL_LINEA_TABLA = 1;
/** Mínimo precio unitario en columnas CANT/PRECIO/TOTAL. */
const MIN_PRECIO_UNIT_TABLA = 0.01;

function formatearLineaNotaProducto(desc, cantRaw) {
  const d = String(desc || '')
    .replace(/\s+/g, ' ')
    .replace(/^[\d.,]+\s+/, '')
    .trim();
  const c = String(cantRaw || '').trim();
  if (d.length < 2 || !tieneLetrasProducto(d)) return '';
  if (descPareceResumenOFiscal(d)) return '';
  let sufijo;
  if (!c) {
    sufijo = 'cant.: —';
  } else if (/^\d+,\d+$/.test(c) || /^\d+\.\d+$/.test(c)) {
    sufijo = `cant.: ${c} kg (peso a granel)`;
  } else if (/^\d+$/.test(c)) {
    if (descripcionTieneReferenciaGramos(d)) {
      sufijo = `cant.: ${c} u. (pieza/empaque; ref. gramaje en nombre)`;
    } else {
      sufijo = `cant.: ${c} u. (por pieza)`;
    }
  } else {
    sufijo = `cant.: ${c}`;
  }
  return `${d} — ${sufijo}`;
}

/**
 * Suma la columna TOTAL de filas tipo tabla (mismo criterio que extraerProductosTablaDetallados)
 * entre cabecera y línea TOTAL. Sirve cuando no hay línea «TOTAL» legible pero sí ítems con importe.
 */
function sumarTotalesFilasTablaTicket(lines) {
  const start = indiceFinCabeceraTicket(lines);
  let sum = 0;
  let n = 0;
  const norm = (s) => String(s || '').replace(/\s+/g, ' ').trim();
  for (let i = start; i < lines.length; i++) {
    const l = norm(lines[i]);
    if (l.length < 2) continue;
    if (RE_TOTAL_CIERRE.test(l)) break;
    if (RE_FOOTER_LINE.test(l)) continue;
    if (RE_SECCION_TICKET.test(l) || RE_HEADER_COL.test(l)) continue;
    const re4d = /^(.{2,88}?)\s+(\d+,\d+|\d+\.\d+)\s+([\d.]+)\s+([\d.]+)\s*$/.exec(l);
    const re4i = /^(.{2,88}?)\s+(\d{1,3})\s+([\d.]+)\s+([\d.]+)\s*$/.exec(l);
    let tot = null;
    if (re4d && tieneLetrasProducto(re4d[1])) {
      tot = parseNumColTicket(re4d[4]);
    } else if (re4i && tieneLetrasProducto(re4i[1])) {
      const cantN = parseInt(re4i[2], 10);
      if (cantN >= 1 && cantN <= 500) tot = parseNumColTicket(re4i[4]);
    }
    if (tot != null && tot >= MIN_TOTAL_LINEA_TABLA && tot < 100000000) {
      sum += tot;
      n++;
    }
  }
  return n > 0 ? sum : null;
}

/** Suma «total línea X» ya volcados en strings de la nota (fallback del parser). */
function inferirTotalDesdeTextosProducto(productos) {
  if (!Array.isArray(productos) || productos.length === 0) return null;
  let sum = 0;
  let n = 0;
  for (const line of productos) {
    const m = /total\s+l[ií]nea\s+([\d.,\s]+)/i.exec(String(line || ''));
    if (m) {
      const v = parseAmount(String(m[1]).replace(/\s/g, ''));
      if (v != null && v > 0) {
        sum += v;
        n++;
      }
    }
  }
  return n > 0 && sum > 0 ? sum : null;
}

/** Inicio de la tabla de ítems cuando el OCR conserva el encabezado (DESCRIPCION / CANT…). */
function indiceInicioTablaProductos(lines) {
  const norml = (s) => String(s || '').replace(/\s+/g, ' ').trim();
  for (let i = 0; i < lines.length; i++) {
    const l = norml(lines[i]);
    if (
      /DESCRIPCION|DESCRIP\w*\s+CANT|DESCRIP|^\s*CANT\.?\s*$/i.test(l) ||
      /CANT\.?\s+PRECIO|CANT\s+PRECIO|PRECIO\s+TOTAL|^\s*DESCRIP/i.test(l)
    ) {
      return i + 1;
    }
  }
  return -1;
}

/**
 * Tabla DESCRIPCION | CANT | PRECIO | TOTAL: filas completas o descripción multilínea + línea numérica.
 * `optStart`: si viene, se usa como primera línea de datos (p. ej. justo después del encabezado de columnas).
 * @returns {string[]}
 */
function extraerProductosTablaDetallados(lines, optStart = null) {
  const salida = [];
  const seen = new Set();
  let start;
  if (optStart != null && optStart >= 0) {
    start = optStart;
  } else {
    start = indiceFinCabeceraTicket(lines);
    if (start < lines.length && lineaEsFechaTicket(lines[start])) start += 1;
    for (let i = 0; i < lines.length; i++) {
      if (RE_HEADER_COL.test(lines[i])) {
        start = Math.max(start, i + 1);
        break;
      }
    }
  }

  let buf = [];
  const norm = (s) => String(s || '').replace(/\s+/g, ' ').trim();

  const mergeBuf = () =>
    buf
      .map(norm)
      .filter(Boolean)
      .join(' ')
      .replace(/\s+/g, ' ')
      .trim();

  const pushItem = (desc, cantRaw) => {
    const d = norm(desc);
    if (d.length < 3 || /^(X\d|REF|COD)/i.test(d)) return;
    if (descPareceResumenOFiscal(d)) return;
    if (!tieneLetrasProducto(d)) return;
    const key = `${d}|${cantRaw}`;
    if (seen.has(key)) return;
    seen.add(key);
    const line = formatearLineaNotaProducto(d, cantRaw);
    if (line) salida.push(line);
  };

  const puedeSerContinuacionDesc = (l) => {
    if (l.length < 2 || l.length > 76) return false;
    if (/^[\d\s.,]+$/.test(l)) return false;
    if (RE_HEADER_COL.test(l)) return false;
    if (/^TOTAL\b/i.test(l)) return false;
    if (lineaPareceDireccion(l)) return false;
    if (/^DISCRIMINAC|^\d+\s?%\s|^IVA\s|^BASE\s|^GRAVABLE/i.test(l)) return false;
    if (/^(V|C)\s+[A-ZÁÉÍÓÚÑa-záéíóúñ]/i.test(l)) return true;
    if (/^X\d{2,5}\s*G\b/i.test(l)) return true;
    if (/^[A-Za-zÁÉÍÓÚÑ0-9][A-Za-zÁÉÍÓÚÑáéíóúñ0-9\s.\-]{1,60}$/i.test(l) && !/^\d+[.,]\d+/.test(l)) return true;
    return false;
  };

  for (let i = start; i < lines.length; i++) {
    const l = norm(lines[i]);
    if (l.length < 2) continue;
    if (RE_TOTAL_CIERRE.test(l)) break;
    if (RE_SECCION_TICKET.test(l)) {
      buf = [];
      continue;
    }
    if (RE_HEADER_COL.test(l)) {
      buf = [];
      continue;
    }
    if (SKIP_MERCHANT.test(l)) continue;
    if (/^DISCRIMINAC|^\d+\s?%\s*$|^IVA\s|^BASE\s|^GRAVABLE/i.test(l)) {
      buf = [];
      continue;
    }
    if (lineaPareceDireccion(l)) {
      buf = [];
      continue;
    }
    if (lineaEsFechaTicket(l) && l.length < 32) continue;
    if (/^-\s*\d+%/i.test(l) || /^-\d{1,3}(?:\.\d{3})+(?:,\d{2})?\s*$/i.test(l)) continue;
    if (RE_FOOTER_LINE.test(l)) continue;

    /** CANT con coma o punto decimal (OCR a veces usa 0.185 en vez de 0,185). */
    const re4d =
      /^(.{2,88}?)\s+(\d+,\d+|\d+\.\d+)\s+([\d.]+)\s+([\d.]+)\s*$/.exec(l);
    if (re4d && tieneLetrasProducto(re4d[1])) {
      const tot = parseNumColTicket(re4d[4]);
      const prec = parseNumColTicket(re4d[3]);
      if (tot != null && tot >= MIN_TOTAL_LINEA_TABLA && prec != null && prec >= MIN_PRECIO_UNIT_TABLA) {
        const pref = mergeBuf();
        const desc = norm((pref ? `${pref} ` : '') + re4d[1]);
        buf = [];
        pushItem(desc, re4d[2].replace('.', ','));
        continue;
      }
    }

    const re4i = /^(.{2,88}?)\s+(\d{1,3})\s+([\d.]+)\s+([\d.]+)\s*$/.exec(l);
    if (re4i && tieneLetrasProducto(re4i[1])) {
      const cantN = parseInt(re4i[2], 10);
      const tot = parseNumColTicket(re4i[4]);
      const prec = parseNumColTicket(re4i[3]);
      if (
        cantN >= 1 &&
        cantN <= 500 &&
        tot != null &&
        tot >= MIN_TOTAL_LINEA_TABLA &&
        prec != null &&
        prec >= MIN_PRECIO_UNIT_TABLA
      ) {
        const pref = mergeBuf();
        const desc = norm((pref ? `${pref} ` : '') + re4i[1]);
        buf = [];
        pushItem(desc, re4i[2]);
        continue;
      }
    }

    const re3 = /^(\d+,\d+|\d+\.\d+|\d{1,3})\s+([\d.]+)\s+([\d.]+)\s*$/.exec(l);
    if (re3 && buf.length > 0) {
      const tot = parseNumColTicket(re3[3]);
      const prec = parseNumColTicket(re3[2]);
      if (tot != null && tot >= MIN_TOTAL_LINEA_TABLA && prec != null && prec >= MIN_PRECIO_UNIT_TABLA) {
        const cantTok = re3[1];
        if (!/,\d|\.\d/.test(cantTok)) {
          const cn = parseInt(cantTok, 10);
          if (cn < 1 || cn > 500) {
            /* sigue */
          } else {
            const desc = mergeBuf();
            buf = [];
            pushItem(desc, cantTok);
            continue;
          }
        } else {
          const desc = mergeBuf();
          buf = [];
          pushItem(desc, cantTok.replace('.', ','));
          continue;
        }
      }
    }

    if (puedeSerContinuacionDesc(l) && buf.length < 5) {
      buf.push(l);
      continue;
    }

    if (buf.length && /^\d/.test(l) && !re3) {
      buf = [];
    }
  }

  return salida.slice(0, 120);
}

/**
 * Tickets sin tabla clara: precio al final o 2–3 columnas pegadas en una línea.
 * @returns {string[]}
 */
function extraerProductosFormatoSimpleFallback(lines) {
  const items = [];
  const seen = new Set();
  let start = indiceFinCabeceraTicket(lines);
  if (start < lines.length && lineaEsFechaTicket(lines[start])) start += 1;

  for (let i = 0; i < lines.length; i++) {
    if (RE_HEADER_COL.test(lines[i])) {
      start = Math.max(start, i + 1);
      break;
    }
  }

  const tryPush = (desc, precioRaw, cantidadOpt) => {
    let descTrim = String(desc || '')
      .replace(/\s+/g, ' ')
      .replace(/^[\d.,]+\s+/, '')
      .trim();
    if (descTrim.length < 3 || descTrim.length > 88) return;
    if (/^(X\d|REF|COD)/i.test(descTrim)) return;
    if (descPareceResumenOFiscal(descTrim)) return;
    const precioVal = parseAmount(precioRaw);
    if (precioVal == null || precioVal <= 0) return;
    const cantRaw =
      cantidadOpt != null && String(cantidadOpt).trim() !== ''
        ? String(cantidadOpt).trim()
        : '';
    /** Sin CANT leída: ignorar importes ridículamente bajos (evita ruido fiscal); bajar umbral para verdulería. */
    if (!cantRaw && precioVal < 15) return;
    const key = `${descTrim}|${cantRaw}|${precioRaw}`;
    if (seen.has(key)) return;
    seen.add(key);
    const line = cantRaw
      ? formatearLineaNotaProducto(descTrim, cantRaw)
      : `${descTrim} — total línea ${String(precioRaw).trim()} (cant. no leída en OCR)`;
    if (line) items.push(line);
  };

  for (let i = start; i < lines.length; i++) {
    const l = lines[i].replace(/\s+/g, ' ').trim();
    if (l.length < 3) continue;
    if (RE_TOTAL_CIERRE.test(l)) break;
    if (RE_SECCION_TICKET.test(l)) continue;
    if (RE_HEADER_COL.test(l)) continue;
    if (SKIP_MERCHANT.test(l)) continue;
    if (/^[\d\s.,]+$/.test(l)) continue;
    if (lineaEsFechaTicket(l) && l.length < 28) continue;
    if (/^-\s*\d+%/i.test(l) || /^-\d{1,3}(?:\.\d{3})*(?:,\d{2})?\s*$/i.test(l)) continue;
    if (RE_FOOTER_LINE.test(l)) continue;

    const mCantEntera = /^(.{3,58}?)\s+(\d{1,3})\s+(\d{3,6})\s*$/.exec(l);
    if (mCantEntera && /[A-Za-zÁÉÍÓÚÑáéíóúñ]/.test(mCantEntera[1])) {
      const totalN = parseInt(mCantEntera[3], 10);
      const cantN = parseInt(mCantEntera[2], 10);
      if (totalN >= 50 && cantN >= 1 && cantN <= 999) {
        tryPush(mCantEntera[1], mCantEntera[3], String(cantN));
        continue;
      }
    }
    const mCantTotal = /^(.{3,58}?)\s+(\d+,\d+|\d+\.\d+)\s+(\d{1,6})\s*$/.exec(l);
    if (mCantTotal && /[A-Za-zÁÉÍÓÚÑáéíóúñ]/.test(mCantTotal[1])) {
      tryPush(mCantTotal[1], mCantTotal[3], mCantTotal[2].replace('.', ','));
      continue;
    }
    const mQty = /^(\d+[.,]?\d*)\s+(.{3,65}?)\s+(\d{1,3}(?:\.\d{3})*(?:,\d{2})?|\d+,\d{2}|\d{1,3}(?:,\d{3})+\.\d{2}|\d{3,6})\s*$/.exec(l);
    if (mQty) {
      tryPush(mQty[2], mQty[3], mQty[1]);
      continue;
    }
    const m1 = /^(.{3,65}?)\s+(\d{1,3}(?:\.\d{3})*(?:,\d{2})?)\s*$/.exec(l);
    if (m1) {
      tryPush(m1[1], m1[2]);
      continue;
    }
    const m2 = /^(.{3,65}?)\s+(\d+,\d{2})\s*$/.exec(l);
    if (m2) {
      tryPush(m2[1], m2[2]);
      continue;
    }
    const m3 = /^(.{3,65}?)\s+(\d{1,3}(?:,\d{3})+\.\d{2})\s*$/.exec(l);
    if (m3) {
      tryPush(m3[1], m3[2]);
      continue;
    }
    const mSmall = /^(.{4,65}?)\s+(\d{3,6})\s*$/.exec(l);
    if (mSmall && /[A-Za-zÁÉÍÓÚÑáéíóúñ]{3,}/.test(mSmall[1])) {
      const pv = parseAmount(mSmall[2]);
      if (pv != null && pv >= 25) tryPush(mSmall[1], mSmall[2]);
    }
  }

  return items.slice(0, 120);
}

/**
 * ML Kit / Vision suelen devolver cada columna en líneas distintas:
 * "V CILANTRO" → "0,185 4800 888" o descripción en 2 líneas y luego los números.
 * @returns {string[]}
 */
function extraerProductosLineasApiladas(lines) {
  const norm = (s) => String(s || '').replace(/\s+/g, ' ').trim();
  const salida = [];
  const seen = new Set();
  let descBuf = [];

  const lineaEsEncTab = (l) =>
    RE_HEADER_COL.test(l) ||
    /DESCRIPCION|DESCRIP\w*\s+CANT|^\s*CANT\.?\s*$/i.test(l) ||
    /CANT\s+PRECIO|PRECIO\s+TOTAL/i.test(l);

  const textoPareceDescProd = (l) => {
    if (l.length < 3 || l.length > 92) return false;
    if (/^[\d\s.,]+$/.test(l)) return false;
    if (descPareceResumenOFiscal(l)) return false;
    if (lineaPareceDireccion(l)) return false;
    if (SKIP_MERCHANT.test(l)) return false;
    if (/^DISCRIMINAC|^IVA\s|^BASE\s|^GRAVABLE|^\d+\s*%/i.test(l)) return false;
    if (!tieneLetrasProducto(l)) return false;
    return true;
  };

  const siguienteEsTripletNums = (idx) => {
    const next = norm(lines[idx + 1] || '');
    return /^(\d+,\d+|\d+\.\d+|\d{1,3})\s+([\d.]+)\s+([\d.]+)\s*$/.test(next);
  };

  let start = indiceInicioTablaProductos(lines);
  if (start < 0) {
    start = lines.findIndex((ln) => {
      const l = norm(ln);
      return (
        /^(V|C)\s+[A-Za-zÁÉÍÓÚÑáéíóúñ]/.test(l) ||
        /^PANELA\b/i.test(l) ||
        /^CHOCOLATE\b/i.test(l)
      );
    });
  }
  /** Carnes, abarrotes, etc. sin prefijo «V »: si la línea siguiente es CANT PRECIO TOTAL, empezamos aquí. */
  if (start < 0) {
    const cab = indiceFinCabeceraTicket(lines);
    for (let i = cab; i < Math.max(0, lines.length - 1); i++) {
      const l = norm(lines[i]);
      if (!l) continue;
      if (RE_TOTAL_CIERRE.test(l)) break;
      if (lineaEsEncTab(l) || RE_SECCION_TICKET.test(l)) continue;
      if (
        textoPareceDescProd(l) &&
        !lineaTienePrecioAlFinal(l) &&
        siguienteEsTripletNums(i)
      ) {
        start = i;
        break;
      }
    }
  }
  if (start < 0) return [];

  for (let i = start; i < lines.length; i++) {
    const l = norm(lines[i]);
    if (l.length < 2) continue;
    if (RE_TOTAL_CIERRE.test(l)) break;
    if (/^DISCRIMINAC\b|^SUBTOTAL\b/i.test(l)) break;
    if (lineaEsEncTab(l)) {
      descBuf = [];
      continue;
    }
    if (RE_FOOTER_LINE.test(l)) continue;

    const mNums = /^(\d+,\d+|\d+\.\d+|\d{1,3})\s+([\d.]+)\s+([\d.]+)\s*$/.exec(l);
    if (mNums) {
      const tot = parseNumColTicket(mNums[3]);
      const prec = parseNumColTicket(mNums[2]);
      if (
        tot != null &&
        tot >= MIN_TOTAL_LINEA_TABLA &&
        prec != null &&
        prec >= MIN_PRECIO_UNIT_TABLA &&
        descBuf.length > 0
      ) {
        const cantTok = mNums[1];
        if (!/,\d|\.\d/.test(cantTok)) {
          const cn = parseInt(cantTok, 10);
          if (cn >= 1 && cn <= 500) {
            const desc = descBuf.join(' ');
            descBuf = [];
            if (!descPareceResumenOFiscal(desc)) {
              const cantFmt =
                cantTok.includes('.') && !cantTok.includes(',') ? cantTok.replace('.', ',') : cantTok;
              const line = formatearLineaNotaProducto(desc, cantFmt);
              if (line && !seen.has(line)) {
                seen.add(line);
                salida.push(line);
              }
            }
          }
        } else {
          const desc = descBuf.join(' ');
          descBuf = [];
          if (!descPareceResumenOFiscal(desc)) {
            const cantFmt = cantTok.replace('.', ',');
            const line = formatearLineaNotaProducto(desc, cantFmt);
            if (line && !seen.has(line)) {
              seen.add(line);
              salida.push(line);
            }
          }
        }
      } else {
        descBuf = [];
      }
      continue;
    }

    if (textoPareceDescProd(l)) {
      descBuf.push(l);
      if (descBuf.length > 4) descBuf.shift();
    } else if (!/^[\d\s.,]+$/.test(l)) {
      descBuf = [];
    }
  }

  return salida.slice(0, 120);
}

/**
 * Une listas de líneas de nota sin duplicar el mismo artículo (clave = texto antes de " — ").
 * Conserva el orden del primer listado útil y, por clave, la variante más larga (más datos).
 * @param {string[][]} listas ordenadas por prioridad (primero la extracción que suele ir en orden de ticket).
 * @returns {string[]}
 */
function fusionarLineasProductoUnicas(listas) {
  function claveLinea(line) {
    const raw = String(line || '').trim();
    const part = raw.split(/\s—\s/)[0]?.trim().toLowerCase().replace(/\s+/g, ' ') || '';
    return part.slice(0, 72);
  }
  const mejorPorClave = new Map();
  for (const lista of listas) {
    if (!Array.isArray(lista)) continue;
    for (const line of lista) {
      const raw = String(line || '').trim();
      if (raw.length < 4) continue;
      const k = claveLinea(raw);
      if (k.length < 3) continue;
      const prev = mejorPorClave.get(k);
      if (!prev || raw.length > prev.length) mejorPorClave.set(k, raw);
    }
  }
  const out = [];
  const ya = new Set();
  for (const lista of listas) {
    if (!Array.isArray(lista)) continue;
    for (const line of lista) {
      const raw = String(line || '').trim();
      const k = claveLinea(raw);
      if (k.length < 3 || ya.has(k)) continue;
      ya.add(k);
      const best = mejorPorClave.get(k);
      if (best) out.push(best);
    }
  }
  return out.slice(0, 120);
}

/**
 * Productos para la nota: fusiona tabla, líneas apiladas (OCR por filas) y fallback.
 * Antes solo se usaba el primer método con resultados → faltaban ítems capturados por otro parser.
 * @returns {string[]}
 */
function extraerLineasProducto(lines) {
  const idx = indiceInicioTablaProductos(lines);
  const variantesTabla = [
    extraerProductosTablaDetallados(lines),
    idx >= 0 ? extraerProductosTablaDetallados(lines, idx) : [],
    extraerProductosTablaDetallados(lines, 0),
  ];
  const tablasOrdenadas = variantesTabla.filter((a) => a.length > 0).sort((a, b) => b.length - a.length);
  const apiladas = extraerProductosLineasApiladas(lines);
  const fallback = extraerProductosFormatoSimpleFallback(lines);
  const listas = [...tablasOrdenadas, apiladas, fallback];
  const merged = fusionarLineasProductoUnicas(listas);
  return merged.length > 0 ? merged : fallback;
}

/**
 * @returns {{ monto: number|null, establecimiento: string|null, fecha: Date|null, productos: string[] }}
 * `productos`: líneas para la nota (nombre + cant. CANT y tipo: kg granel / u. pieza / empaque con X---g).
 */
export function parseDatosTicketDesdeTexto(texto) {
  const raw = aplicarCorreccionesOCRComunes(normalizarTextoOCR(texto))
    .replace(/\r/g, '\n')
    .trim();
  const lines = raw
    .split('\n')
    .map((l) => l.trim())
    .filter((l) => l.length > 0);

  let monto = pickTotalFromLines(lines);
  if (monto == null || monto <= 0) {
    const sumTab = sumarTotalesFilasTablaTicket(lines);
    if (sumTab != null && sumTab > 0) monto = sumTab;
  }

  const fecha = guessDate(raw);
  let establecimiento = guessEstablishment(lines);
  if (establecimiento && lineaPareceRuidoOCR(establecimiento)) establecimiento = null;
  if (establecimiento) {
    establecimiento = tituloLugarTicket(aplicarCorreccionesOCRComunes(establecimiento));
  }
  const productos = extraerLineasProducto(lines);
  if (monto == null || monto <= 0) {
    const sumNote = inferirTotalDesdeTextosProducto(productos);
    if (sumNote != null && sumNote > 0) monto = sumNote;
  }

  return {
    monto: monto != null && monto > 0 ? monto : null,
    establecimiento,
    fecha,
    productos,
  };
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
