/**
 * Respaldo completo para migrar entre instalaciones sin perder historial.
 *
 * Export por defecto: CSV con cabecera `MoneyTrack-CSV-Export;v1` (ver `respaldoCsvMoneyTrack.js`).
 * Import: CSV nuevo o JSON legado (`format: moneytrack-respaldo`); el picker acepta ambos.
 *
 * Share: `expo-sharing` (archivo adjunto); fallback `Share` con URI de contenido (sin `message` en Android).
 * FS: API legacy (SDK 54).
 */
import { Platform, Share } from 'react-native';
import * as Sharing from 'expo-sharing';
import * as FileSystem from 'expo-file-system/legacy';
import * as DocumentPicker from 'expo-document-picker';
import { CSV_EXPORT_MAGIC, serializarDataACsv, parsearRespaldoCsv } from './respaldoCsvMoneyTrack';

export const BACKUP_FORMAT = 'moneytrack-respaldo';
export const BACKUP_VERSION = 1;

/** Mismo orden que devuelve `normalizeState` en AppContext (export estable y legible). */
export const BACKUP_DATA_KEYS_ORDERED = [
  'moneda',
  'nombreUsuario',
  'saldosCuentas',
  'bancosDetalle',
  'limiteTarjetaCredito',
  'presupuestoMensual',
  'presupuestoDesdeFecha',
  'ingresos',
  'gastos',
  'categorias',
  'metas',
  'contribucionesMetas',
  'pagosProgramados',
  'saldoInicialNota',
  'plataformasDetalle',
  'tarjetasCredito',
  'extractosTarjetasHistorial',
  'bolsillos',
  'recordatoriosPagoRegistrado',
  'intencionesCompra',
  'asistenteUmbral48h',
  'listaSuperCategoriaPreferida',
  'listaSuperArticulosExtra',
  'listaSuperCompraItems',
  'avisosGastosMovimiento',
];

/** Notas breves por campo en `data` (interoperabilidad / migración). */
export const BACKUP_DATA_FIELD_NOTES = {
  moneda: 'Etiqueta o código de moneda principal (texto).',
  nombreUsuario: 'Nombre para personalizar avisos en campana y notificaciones locales (texto corto).',
  saldosCuentas: 'Objeto con totales por tipo: efectivo, banco, plataforma, etc.',
  bancosDetalle: 'Array de cuentas bancarias { id, nombre, saldo, … }.',
  limiteTarjetaCredito: 'Límite global legado (número); puede convivir con tarjetasCredito.',
  presupuestoMensual: 'Tope de gasto mensual.',
  presupuestoDesdeFecha: 'Fecha inicio presupuesto (YYYY-MM-DD o vacío).',
  ingresos: 'Array de ingresos registrados.',
  gastos: 'Array de gastos (incluye cuotas, origen cuenta, categoría, etc.).',
  categorias:
    'Array de categorías: nombre, color/color_hex, icono (emoji), iconoIon (Ionicons opcional), grupo503020 (necesidades|deseos|ahorro_deuda), limite.',
  metas: 'Metas de ahorro u objetivos.',
  contribucionesMetas: 'Aportes vinculados a metas.',
  pagosProgramados: 'Pagos recurrentes o programados.',
  saldoInicialNota: 'Nota libre del arranque / saldos iniciales.',
  plataformasDetalle: 'Cuentas tipo billetera/plataforma.',
  tarjetasCredito: 'Tarjetas con cortes, cupos, deuda inicial en cuotas, etc.',
  extractosTarjetasHistorial: 'Copias de estado de cuenta guardadas por periodo.',
  bolsillos: 'Sub-ahorros / bolsillos.',
  recordatoriosPagoRegistrado: 'Claves de recordatorios de pago ya cumplidos (strings).',
  intencionesCompra: 'Lista de intenciones del asistente de compras (pendientes).',
  asistenteUmbral48h: 'Umbral monetario para regla 48h del asistente.',
  listaSuperCategoriaPreferida: 'Categoría por defecto para lista súper.',
  listaSuperArticulosExtra: 'Ítems extra sugeridos (strings).',
  listaSuperCompraItems: 'Líneas de lista súper normalizadas.',
  avisosGastosMovimiento: 'Historial breve de avisos de campana por editar o quitar gastos.',
};

const ORDERED_SET = new Set(BACKUP_DATA_KEYS_ORDERED);

/**
 * Clona y reordena las claves de primer nivel de `data` para un dump estable.
 * Las claves desconocidas se añaden al final (compatibilidad hacia adelante).
 */
export function ordenarDataParaExportacion(state) {
  const clone = JSON.parse(JSON.stringify(state));
  const out = {};
  for (const k of BACKUP_DATA_KEYS_ORDERED) {
    if (Object.prototype.hasOwnProperty.call(clone, k)) {
      out[k] = clone[k];
    }
  }
  for (const k of Object.keys(clone)) {
    if (!ORDERED_SET.has(k)) {
      out[k] = clone[k];
    }
  }
  return out;
}

/**
 * @param {object} state - Estado normalizado en memoria (mismo shape que usa AppContext).
 * @param {boolean} onboardingCompletado
 */
export function crearPayloadRespaldo(state, onboardingCompletado) {
  const dataOrdenada = ordenarDataParaExportacion(state);
  return {
    format: BACKUP_FORMAT,
    version: BACKUP_VERSION,
    exportedAt: new Date().toISOString(),
    app: 'MoneyTrack',
    onboardingCompletado: !!onboardingCompletado,
    schema: {
      purpose: 'Mapa de campos en `data` para importación o migración a otra app.',
      dataKeysOrdered: [...BACKUP_DATA_KEYS_ORDERED],
      fieldNotes: { ...BACKUP_DATA_FIELD_NOTES },
      importHint: 'Pasar `data` por la misma lógica que normalizeState en AppContext antes de persistir.',
    },
    data: dataOrdenada,
  };
}

export function serializarRespaldo(payload, pretty = false) {
  return pretty ? JSON.stringify(payload, null, 2) : JSON.stringify(payload);
}

/**
 * @returns {{ ok: true, onboardingCompletado: boolean, data: object, exportedAt: string|null } | { ok: false, error: string }}
 */
export function parsearRespaldoJson(texto) {
  if (!texto || typeof texto !== 'string') {
    return { ok: false, error: 'Archivo vacío o no legible.' };
  }
  const limpio = texto.replace(/^\uFEFF/, '').trim();
  let obj;
  try {
    obj = JSON.parse(limpio);
  } catch {
    return { ok: false, error: 'No es un JSON válido.' };
  }
  if (!obj || typeof obj !== 'object') {
    return { ok: false, error: 'Contenido inválido.' };
  }
  if (obj.format !== BACKUP_FORMAT) {
    return {
      ok: false,
      error:
        'Este archivo no es un respaldo de MoneyTrack. Elige el .csv exportado desde Administrar o un .json de respaldo antiguo.',
    };
  }
  if (obj.version !== BACKUP_VERSION) {
    return {
      ok: false,
      error: `Versión de respaldo (${obj.version}) no compatible. Actualiza la app.`,
    };
  }
  if (!obj.data || typeof obj.data !== 'object') {
    return { ok: false, error: 'El respaldo no contiene la sección de datos.' };
  }
  return {
    ok: true,
    onboardingCompletado: !!obj.onboardingCompletado,
    data: obj.data,
    exportedAt: typeof obj.exportedAt === 'string' ? obj.exportedAt : null,
  };
}

function nombreArchivoRespaldoLegible() {
  const d = new Date();
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  const h = String(d.getHours()).padStart(2, '0');
  const min = String(d.getMinutes()).padStart(2, '0');
  return `MoneyTrack-respaldo-${y}-${m}-${day}_${h}${min}.csv`;
}

const SHARE_CSV_OPTS = {
  mimeType: 'text/comma-separated-values',
  dialogTitle: 'Guardar respaldo MoneyTrack',
  UTI: 'public.comma-separated-values-text',
};

/**
 * Comparte el .csv como archivo. En Android, no mezclar `message` con `url` en `Share`: WhatsApp y otras apps
 * envían solo el texto y ignoran el adjunto.
 */
async function compartirArchivoRespaldo(fileUri) {
  try {
    if (await Sharing.isAvailableAsync()) {
      await Sharing.shareAsync(fileUri, SHARE_CSV_OPTS);
      return true;
    }
  } catch {
    // Fallback nativo abajo
  }

  if (Platform.OS === 'android') {
    try {
      const contentUri = await FileSystem.getContentUriAsync(fileUri);
      await Share.share({
        title: 'Respaldo MoneyTrack (.csv)',
        url: contentUri,
      });
      return true;
    } catch {
      return false;
    }
  }

  if (Platform.OS === 'ios') {
    try {
      await Share.share({
        title: 'Respaldo MoneyTrack (.csv)',
        url: fileUri,
      });
      return true;
    } catch {
      return false;
    }
  }

  return false;
}

/**
 * Genera un archivo de respaldo y abre «Compartir» para guardarlo (Drive, Descargas, correo…). Importar solo elige ese archivo.
 */
export async function exportarRespaldoCompartir(state, onboardingCompletado) {
  if (Platform.OS === 'web') {
    return {
      ok: false,
      mensaje: 'Exportar archivo está disponible en la app para Android o iPhone.',
    };
  }
  const data = ordenarDataParaExportacion(state);
  const exportedAt = new Date().toISOString();
  const csv = serializarDataACsv(data, exportedAt, onboardingCompletado);

  const base = FileSystem.cacheDirectory || FileSystem.documentDirectory;
  if (!base) {
    return {
      ok: false,
      mensaje:
        'No hay carpeta temporal para crear el archivo en este entorno. Usa la app instalada en el teléfono (no solo la versión web).',
    };
  }

  const fileUri = `${base}${nombreArchivoRespaldoLegible()}`;
  try {
    await FileSystem.writeAsStringAsync(fileUri, csv, {
      encoding: FileSystem.EncodingType.UTF8,
    });
  } catch (e) {
    return {
      ok: false,
      mensaje: e?.message || 'No se pudo crear el archivo de respaldo.',
    };
  }

  const compartido = await compartirArchivoRespaldo(fileUri);
  if (compartido) {
    return {
      ok: true,
      mensaje:
        'Deberías ver el archivo .csv como adjunto. Elige WhatsApp, Drive, Archivos o Correo y guárdalo. En el otro teléfono: Administrar → Importar datos y abre ese .csv.',
    };
  }

  return {
    ok: false,
    mensaje:
      'No se pudo abrir el menú para compartir el archivo .csv (permisos o versión del sistema). Reintenta; si el teléfono pide acceso a archivos, acéptalo. No se envió el respaldo como texto para que puedas importar un archivo real.',
  };
}

/**
 * @returns {Promise<
 *   | { ok: true; cancelado?: false; onboardingCompletado: boolean; data: object; exportedAt: string|null }
 *   | { ok: false; cancelado: true }
 *   | { ok: false; error: string }
 * >}
 */
export async function importarRespaldoElegirArchivo() {
  if (Platform.OS === 'web') {
    return { ok: false, error: 'Importar solo en la app móvil.' };
  }
  let pick;
  try {
    pick = await DocumentPicker.getDocumentAsync({
      type: [
        'text/csv',
        'text/comma-separated-values',
        'application/csv',
        'application/json',
        'text/json',
        'text/plain',
        'application/octet-stream',
        'application/x-json',
        '*/*',
      ],
      copyToCacheDirectory: true,
    });
  } catch {
    return {
      ok: false,
      error:
        'No está disponible el selector de archivos. Vuelve a generar la app con `npx expo prebuild` y `npx expo run:android` (o tu build EAS) para incluir módulos Expo.',
    };
  }
  if (pick.canceled) {
    return { ok: false, cancelado: true };
  }
  const asset = pick.assets?.[0];
  if (!asset?.uri) {
    return { ok: false, error: 'No se obtuvo la ruta del archivo.' };
  }
  const text = await FileSystem.readAsStringAsync(asset.uri, {
    encoding: FileSystem.EncodingType.UTF8,
  });
  const inicio = text.replace(/^\uFEFF/, '').trimStart();
  if (inicio.startsWith(CSV_EXPORT_MAGIC)) {
    const parsedCsv = parsearRespaldoCsv(text);
    if (!parsedCsv.ok) {
      return { ok: false, error: parsedCsv.error };
    }
    return {
      ok: true,
      onboardingCompletado: parsedCsv.onboardingCompletado,
      data: parsedCsv.data,
      exportedAt: parsedCsv.exportedAt,
    };
  }
  const parsed = parsearRespaldoJson(text);
  if (!parsed.ok) {
    const parsedCsv2 = parsearRespaldoCsv(text);
    if (parsedCsv2.ok) {
      return {
        ok: true,
        onboardingCompletado: parsedCsv2.onboardingCompletado,
        data: parsedCsv2.data,
        exportedAt: parsedCsv2.exportedAt,
      };
    }
    return { ok: false, error: parsed.error };
  }
  return {
    ok: true,
    onboardingCompletado: parsed.onboardingCompletado,
    data: parsed.data,
    exportedAt: parsed.exportedAt,
  };
}
