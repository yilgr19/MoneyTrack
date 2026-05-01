/**
 * Respaldo completo de datos locales (JSON) para migrar entre instalaciones sin perder historial.
 * Extensión recomendada: .moneytrack.json
 *
 * El objeto raíz incluye `schema` (orden de campos y notas) para que otra app o un script
 * pueda mapear los mismos datos sin adivinar el shape.
 *
 * Export: archivo .moneytrack.json + compartir. iOS: Share con `url`. Android: expo-sharing solo si existe el nativo
 * (NativeModules.ExpoSharing); si no, nunca se hace require y no hay error — reserva texto.
 * Import: `expo-document-picker`. FS: API legacy (SDK 54).
 */
import { NativeModules, Platform, Share } from 'react-native';
import * as FileSystem from 'expo-file-system/legacy';
import * as DocumentPicker from 'expo-document-picker';

export const BACKUP_FORMAT = 'moneytrack-respaldo';
export const BACKUP_VERSION = 1;

/** Mismo orden que devuelve `normalizeState` en AppContext (export estable y legible). */
export const BACKUP_DATA_KEYS_ORDERED = [
  'moneda',
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

/** Límite práctico para compartir JSON como texto en `message` (varía por fabricante). */
const MAX_JSON_SHARE_CHARS = 350_000;

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
  let obj;
  try {
    obj = JSON.parse(texto);
  } catch {
    return { ok: false, error: 'No es un JSON válido.' };
  }
  if (!obj || typeof obj !== 'object') {
    return { ok: false, error: 'Contenido inválido.' };
  }
  if (obj.format !== BACKUP_FORMAT) {
    return {
      ok: false,
      error: 'Este archivo no es un respaldo de MoneyTrack (.moneytrack.json).',
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
  return `MoneyTrack-respaldo-${y}-${m}-${day}_${h}${min}.moneytrack.json`;
}

/**
 * Comparte el archivo de respaldo sin provocar "Cannot find native module 'ExpoSharing'":
 * en Android no se hace `require('expo-sharing')` si el nativo no está enlazado.
 */
async function compartirArchivoRespaldo(fileUri) {
  if (Platform.OS === 'ios') {
    try {
      await Share.share({
        title: 'Respaldo MoneyTrack',
        url: fileUri,
      });
      return true;
    } catch {
      return false;
    }
  }

  if (Platform.OS === 'android' && NativeModules.ExpoSharing) {
    try {
      const Sharing = require('expo-sharing');
      if (!(await Sharing.isAvailableAsync())) return false;
      await Sharing.shareAsync(fileUri, {
        mimeType: 'application/json',
        dialogTitle: 'Guardar respaldo MoneyTrack',
      });
      return true;
    } catch {
      return false;
    }
  }

  return false;
}

async function exportarRespaldoSoloTexto(jsonCompacto) {
  if (jsonCompacto.length > MAX_JSON_SHARE_CHARS) {
    const kb = Math.round(jsonCompacto.length / 1000);
    return {
      ok: false,
      mensaje: `El respaldo es muy grande (~${kb} KB). Instala la última versión de la app (incluye exportar como archivo) o reduce datos.`,
    };
  }
  try {
    await Share.share({
      title: 'Respaldo MoneyTrack',
      message: jsonCompacto,
    });
    return {
      ok: true,
      mensaje:
        Platform.OS === 'android' && !NativeModules.ExpoSharing
          ? 'Esta instalación aún no puede adjuntar un archivo automático. Instala la última versión de MoneyTrack (APK generado de nuevo en tu PC o EAS) y el export será un archivo listo para importar. Mientras tanto puedes guardar este texto como respaldo.'
          : 'Tu teléfono compartió el respaldo como texto. Si puedes, actualiza la app para exportar un archivo directo.',
    };
  } catch (e) {
    return { ok: false, mensaje: e?.message || 'No se pudo abrir el menú compartir.' };
  }
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
  const payload = crearPayloadRespaldo(state, onboardingCompletado);
  const jsonPretty = serializarRespaldo(payload, true);
  const jsonCompacto = serializarRespaldo(payload, false);

  const base = FileSystem.cacheDirectory;
  if (!base) {
    return exportarRespaldoSoloTexto(jsonCompacto);
  }

  const fileUri = `${base}${nombreArchivoRespaldoLegible()}`;
  try {
    await FileSystem.writeAsStringAsync(fileUri, jsonPretty, {
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
        'Elige dónde guardar el archivo (por ejemplo Descargas, Drive o Archivos). En el otro dispositivo abre MoneyTrack → Administrar → Importar datos y selecciona ese archivo.',
    };
  }

  return exportarRespaldoSoloTexto(jsonCompacto);
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
      type: ['application/json', 'text/plain', 'application/octet-stream', '*/*'],
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
  const parsed = parsearRespaldoJson(text);
  if (!parsed.ok) {
    return { ok: false, error: parsed.error };
  }
  return {
    ok: true,
    onboardingCompletado: parsed.onboardingCompletado,
    data: parsed.data,
    exportedAt: parsed.exportedAt,
  };
}
