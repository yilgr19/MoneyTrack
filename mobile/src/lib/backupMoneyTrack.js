/**
 * Respaldo completo de datos locales (JSON) para migrar entre instalaciones APK sin perder historial.
 * Extensión recomendada: .moneytrack.json
 *
 * Export usa `Share` de React Native (no requiere módulo nativo extra). Import usa `expo-document-picker`
 * al pulsar «Importar» (carga diferida).
 */
import { Platform, Share } from 'react-native';
import * as FileSystem from 'expo-file-system';

export const BACKUP_FORMAT = 'moneytrack-respaldo';
export const BACKUP_VERSION = 1;

/** Límite práctico para compartir JSON como texto (varía por fabricante). */
const MAX_JSON_SHARE_CHARS = 350_000;

/**
 * @param {object} state - Estado normalizado en memoria (mismo shape que usa AppContext).
 * @param {boolean} onboardingCompletado
 */
export function crearPayloadRespaldo(state, onboardingCompletado) {
  return {
    format: BACKUP_FORMAT,
    version: BACKUP_VERSION,
    exportedAt: new Date().toISOString(),
    app: 'MoneyTrack',
    onboardingCompletado: !!onboardingCompletado,
    data: JSON.parse(JSON.stringify(state)),
  };
}

export function serializarRespaldo(payload) {
  return JSON.stringify(payload);
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

/**
 * Abre el menú «Compartir» del sistema con el JSON (no depende de expo-sharing / rebuild nativo).
 */
export async function exportarRespaldoCompartir(state, onboardingCompletado) {
  if (Platform.OS === 'web') {
    return {
      ok: false,
      mensaje: 'Exportar a archivo está disponible en la app instalada (Android/iOS).',
    };
  }
  const payload = crearPayloadRespaldo(state, onboardingCompletado);
  const json = serializarRespaldo(payload);
  if (json.length > MAX_JSON_SHARE_CHARS) {
    const kb = Math.round(json.length / 1000);
    return {
      ok: false,
      mensaje: `El respaldo es muy grande (~${kb} KB) para compartir como texto en este dispositivo. Pronto añadiremos exportación a archivo; mientras tanto reduce datos o usa la versión web si aplica.`,
    };
  }
  try {
    await Share.share({
      title: 'Respaldo MoneyTrack',
      message: json,
    });
    return {
      ok: true,
      mensaje:
        'Elige una app (archivos, notas, correo…). Guarda el contenido como archivo .moneytrack.json si tu app lo permite, o copia el texto a un archivo con ese nombre.',
    };
  } catch (e) {
    return { ok: false, mensaje: e?.message || 'No se pudo abrir el menú compartir.' };
  }
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
  let DocumentPicker;
  try {
    DocumentPicker = await import('expo-document-picker');
  } catch (e) {
    return {
      ok: false,
      error:
        'No está disponible el selector de archivos. Vuelve a generar la app con `npx expo prebuild` y `npx expo run:android` (o tu build EAS) para incluir módulos Expo.',
    };
  }
  const pick = await DocumentPicker.getDocumentAsync({
    type: ['application/json', 'text/plain', 'application/octet-stream', '*/*'],
    copyToCacheDirectory: true,
  });
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
