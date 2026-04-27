import AsyncStorage from '@react-native-async-storage/async-storage';

export const NOTIFICACIONES_LECTURA_KEY = 'notificacionesFirmasLectura';
const KEY = NOTIFICACIONES_LECTURA_KEY;

/**
 * Identifica la “versión” de un aviso. Si el texto cambia, vuelve a contar como no leído.
 */
export function firmaNotificacion(item) {
  return `${String(item.titulo || '')}||${String(item.detalle || '')}`;
}

/**
 * @param {Array<{id,titulo,detalle}>} items
 * @param {Record<string, string> | null} firmas id → última firma leída
 */
export function contarNoLeidas(items, firmas) {
  if (!items.length) return 0;
  if (!firmas) return items.length;
  return items.filter((it) => firmas[it.id] !== firmaNotificacion(it)).length;
}

export async function loadFirmasLectura() {
  try {
    const s = await AsyncStorage.getItem(KEY);
    if (!s) return {};
    const o = JSON.parse(s);
    return o && typeof o === 'object' ? o : {};
  } catch {
    return {};
  }
}

export async function saveFirmasLectura(firmas) {
  await AsyncStorage.setItem(KEY, JSON.stringify(firmas));
}

/**
 * Graba que el usuario vio el estado actual de la lista: cada id pasa a la firma actual.
 * Si una notificación vuelve con el mismo id pero detalle distinto, volverá a no leída.
 */
export function marcarAvisosActualesComoVistos(items, firmasPrevias) {
  const next = { ...firmasPrevias };
  for (const it of items) {
    next[it.id] = firmaNotificacion(it);
  }
  return next;
}
