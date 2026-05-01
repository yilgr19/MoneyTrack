/**
 * Solicitud unificada de permisos para notificaciones locales (y push).
 * Android 13+: debe existir al menos un canal antes del diálogo POST_NOTIFICATIONS (Expo docs).
 */
import { Platform, Linking } from 'react-native';
import * as Notifications from 'expo-notifications';
import { AndroidImportance } from 'expo-notifications';
import { notificacionesSistemaDisponibles } from './notificacionesLocalesEntorno';

const CANAL_ANTES_PERMISO = 'moneytrack-inicial-permisos-v1';
let canalBootstrapAndroidListo = false;

/** Opciones alineadas con expo-notifications / iOS UNAuthorizationOptions. */
export const solicitudPermisosNotificacionesOpciones = {
  ios: {
    allowAlert: true,
    allowBadge: true,
    allowSound: true,
  },
};

/** Alineado a UNAuthorizationStatus (iOS); evita depender del export en tiempo de carga. */
const IOS_UN_AUTH = { AUTHORIZED: 2, PROVISIONAL: 3, EPHEMERAL: 4 };

function permisoNotificacionesConcedido(r) {
  if (!r) return false;
  if (r.granted) return true;
  if (Platform.OS === 'ios' && r.ios && typeof r.ios.status === 'number') {
    const s = r.ios.status;
    const Auth = Notifications.IosAuthorizationStatus || IOS_UN_AUTH;
    return s === Auth.AUTHORIZED || s === Auth.PROVISIONAL || s === Auth.EPHEMERAL;
  }
  return false;
}

/**
 * Crea un canal en Android antes de pedir permiso (requisito API 33+).
 */
export async function asegurarCanalAndroidAntesDeSolicitarPermiso() {
  if (Platform.OS !== 'android' || canalBootstrapAndroidListo) return;
  await Notifications.setNotificationChannelAsync(CANAL_ANTES_PERMISO, {
    name: 'Recordatorios MoneyTrack',
    importance: AndroidImportance.HIGH,
    vibrationPattern: [0, 250, 250, 250],
    lightColor: '#4B246C',
    enableVibrate: true,
  });
  canalBootstrapAndroidListo = true;
}

/** Evita dos `requestPermissionsAsync` en paralelo (p. ej. arranque + sincronización). */
let solicitudPermisosEnCurso = null;

/**
 * Tras crear los canales de la app: comprueba y, si hace falta, muestra el diálogo del sistema.
 * @returns {Promise<boolean>} true si se puede mostrar notificaciones
 */
export async function solicitarPermisosSiNoConcedidos() {
  if (Platform.OS === 'web') return false;
  if (solicitudPermisosEnCurso) return solicitudPermisosEnCurso;
  solicitudPermisosEnCurso = (async () => {
    try {
      const prev = await Notifications.getPermissionsAsync();
      if (permisoNotificacionesConcedido(prev)) return true;
      if (prev.status === 'denied' && prev.canAskAgain === false) return false;
      const next = await Notifications.requestPermissionsAsync(solicitudPermisosNotificacionesOpciones);
      return permisoNotificacionesConcedido(next);
    } finally {
      solicitudPermisosEnCurso = null;
    }
  })();
  return solicitudPermisosEnCurso;
}

/**
 * Entrada temprana al arranque: canal bootstrap en Android + solicitud si aplica.
 * @returns {Promise<{ ok: boolean, skipped?: boolean, needsSettings?: boolean }>}
 */
export async function solicitarPermisosNotificacionesAlIniciar() {
  if (Platform.OS === 'web') {
    return { ok: false, skipped: true };
  }
  if (!notificacionesSistemaDisponibles()) {
    return { ok: false, skipped: true };
  }

  if (Platform.OS === 'android') {
    await asegurarCanalAndroidAntesDeSolicitarPermiso();
  }

  const prev = await Notifications.getPermissionsAsync();
  if (permisoNotificacionesConcedido(prev)) {
    return { ok: true };
  }
  if (prev.status === 'denied' && prev.canAskAgain === false) {
    return { ok: false, needsSettings: true };
  }

  const next = await Notifications.requestPermissionsAsync(solicitudPermisosNotificacionesOpciones);
  const ok = permisoNotificacionesConcedido(next);
  return {
    ok,
    needsSettings: !ok && next.canAskAgain === false,
  };
}

/** Abre los ajustes de la app (notificaciones / alarmas exactas según fabricante). */
export function abrirAjustesAppParaNotificaciones() {
  if (Platform.OS === 'web') return;
  Linking.openSettings().catch(() => {});
}
