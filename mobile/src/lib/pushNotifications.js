import { Platform } from 'react-native';
import AsyncStorage from '@react-native-async-storage/async-storage';
import Constants from 'expo-constants';
import * as Notifications from 'expo-notifications';
import { AndroidImportance } from 'expo-notifications';
import { rootNavigationRef } from '../navigation/rootNavigationRef';
import { notificacionesSistemaDisponibles } from './notificacionesLocalesEntorno';

const STORAGE_KEY = '@moneytrack/expoPushToken';
const STORAGE_DEVICE_INSTALL_ID = '@moneytrack/deviceInstallId';
const CANAL_DEFAULT = 'default';

function generarUuidInstalacion() {
  const bytes = new Uint8Array(16);
  for (let i = 0; i < 16; i += 1) bytes[i] = Math.floor(Math.random() * 256);
  bytes[6] = (bytes[6] & 0x0f) | 0x40;
  bytes[8] = (bytes[8] & 0x3f) | 0x80;
  const hex = Array.from(bytes, (b) => b.toString(16).padStart(2, '0')).join('');
  return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}`;
}

async function obtenerDeviceInstallId() {
  try {
    let id = await AsyncStorage.getItem(STORAGE_DEVICE_INSTALL_ID);
    if (!id || String(id).length < 8) {
      id = generarUuidInstalacion();
      await AsyncStorage.setItem(STORAGE_DEVICE_INSTALL_ID, id);
    }
    return id;
  } catch {
    return generarUuidInstalacion();
  }
}

let canalDefaultListo = false;
let listenersRegistrados = false;

/**
 * Push remoto (Expo Push Service).
 *
 * Requisitos:
 * - Build con development client o producción (no Expo Go).
 * - Proyecto EAS: `npx eas-cli@latest init` y copiar `projectId` en app.json → expo.extra.eas.projectId
 * - Credenciales FCM (Android) y APNs (iOS): `eas credentials` o panel EAS.
 * - Envío: https://docs.expo.dev/push-notifications/sending-notifications/
 *
 * Opcional:
 * - `EXPO_PUBLIC_PUSH_REGISTER_URL` — POST JSON { expoPushToken, platform, deviceInstallId }.
 * - `EXPO_PUBLIC_PUSH_API_KEY` — cabecera X-MoneyTrack-Api-Key (misma clave que MONEYTRACK_PUSH_API_KEY en Laravel).
 *
 * Payload sugerido al tocar la notificación (campo `data`):
 * { "rootScreen": "Mas", "nestedScreen": "PagosProgramados", "nestedParams": {} }
 * Pestañas raíz: Inicio | Gastos | Saldo | Mas
 */

export function getExpoProjectId() {
  return Constants.expoConfig?.extra?.eas?.projectId ?? Constants.easConfig?.projectId ?? undefined;
}

export async function leerTokenPushGuardado() {
  try {
    return await AsyncStorage.getItem(STORAGE_KEY);
  } catch {
    return null;
  }
}

async function asegurarCanalDefaultAndroid() {
  if (Platform.OS !== 'android' || canalDefaultListo) return;
  await Notifications.setNotificationChannelAsync(CANAL_DEFAULT, {
    name: 'Avisos',
    importance: AndroidImportance.HIGH,
    vibrationPattern: [0, 250, 250, 250],
    lightColor: '#4B246C',
  });
  canalDefaultListo = true;
}

async function permisosPushOk() {
  const { status: prev } = await Notifications.getPermissionsAsync();
  if (prev === 'granted') return true;
  const { status } = await Notifications.requestPermissionsAsync({
    ios: { allowAlert: true, allowBadge: true, allowSound: true },
  });
  return status === 'granted';
}

async function enviarTokenAlBackendOpcional(token) {
  const url =
    typeof process !== 'undefined' && process.env && process.env.EXPO_PUBLIC_PUSH_REGISTER_URL
      ? String(process.env.EXPO_PUBLIC_PUSH_REGISTER_URL).trim()
      : '';
  if (!url || !token) return;
  const apiKey =
    typeof process !== 'undefined' && process.env && process.env.EXPO_PUBLIC_PUSH_API_KEY
      ? String(process.env.EXPO_PUBLIC_PUSH_API_KEY).trim()
      : '';
  const deviceInstallId = await obtenerDeviceInstallId();
  const platform = Platform.OS === 'ios' ? 'ios' : 'android';
  try {
    const headers = {
      'Content-Type': 'application/json',
      Accept: 'application/json',
    };
    if (apiKey) headers['X-MoneyTrack-Api-Key'] = apiKey;
    await fetch(url, {
      method: 'POST',
      headers,
      body: JSON.stringify({
        expoPushToken: token,
        platform,
        deviceInstallId,
      }),
    });
  } catch {
    /* backend opcional */
  }
}

/**
 * Obtiene el token Expo Push, lo guarda y opcionalmente lo envía al servidor.
 * @returns {Promise<string|null>}
 */
export async function registrarTokenPushExpo() {
  if (Platform.OS === 'web' || !notificacionesSistemaDisponibles()) return null;

  const ok = await permisosPushOk();
  if (!ok) return null;

  await asegurarCanalDefaultAndroid();

  const projectId = getExpoProjectId();
  try {
    const tokenRes = await Notifications.getExpoPushTokenAsync(
      projectId ? { projectId } : undefined
    );
    const token = tokenRes?.data;
    if (!token) return null;
    await AsyncStorage.setItem(STORAGE_KEY, token);
    await enviarTokenAlBackendOpcional(token);
    return token;
  } catch (e) {
    if (typeof __DEV__ !== 'undefined' && __DEV__) {
      console.warn('[MoneyTrack] Expo Push:', e?.message || e);
      if (!projectId) {
        console.warn(
          '[MoneyTrack] Falta expo.extra.eas.projectId en app.json. Ejecuta eas init y añádelo, o usa EAS Build.'
        );
      }
    }
    return null;
  }
}

function navegarDesdeDataPush(data) {
  if (!data || typeof data !== 'object') return;
  const rootScreen = data.rootScreen;
  if (!rootScreen || typeof rootScreen !== 'string') return;
  if (!rootNavigationRef.isReady()) return;
  try {
    const nested = data.nestedScreen;
    if (nested && typeof nested === 'string') {
      rootNavigationRef.navigate(rootScreen, {
        screen: nested,
        params: data.nestedParams && typeof data.nestedParams === 'object' ? data.nestedParams : undefined,
      });
    } else {
      const rp = data.rootParams;
      rootNavigationRef.navigate(rootScreen, rp && typeof rp === 'object' ? rp : undefined);
    }
  } catch {
    /* best-effort */
  }
}

function registrarListenersUnaVez() {
  if (listenersRegistrados) return;
  listenersRegistrados = true;

  Notifications.addNotificationResponseReceivedListener((response) => {
    const data = response?.notification?.request?.content?.data;
    navegarDesdeDataPush(data);
  });

  Notifications.addPushTokenListener(() => {
    registrarTokenPushExpo().catch(() => {});
  });
}

/**
 * Si la app se abrió desde una notificación (cold start), aplica navegación cuando el árbol esté listo.
 */
export function aplicarUltimaRespuestaNotificacionSiAplica() {
  if (Platform.OS === 'web' || !notificacionesSistemaDisponibles()) return;
  try {
    const last = Notifications.getLastNotificationResponse();
    const data = last?.notification?.request?.content?.data;
    if (!data) return;
    let intentos = 0;
    const max = 40;
    const t = setInterval(() => {
      intentos += 1;
      if (rootNavigationRef.isReady()) {
        clearInterval(t);
        navegarDesdeDataPush(data);
        try {
          Notifications.clearLastNotificationResponse();
        } catch {
          /* SDK antiguo */
        }
        return;
      }
      if (intentos >= max) clearInterval(t);
    }, 50);
  } catch {
    /* no bloquear arranque */
  }
}

/**
 * Registrar listeners + token. Llamar una vez tras `registrarHandlerNotificacionesLocales` y con la app lista.
 */
export async function inicializarPushNotificaciones() {
  if (Platform.OS === 'web' || !notificacionesSistemaDisponibles()) return;
  registrarListenersUnaVez();
  await registrarTokenPushExpo();
  aplicarUltimaRespuestaNotificacionSiAplica();
}
