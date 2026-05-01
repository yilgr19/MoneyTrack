/**
 * Recordatorios locales para artículos de la lista de compras marcados como **urgentes**.
 * Misma disponibilidad que pagos/TC: no en Expo Go; sí en dev build / APK.
 */
import { Platform } from 'react-native';
import * as Notifications from 'expo-notifications';
import { AndroidImportance } from 'expo-notifications';
import { notificacionesSistemaDisponibles } from './notificacionesLocalesEntorno';
import { AVISOS_LOCALES_USAR_HORA_CERCANA } from './notificacionesLocalesPagosProgramados';
import { varianteListaSuperUrgente } from './notificacionesVariantesAmigables';

const CANAL_LISTA_SUPER = 'lista-super-urgente-v1';
const TIPO_LISTA_SUPER_URGENTE = 'listaSuperUrgente';
const HORIZONTE_DIAS = 14;
const HORA_AVISO = 10;
const MINUTO_AVISO = 30;
const MINUTOS_TRAS_SINCRON_PARA_HOY = 2;

const MAX_SEC_TIME_INTERVAL_ANDROID = 86400 * 60;

let canalListaListo = false;

function mismoDiaCalendario(dA, dB) {
  return (
    dA.getFullYear() === dB.getFullYear() &&
    dA.getMonth() === dB.getMonth() &&
    dA.getDate() === dB.getDate()
  );
}

function diaHoraAvisoListaSuper(refDia, ahora) {
  if (AVISOS_LOCALES_USAR_HORA_CERCANA && mismoDiaCalendario(refDia, ahora)) {
    return new Date(ahora.getTime() + MINUTOS_TRAS_SINCRON_PARA_HOY * 60_000);
  }
  return new Date(
    refDia.getFullYear(),
    refDia.getMonth(),
    refDia.getDate(),
    HORA_AVISO,
    MINUTO_AVISO,
    0
  );
}

function ajustarDisparadorSiHoraPasada(refDia, ahora, triggerAt) {
  const t = triggerAt.getTime();
  if (t > ahora.getTime()) return triggerAt;
  const finDiaRef = new Date(refDia.getFullYear(), refDia.getMonth(), refDia.getDate(), 23, 59, 59, 999);
  if (ahora.getTime() > finDiaRef.getTime()) return null;
  return new Date(ahora.getTime() + 60_000);
}

function construirDisparador(triggerAt) {
  const ms = triggerAt.getTime() - Date.now();
  const sec = Math.ceil(ms / 1000);
  if (sec < 15) return null;
  if (Platform.OS === 'android' && sec <= MAX_SEC_TIME_INTERVAL_ANDROID) {
    return {
      type: Notifications.SchedulableTriggerInputTypes.TIME_INTERVAL,
      seconds: sec,
      channelId: CANAL_LISTA_SUPER,
    };
  }
  return {
    type: Notifications.SchedulableTriggerInputTypes.DATE,
    date: triggerAt,
    ...(Platform.OS === 'android' && { channelId: CANAL_LISTA_SUPER }),
  };
}

export function resumenTextoItemsUrgentes(lineas) {
  const nombres = (lineas || [])
    .map((l) => String(l?.nombre || '').trim())
    .filter(Boolean);
  if (nombres.length === 0) return '';
  if (nombres.length === 1) return nombres[0];
  if (nombres.length === 2) return `${nombres[0]} y ${nombres[1]}`;
  return `${nombres[0]}, ${nombres[1]} y ${nombres.length - 2} más`;
}

async function asegurarCanalListaSuper() {
  if (Platform.OS === 'web') return false;
  if (Platform.OS === 'android' && !canalListaListo) {
    await Notifications.setNotificationChannelAsync(CANAL_LISTA_SUPER, {
      name: 'Lista de compras · urgente',
      importance: AndroidImportance.HIGH,
      vibrationPattern: [0, 200, 150, 200],
      lightColor: '#2dd4bf',
      enableVibrate: true,
    });
    canalListaListo = true;
  }
  let { status } = await Notifications.getPermissionsAsync();
  if (status !== 'granted') {
    const r = await Notifications.requestPermissionsAsync({
      ios: { allowAlert: true, allowBadge: true, allowSound: true },
    });
    status = r.status;
  }
  return status === 'granted';
}

async function cancelarProgramadasListaSuper() {
  const list = await Notifications.getAllScheduledNotificationsAsync();
  await Promise.all(
    list
      .filter((r) => r?.content?.data?.tipo === TIPO_LISTA_SUPER_URGENTE)
      .map((r) => Notifications.cancelScheduledNotificationAsync(r.identifier))
  );
}

/**
 * Programa un aviso al día (variación aleatoria entre 20 textos) mientras haya ítems urgentes en la lista.
 */
export async function sincronizarNotificacionesListaSuperUrgente(state) {
  if (Platform.OS === 'web' || !notificacionesSistemaDisponibles()) return;
  const ok = await asegurarCanalListaSuper();
  if (!ok) return;

  await cancelarProgramadasListaSuper();

  const items = (state?.listaSuperCompraItems || []).filter(
    (x) => x && String(x.nombre || '').trim() && x.urgencia === 'urgente'
  );
  if (items.length === 0) return;

  const resumen = resumenTextoItemsUrgentes(items);
  const ahora = new Date();

  for (let offset = 0; offset <= HORIZONTE_DIAS; offset += 1) {
    const refDia = new Date(ahora.getFullYear(), ahora.getMonth(), ahora.getDate() + offset, 12, 0, 0);
    const y = refDia.getFullYear();
    const mo = refDia.getMonth() + 1;
    const da = refDia.getDate();
    const idClave = `ls-urg-${y}-${mo}-${da}`;

    const { title: titulo, body: cuerpo } = varianteListaSuperUrgente(resumen);

    const ideal = diaHoraAvisoListaSuper(refDia, ahora);
    const triggerAt = ajustarDisparadorSiHoraPasada(refDia, ahora, ideal);
    if (triggerAt == null) continue;
    const disparador = construirDisparador(triggerAt);
    if (disparador == null) continue;

    try {
      await Notifications.scheduleNotificationAsync({
        identifier: idClave,
        content: {
          title: titulo,
          body: cuerpo,
          data: { tipo: TIPO_LISTA_SUPER_URGENTE, offsetDia: offset },
          sound: true,
          ...(Platform.OS === 'android' && {
            channelId: CANAL_LISTA_SUPER,
            priority: 'high',
          }),
        },
        trigger: disparador,
      });
    } catch (e) {
      if (typeof __DEV__ !== 'undefined' && __DEV__) {
        console.warn('[MoneyTrack] schedule lista super:', e?.message || e);
      }
    }
  }
}
