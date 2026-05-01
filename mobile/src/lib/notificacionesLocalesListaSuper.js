/**
 * Recordatorios locales para artículos de la lista de compras marcados como **urgentes**.
 * Mismas franjas que pagos/TC: 9:00, 12:00 y 18:00 (Bogotá), con opcional aviso ~5 min tras sync el mismo día
 * si aplica (ver `instantesAvisoDiaBogota`). No en Expo Go; sí en dev build / APK.
 */
import { Platform } from 'react-native';
import * as Notifications from 'expo-notifications';
import { AndroidImportance } from 'expo-notifications';
import { notificacionesSistemaDisponibles } from './notificacionesLocalesEntorno';
import { instantesAvisoDiaBogota, ymdBogotaMasDias } from './notificacionesLocalesPagosProgramados';
import { varianteListaSuperUrgente } from './notificacionesVariantesAmigables';
import { tituloNotifConNombre } from './notificacionesPersonalizacion';
import { solicitarPermisosSiNoConcedidos } from './notificacionesPermisos';

const CANAL_LISTA_SUPER = 'lista-super-urgente-v1';
const TIPO_LISTA_SUPER_URGENTE = 'listaSuperUrgente';
const HORIZONTE_DIAS = 14;

const MAX_SEC_TIME_INTERVAL_ANDROID = 86400 * 60;

let canalListaListo = false;

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
  return solicitarPermisosSiNoConcedidos();
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
 * Programa avisos en franjas Bogotá (9 / 12 / 18 y opcional +5 min el mismo día) mientras haya ítems urgentes.
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

  const nombres = items.map((x) => String(x.nombre || '').trim()).filter(Boolean);
  const resumen = resumenTextoItemsUrgentes(items);
  const ahora = new Date();

  for (let offset = 0; offset <= HORIZONTE_DIAS; offset += 1) {
    const { y, mo, d } = ymdBogotaMasDias(ahora, offset);
    const instantes = instantesAvisoDiaBogota(y, mo, d, ahora);

    for (let slotIx = 0; slotIx < instantes.length; slotIx += 1) {
      const triggerAt = instantes[slotIx];
      const disparador = construirDisparador(triggerAt);
      if (disparador == null) continue;

      const idClave = `ls-urg-${y}-${mo}-${d}-s${slotIx}`;
      const { title: tituloRaw, body: cuerpo } = varianteListaSuperUrgente({ nombres, resumen });
      const titulo = tituloNotifConNombre(state, tituloRaw);

      try {
        await Notifications.scheduleNotificationAsync({
          identifier: idClave,
          content: {
            title: titulo,
            body: cuerpo,
            data: { tipo: TIPO_LISTA_SUPER_URGENTE, offsetDia: offset, slot: slotIx },
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
}
