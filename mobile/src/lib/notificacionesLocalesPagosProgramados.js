/**
 * Notificaciones que el SO entrega con la app cerrada (alarmas locales programadas).
 * No funcionan en Expo Go — hace falta development build o APK/IPA (ver notificacionesLocalesEntorno).
 */
import { Platform } from 'react-native';
import * as Notifications from 'expo-notifications';
import { AndroidImportance } from 'expo-notifications';
import { diasHastaPagoProgramado } from './notificacionesApp';
import { entornoEsExpoGo, notificacionesSistemaDisponibles } from './notificacionesLocalesEntorno';
import {
  formatearNumero,
  resumenAlertasTarjetasCredito,
  construirExtractoBancarioTarjeta,
  montoPagoSugeridoDesdeExtracto,
} from './finance';
import {
  varianteNotifPagoProgramado,
  varianteNotifPagoPruebaLejano,
  varianteNotifPruebaSistema,
  varianteNotifTcCorteHoy,
  varianteNotifTcCorteManana,
  varianteNotifTcCorte2d,
  varianteNotifTcPago,
} from './notificacionesVariantesAmigables';

/** `-v2`: en Android los canales no se pueden “subir” de importancia; nuevo id fuerza cabecera/sonido correctos. */
const CANAL_ANDROID = 'pagos-programados-v2';
const CANAL_TARJETAS = 'tarjetas-credito-v2';
const TIPO_DATA = 'pagoProgramado';
const TIPO_TC_CORTE = 'tcCorteLocal';
const TIPO_TC_PAGO = 'tcPagoLocal';
const TIPO_PRUEBA = 'notifPruebaMoneyTrack';
/** Mismo rango que la campana: hoy (0) y 1–3 días antes del vencimiento. */
const DIAS_MAX = 3;
/** Días hacia adelante en que programamos avisos (incluye varios ciclos mensuales). */
const HORIZONTE_DIAS = 45;
const HORA_AVISO_PRODUCCION = 9;
/**
 * `false` (producción): avisos a las ~9:00 del día que corresponde.
 * `true` (pruebas): los avisos cuyo día es HOY se programan unos minutos después de abrir/sincronizar la app;
 *   los de mañana o después siguen a las 9:00. Pon `false` cuando termines de probar en barra.
 */
export const AVISOS_LOCALES_USAR_HORA_CERCANA = true;
const MINUTOS_TRAS_SINCRON_PARA_HOY = 2;

let canalesAndroidListos = false;

function mismoDiaCalendario(dA, dB) {
  return (
    dA.getFullYear() === dB.getFullYear() &&
    dA.getMonth() === dB.getMonth() &&
    dA.getDate() === dB.getDate()
  );
}

function diaHoraAviso(refDia, ahora) {
  if (AVISOS_LOCALES_USAR_HORA_CERCANA && mismoDiaCalendario(refDia, ahora)) {
    return new Date(ahora.getTime() + MINUTOS_TRAS_SINCRON_PARA_HOY * 60_000);
  }
  return new Date(
    refDia.getFullYear(),
    refDia.getMonth(),
    refDia.getDate(),
    HORA_AVISO_PRODUCCION,
    0,
    0
  );
}

/** Si la hora objetivo ya pasó ese día, programa en ~1 min (p. ej. modo 9:00 y ya son las 15:00). */
function ajustarDisparadorSiYaPasaronLasNueve(refDia, ahora, triggerAt) {
  const t = triggerAt.getTime();
  if (t > ahora.getTime()) return triggerAt;
  const finDiaRef = new Date(refDia.getFullYear(), refDia.getMonth(), refDia.getDate(), 23, 59, 59, 999);
  if (ahora.getTime() > finDiaRef.getTime()) return null;
  return new Date(ahora.getTime() + 60_000);
}

/** Límite ~60 días: en Android `TIME_INTERVAL` suele disparar mejor que `DATE` en MIUI/Samsung con app cerrada. */
const MAX_SEC_TIME_INTERVAL_ANDROID = 86400 * 60;

/**
 * Programa un disparo único. En Android usa `TIME_INTERVAL` hasta ~60 días; iOS usa `DATE`.
 */
function construirDisparador(triggerAt, channelIdAndroid = CANAL_ANDROID) {
  const ms = triggerAt.getTime() - Date.now();
  const sec = Math.ceil(ms / 1000);
  if (sec < 15) {
    if (typeof __DEV__ !== 'undefined' && __DEV__) {
      console.warn('[MoneyTrack] Aviso local omitido: el disparador queda a menos de 15 s.');
    }
    return null;
  }
  if (Platform.OS === 'android' && sec <= MAX_SEC_TIME_INTERVAL_ANDROID) {
    return {
      type: Notifications.SchedulableTriggerInputTypes.TIME_INTERVAL,
      seconds: sec,
      channelId: channelIdAndroid,
    };
  }
  return {
    type: Notifications.SchedulableTriggerInputTypes.DATE,
    date: triggerAt,
    ...(Platform.OS === 'android' && { channelId: channelIdAndroid }),
  };
}

function epochAproximadoDesdeTrigger(trig) {
  if (!trig || typeof trig !== 'object') return NaN;
  const t = trig.type;
  if (t === Notifications.SchedulableTriggerInputTypes.DATE || t === 'date') {
    const raw = trig.date ?? trig.value ?? trig.timestamp;
    if (typeof raw === 'number') return raw;
    if (raw instanceof Date) return raw.getTime();
  }
  if (t === Notifications.SchedulableTriggerInputTypes.TIME_INTERVAL || t === 'timeInterval') {
    const s = Number(trig.seconds);
    if (Number.isFinite(s) && s > 0) return Date.now() + s * 1000;
  }
  return NaN;
}

function lineasProximosDisparos(list, maxN = 6) {
  if (!Array.isArray(list) || list.length === 0) return [];
  const withT = list
    .map((req) => {
      const ep = epochAproximadoDesdeTrigger(req.trigger);
      const title = String(req.content?.title || '(aviso)').trim() || '(aviso)';
      return { ep, title };
    })
    .filter((x) => Number.isFinite(x.ep));
  if (withT.length === 0) return ['  (disparadores no legibles: revisa «Alarmas exactas» en ajustes del teléfono)'];
  withT.sort((a, b) => a.ep - b.ep);
  const now = Date.now();
  return withT.slice(0, maxN).map((x) => {
    const min = Math.max(0, Math.round((x.ep - now) / 60_000));
    const fecha = new Date(x.ep).toLocaleString('es', { dateStyle: 'short', timeStyle: 'short' });
    const rel = min < 180 ? `~${min} min` : fecha;
    return `  → ${x.title.slice(0, 40)}: ${rel}`;
  });
}

/**
 * Debe llamarse al arranque (App.js) para que las notificaciones locales se muestren con el handler por defecto.
 */
export function registrarHandlerNotificacionesLocales() {
  if (Platform.OS === 'web') return;
  Notifications.setNotificationHandler({
    handleNotification: async () => ({
      shouldShowAlert: true,
      shouldPlaySound: true,
      shouldSetBadge: false,
    }),
  });
}

async function asegurarCanalYPermisos() {
  if (Platform.OS === 'web') return false;
  if (Platform.OS === 'android' && !canalesAndroidListos) {
    await Notifications.setNotificationChannelAsync(CANAL_ANDROID, {
      name: 'Pagos programados',
      importance: AndroidImportance.MAX,
      vibrationPattern: [0, 250, 250, 250],
      lightColor: '#4B246C',
      enableVibrate: true,
    });
    await Notifications.setNotificationChannelAsync(CANAL_TARJETAS, {
      name: 'Tarjetas · corte y pago',
      importance: AndroidImportance.MAX,
      vibrationPattern: [0, 250, 250, 250],
      lightColor: '#4B246C',
      enableVibrate: true,
    });
    canalesAndroidListos = true;
  }
  let { status } = await Notifications.getPermissionsAsync();
  if (status !== 'granted') {
    const r = await Notifications.requestPermissionsAsync({
      ios: { allowAlert: true, allowBadge: true, allowSound: true },
    });
    status = r.status;
  }
  if (status !== 'granted') {
    if (typeof __DEV__ !== 'undefined' && __DEV__) {
      console.warn('[MoneyTrack] Sin permiso de notificaciones: no aparecerán en la barra del sistema.');
    }
    return false;
  }
  return true;
}

async function cancelarProgramadasPagos() {
  const list = await Notifications.getAllScheduledNotificationsAsync();
  await Promise.all(
    list
      .filter((r) => r?.content?.data?.tipo === TIPO_DATA)
      .map((r) => Notifications.cancelScheduledNotificationAsync(r.identifier))
  );
}

async function cancelarProgramadasTarjetas() {
  const list = await Notifications.getAllScheduledNotificationsAsync();
  await Promise.all(
    list
      .filter((r) => {
        const t = r?.content?.data?.tipo;
        return t === TIPO_TC_CORTE || t === TIPO_TC_PAGO;
      })
      .map((r) => Notifications.cancelScheduledNotificationAsync(r.identifier))
  );
}

/**
 * Programa notificaciones locales (mismo día ~9:00) cada día en que falten 1–3 días para un pago activo.
 * Debe llamarse tras cambios en `pagosProgramados` o al abrir la app.
 */
export async function sincronizarNotificacionesLocalesPagosProgramados(state) {
  if (Platform.OS === 'web') return;
  const ok = await asegurarCanalYPermisos();
  if (!ok) return;

  await cancelarProgramadasPagos();

  const pagos = state?.pagosProgramados || [];
  const moneda = String(state?.moneda || '').trim();
  const ahora = new Date();
  let programadas = 0;

  for (const p of pagos) {
    if (!p || p.activo === false) continue;

    for (let offset = 0; offset <= HORIZONTE_DIAS; offset += 1) {
      const refDia = new Date(ahora.getFullYear(), ahora.getMonth(), ahora.getDate() + offset, 12, 0, 0);
      const d = diasHastaPagoProgramado(p, refDia);
      if (d == null || d < 0 || d > DIAS_MAX) continue;

      const ideal = diaHoraAviso(refDia, ahora);
      const triggerAt = ajustarDisparadorSiYaPasaronLasNueve(refDia, ahora, ideal);
      if (triggerAt == null) continue;

      const disparador = construirDisparador(triggerAt);
      if (disparador == null) continue;

      const concepto = String(p.concepto || 'Pago').trim() || 'Pago programado';
      const montoStr = formatearNumero(p.monto);
      const { title: titulo, body: cuerpo } = varianteNotifPagoProgramado(d, {
        concepto,
        montoStr,
        moneda,
      });
      const idBase = String(p.id != null ? p.id : 'sin-id');
      const idClave = `pp-loc-${idBase}-${refDia.getFullYear()}-${refDia.getMonth() + 1}-${refDia.getDate()}`;

      try {
        await Notifications.scheduleNotificationAsync({
          identifier: idClave,
          content: {
            title: titulo,
            body: cuerpo,
            data: { tipo: TIPO_DATA, pagoId: p.id, dias: d },
            sound: true,
            ...(Platform.OS === 'android' && {
              channelId: CANAL_ANDROID,
              priority: 'high',
            }),
          },
          trigger: disparador,
        });
        programadas += 1;
      } catch (e) {
        console.warn('[MoneyTrack] scheduleNotificationAsync:', e?.message || e);
      }
    }
  }

  /**
   * Con `AVISOS_LOCALES_USAR_HORA_CERCANA`, si no hubo ninguna programación (vencimiento a más de 3 días),
   * agenda un aviso de prueba en ~2 min para validar permisos y app cerrada. En producción (`false`) no aplica.
   */
  if (AVISOS_LOCALES_USAR_HORA_CERCANA && programadas === 0) {
    const refHoyMediodia = new Date(ahora.getFullYear(), ahora.getMonth(), ahora.getDate(), 12, 0, 0);
    let idxPrueba = 0;
    for (const p of pagos) {
      if (!p || p.activo === false) continue;
      const dHoy = diasHastaPagoProgramado(p, refHoyMediodia);
      if (dHoy == null || dHoy < 0 || dHoy <= DIAS_MAX || dHoy > 366) continue;

      const ideal = new Date(
        ahora.getTime() + MINUTOS_TRAS_SINCRON_PARA_HOY * 60_000 + idxPrueba * 45_000
      );
      idxPrueba += 1;
      const triggerAt = ajustarDisparadorSiYaPasaronLasNueve(refHoyMediodia, ahora, ideal);
      if (triggerAt == null) continue;
      const disparador = construirDisparador(triggerAt);
      if (disparador == null) continue;

      const concepto = String(p.concepto || 'Pago').trim() || 'Pago programado';
      const montoStr = formatearNumero(p.monto);
      const idBase = String(p.id != null ? p.id : 'sin-id');
      const y = refHoyMediodia.getFullYear();
      const mo = refHoyMediodia.getMonth() + 1;
      const da = refHoyMediodia.getDate();
      const idClave = `pp-loc-prueba-lejano-${idBase}-${y}-${mo}-${da}`;
      const pruebaTxt = varianteNotifPagoPruebaLejano({ concepto, montoStr, moneda, dHoy });

      try {
        await Notifications.scheduleNotificationAsync({
          identifier: idClave,
          content: {
            title: pruebaTxt.title,
            body: pruebaTxt.body,
            data: { tipo: TIPO_DATA, pagoId: p.id, dias: dHoy, pruebaLejano: true },
            sound: true,
            ...(Platform.OS === 'android' && {
              channelId: CANAL_ANDROID,
              priority: 'high',
            }),
          },
          trigger: disparador,
        });
        programadas += 1;
      } catch (e) {
        console.warn('[MoneyTrack] scheduleNotificationAsync (prueba lejano):', e?.message || e);
      }
    }
  }

  const hayActivos = pagos.some((x) => x && x.activo !== false);
  if (hayActivos && programadas === 0) {
    console.warn(
      '[MoneyTrack] Sin notificaciones locales de pagos: solo se programan cuando faltan 0–3 días para el vencimiento, ' +
        'o el pago está inactivo / sin fechas válidas. En Expo Go no hay barra del sistema; usa dev build o APK. ' +
        'Tarjetas (corte/pago) van por otro canal.'
    );
  }
}

/**
 * Avisos en la barra del sistema (app cerrada) para corte y pago de TC, alineados con la campana.
 * Reprogramar al abrir la app o al cambiar tarjetas / gastos (pago al día).
 */
export async function sincronizarNotificacionesLocalesTarjetasCredito(state) {
  if (Platform.OS === 'web') return;
  const ok = await asegurarCanalYPermisos();
  if (!ok) return;

  await cancelarProgramadasTarjetas();

  const tcs = state?.tarjetasCredito || [];
  if (!Array.isArray(tcs) || tcs.length === 0) return;

  const moneda = String(state?.moneda || '').trim();
  const ahora = new Date();

  for (let offset = 0; offset <= HORIZONTE_DIAS; offset += 1) {
    const refDia = new Date(ahora.getFullYear(), ahora.getMonth(), ahora.getDate() + offset, 12, 0, 0);
    const r = resumenAlertasTarjetasCredito(state, refDia);

    for (const t of r.tarjetas || []) {
      if (!t || t.pagoCorteMuestraAlDia) continue;

      const tRaw = tcs.find((x) => x && x.id === t.id);
      let pagoSug = '';
      if (tRaw) {
        try {
          const ex = construirExtractoBancarioTarjeta(tRaw, state, refDia);
          pagoSug = `${formatearNumero(montoPagoSugeridoDesdeExtracto(ex), 0)}${moneda ? ` ${moneda}` : ''}`;
        } catch {
          /* ignorar */
        }
      }

      const y = refDia.getFullYear();
      const mo = refDia.getMonth() + 1;
      const da = refDia.getDate();
      const nom = String(t.nombreEntidad || 'Tarjeta').trim();
      const ctxTc = { nom, pagoSug };

      const programar = async (identifier, titulo, cuerpo, tipo, extra = {}) => {
        const ideal = diaHoraAviso(refDia, ahora);
        const triggerAt = ajustarDisparadorSiYaPasaronLasNueve(refDia, ahora, ideal);
        if (triggerAt == null) return;
        const disparador = construirDisparador(triggerAt, CANAL_TARJETAS);
        if (disparador == null) return;
        try {
          await Notifications.scheduleNotificationAsync({
            identifier,
            content: {
              title: titulo,
              body: cuerpo,
              data: { tipo, tarjetaId: t.id, ...extra },
              sound: true,
              ...(Platform.OS === 'android' && {
                channelId: CANAL_TARJETAS,
                priority: 'high',
              }),
            },
            trigger: disparador,
          });
        } catch (e) {
          console.warn('[MoneyTrack] scheduleNotificationAsync TC:', e?.message || e);
        }
      };

      if (t.alertaCorte) {
        const dc = t.diasCorte;
        if (dc === 0 || t.corteHoy) {
          const v = varianteNotifTcCorteHoy(ctxTc);
          await programar(`tc-corte-hoy-${t.id}-${y}-${mo}-${da}`, v.title, v.body, TIPO_TC_CORTE, {
            aviso: 'corte_hoy',
          });
        } else if (dc === 1) {
          const v = varianteNotifTcCorteManana(ctxTc);
          await programar(`tc-corte-d1-${t.id}-${y}-${mo}-${da}`, v.title, v.body, TIPO_TC_CORTE, {
            aviso: 'corte_manana',
          });
        } else if (dc === 2) {
          const v = varianteNotifTcCorte2d(ctxTc);
          await programar(`tc-corte-d2-${t.id}-${y}-${mo}-${da}`, v.title, v.body, TIPO_TC_CORTE, {
            aviso: 'corte_2d',
          });
        }
      }

      if (t.alertaPagoUrgente && t.diasPago >= 0 && t.diasPago <= 3) {
        const dp = t.diasPago;
        const v = varianteNotifTcPago(dp, ctxTc);
        await programar(`tc-pago-d${dp}-${t.id}-${y}-${mo}-${da}`, v.title, v.body, TIPO_TC_PAGO, {
          diasPago: dp,
        });
      }
    }
  }
}

/**
 * Programa una notificación local en ~15 s para comprobar permisos y canal con la app en segundo plano o cerrada.
 */
export async function programarNotificacionLocalDePrueba() {
  if (Platform.OS === 'web') {
    return { ok: false, mensaje: 'En web no hay notificaciones nativas.' };
  }
  if (!notificacionesSistemaDisponibles()) {
    return {
      ok: false,
      mensaje:
        'Expo Go no activa notificaciones del sistema. Usa un development build (`eas build`) o instala el APK/AAB generado.',
    };
  }
  const ok = await asegurarCanalYPermisos();
  if (!ok) {
    return { ok: false, mensaje: 'Activa notificaciones para MoneyTrack en los ajustes del teléfono.' };
  }
  try {
    const list = await Notifications.getAllScheduledNotificationsAsync();
    await Promise.all(
      list
        .filter((r) => r?.content?.data?.tipo === TIPO_PRUEBA)
        .map((r) => Notifications.cancelScheduledNotificationAsync(r.identifier))
    );
    const trigger =
      Platform.OS === 'android'
        ? {
            type: Notifications.SchedulableTriggerInputTypes.TIME_INTERVAL,
            seconds: 20,
            channelId: CANAL_ANDROID,
          }
        : {
            type: Notifications.SchedulableTriggerInputTypes.DATE,
            date: new Date(Date.now() + 20_000),
          };
    const prueba = varianteNotifPruebaSistema();
    await Notifications.scheduleNotificationAsync({
      identifier: 'moneytrack-prueba-local',
      content: {
        title: prueba.title,
        body: prueba.body,
        data: { tipo: TIPO_PRUEBA },
        sound: true,
        ...(Platform.OS === 'android' && {
          channelId: CANAL_ANDROID,
          priority: 'high',
        }),
      },
      trigger,
    });
    return {
      ok: true,
      mensaje:
        'Listo: en unos 20 s debería aparecer el aviso. Puedes minimizar o cerrar MoneyTrack y esperar.',
    };
  } catch (e) {
    return { ok: false, mensaje: e?.message || String(e) };
  }
}

/**
 * Resumen para pantalla Administrar: entorno, permiso, cuántas hay programadas en el SO.
 */
export async function diagnosticarNotificacionesLocales(state) {
  if (Platform.OS === 'web') {
    return {
      texto: 'En la versión web no hay notificaciones nativas.',
      permiso: 'n/a',
      pendientes: 0,
      pendientesPagos: 0,
      expoGo: false,
      activos: 0,
    };
  }
  const expoGo = entornoEsExpoGo();
  let permiso = 'desconocido';
  try {
    const r = await Notifications.getPermissionsAsync();
    permiso = r.status;
  } catch {
    permiso = 'error';
  }
  let pendientes = 0;
  let pendientesPagos = 0;
  let listProgramadas = [];
  try {
    listProgramadas = await Notifications.getAllScheduledNotificationsAsync();
    pendientes = listProgramadas.length;
    pendientesPagos = listProgramadas.filter((x) => x?.content?.data?.tipo === TIPO_DATA).length;
  } catch {
    pendientes = -1;
    listProgramadas = [];
  }
  const activos = (state?.pagosProgramados || []).filter((p) => p && p.activo !== false).length;
  const lineas = [
    expoGo
      ? '• Expo Go: Android/iOS no programan estas alarmas en la barra.'
      : '• Build instalada / dev client: entorno correcto para alarmas.',
    `• Permiso de notificaciones: ${permiso}`,
    !expoGo && permiso !== 'granted'
      ? '• Si no es «granted»: Ajustes del teléfono → Apps → MoneyTrack → Notificaciones (activar).'
      : null,
    Platform.OS === 'android' && !expoGo
      ? '• Android: revisa también energía sin restricciones y «Alarmas y recordatorios» para MoneyTrack.'
      : null,
    `• Avisos pendientes en el sistema: ${pendientes}${pendientes >= 0 ? ` (${pendientesPagos} de pagos programados)` : ' (no se pudo leer)'}`,
    `• Pagos programados activos guardados: ${activos}`,
    '• Tras cambiar pagos, deja la app abierta unos segundos o vuelve a entrar para reprogramar.',
    ...(pendientes > 0 && listProgramadas.length > 0
      ? ['• La siguiente alarma (las 9 pueden ser días distintos; no suenan todas a la vez):', ...lineasProximosDisparos(listProgramadas)]
      : []),
  ].filter(Boolean);
  return {
    texto: lineas.join('\n'),
    permiso,
    pendientes,
    pendientesPagos,
    expoGo,
    activos,
  };
}

/** Cuántas notificaciones locales hay pendientes (pagos, TC, pruebas, etc.). */
export async function contarNotificacionesLocalesProgramadas() {
  if (Platform.OS === 'web') return 0;
  try {
    const list = await Notifications.getAllScheduledNotificationsAsync();
    return list.length;
  } catch {
    return -1;
  }
}
