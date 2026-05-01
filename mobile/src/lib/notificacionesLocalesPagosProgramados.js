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
  normalizarMeta,
  resumenAlertasTarjetasCredito,
  construirExtractoBancarioTarjeta,
  montoPagoSugeridoDesdeExtracto,
  resumenPresupuestoMensualParaNotificacion,
} from './finance';
import { tituloNotifConNombre } from './notificacionesPersonalizacion';
import { solicitarPermisosSiNoConcedidos } from './notificacionesPermisos';
import {
  varianteNotifPagoProgramado,
  varianteNotifPagoPruebaLejano,
  varianteNotifPruebaSistema,
  varianteNotifTcCorteHoy,
  varianteNotifTcCorteManana,
  varianteNotifTcCorte2d,
  varianteNotifTcPago,
  varianteNotifPresupuestoInforme,
  varianteNotifPresupuestoAlerta,
} from './notificacionesVariantesAmigables';

/** `-v2`: en Android los canales no se pueden “subir” de importancia; nuevo id fuerza cabecera/sonido correctos. */
const CANAL_ANDROID = 'pagos-programados-v2';
const CANAL_TARJETAS = 'tarjetas-credito-v2';
const CANAL_PRESUPUESTO = 'presupuesto-mensual-v1';
const CANAL_METAS = 'metas-objetivos-v1';
const TIPO_DATA = 'pagoProgramado';
const TIPO_PRESUPUESTO = 'presupuestoMensualLocal';
const TIPO_TC_CORTE = 'tcCorteLocal';
const TIPO_TC_PAGO = 'tcPagoLocal';
const TIPO_PRUEBA = 'notifPruebaMoneyTrack';
export const TIPO_META_LOCAL = 'metaObjetivoLocal';
/** Mismo rango que la campana: hoy (0) y 1–3 días antes del vencimiento. */
const DIAS_MAX = 3;
/** Días hacia adelante en que programamos avisos (incluye varios ciclos mensuales). */
const HORIZONTE_DIAS = 45;
/** Recordatorios de metas: un disparo al día por meta (primer slot Bogotá). */
const HORIZONTE_METAS_DIAS = 14;

/** Zona horaria fija para recordatorios (Colombia, sin DST). */
const TZ_BOGOTA = 'America/Bogota';
/** Franjas fijas (Bogotá), en punto — respetuosas: solo 9:00, 12:00 y 18:00. */
const HORAS_AVISO_BOGOTA = [9, 12, 18];
const MINUTOS_TRAS_LA_HORA_EN_PUNTO = 0;
/**
 * El mismo día en Bogotá, un único aviso de comprobación ~5 min después de sincronizar,
 * solo si cae *antes* del primer disparo de franja (evita duplicar y evita re-programar +5 en cada apertura).
 */
const MINUTOS_PRUEBA_TRAS_SINCRO = 5;
/**
 * `true`: si no hay vencimientos en ventana 0–3 días pero sí pagos activos, una sola notificación de prueba
 * a los 5 min (no una por cada pago).
 */
export const AVISOS_LOCALES_USAR_HORA_CERCANA = true;
const MINUTOS_PRUEBA_LEJANO = 5;

let canalesAndroidListos = false;

/** Fecha calendario (y, mo, d) en Bogotá para el instante `ahora`. */
function ymdBogota(ahora) {
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: TZ_BOGOTA,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
  }).formatToParts(ahora);
  const y = parseInt(parts.find((x) => x.type === 'year').value, 10);
  const mo = parseInt(parts.find((x) => x.type === 'month').value, 10);
  const d = parseInt(parts.find((x) => x.type === 'day').value, 10);
  return { y, mo, d };
}

/** Suma días al calendario Bogotá de `desde` (números y/m/d coherentes con Date.UTC). */
export function ymdBogotaMasDias(desde, deltaDias) {
  const { y, mo, d } = ymdBogota(desde);
  const t = Date.UTC(y, mo - 1, d + deltaDias);
  const nd = new Date(t);
  return { y: nd.getUTCFullYear(), mo: nd.getUTCMonth() + 1, d: nd.getUTCDate() };
}

function esMismoDiaBogota(y, mo, d, ahora) {
  const h = ymdBogota(ahora);
  return h.y === y && h.mo === mo && h.d === d;
}

/**
 * Instante UTC para una hora civil en Bogotá (UTC−5 fijo).
 * Ej. 9:00 Bogotá → 14:00 UTC el mismo día civil.
 */
function instanteWallBogota(y, month, day, hourBOG, minuteBOG) {
  return new Date(Date.UTC(y, month - 1, day, hourBOG + 5, minuteBOG, 0, 0));
}

/**
 * Mediodía **local** en el dispositivo con año/mes/día Bogotá — para `diasHastaPagoProgramado` (calendario del usuario; en CO coincide con Bogotá).
 */
function refMediodiaParaCalculoDias(y, mo, d) {
  return new Date(y, mo - 1, d, 12, 0, 0, 0);
}

/**
 * Instantes ese día Bogotá: 9:00, 12:00 y 18:00 que aún no pasaron (hora Colombia).
 * Si es hoy en Bogotá y queda al menos una franja: opcionalmente se antepone un aviso ~5 min
 * tras la sincronización, solo si es anterior al primer disparo de franja (prueba de canal).
 * Si ya pasaron las tres franjas hoy: no se programa nada ese día (no hay +5 min en cada reapertura).
 */
export function instantesAvisoDiaBogota(y, mo, d, ahora) {
  const candidatos = HORAS_AVISO_BOGOTA.map((hh) =>
    instanteWallBogota(y, mo, d, hh, MINUTOS_TRAS_LA_HORA_EN_PUNTO)
  );
  const futuros = candidatos.filter((t) => t.getTime() > ahora.getTime());

  if (!esMismoDiaBogota(y, mo, d, ahora)) {
    return futuros;
  }

  const merged = [];
  if (futuros.length > 0) {
    const prueba = new Date(ahora.getTime() + MINUTOS_PRUEBA_TRAS_SINCRO * 60_000);
    const primero = futuros[0];
    if (prueba.getTime() < primero.getTime() && prueba.getTime() > ahora.getTime() + 14_000) {
      merged.push(prueba);
    }
    merged.push(...futuros);
  }
  return merged;
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
    await Notifications.setNotificationChannelAsync(CANAL_PRESUPUESTO, {
      name: 'Presupuesto mensual',
      importance: AndroidImportance.HIGH,
      vibrationPattern: [0, 250, 250, 250],
      lightColor: '#2DD4BF',
      enableVibrate: true,
    });
    await Notifications.setNotificationChannelAsync(CANAL_METAS, {
      name: 'Metas y ahorro',
      importance: AndroidImportance.HIGH,
      vibrationPattern: [0, 250, 250, 250],
      lightColor: '#34D399',
      enableVibrate: true,
    });
    canalesAndroidListos = true;
  }
  const ok = await solicitarPermisosSiNoConcedidos();
  if (!ok) {
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

async function cancelarProgramadasPresupuesto() {
  const list = await Notifications.getAllScheduledNotificationsAsync();
  await Promise.all(
    list
      .filter((r) => r?.content?.data?.tipo === TIPO_PRESUPUESTO)
      .map((r) => Notifications.cancelScheduledNotificationAsync(r.identifier))
  );
}

async function cancelarProgramadasMetasLocales() {
  const list = await Notifications.getAllScheduledNotificationsAsync();
  await Promise.all(
    list
      .filter((r) => r?.content?.data?.tipo === TIPO_META_LOCAL)
      .map((r) => Notifications.cancelScheduledNotificationAsync(r.identifier))
  );
}

function ctxParaVariantesPresupuesto(r) {
  const moneda = r.moneda;
  const mk = (n) => (moneda ? `${formatearNumero(n)} ${moneda}` : formatearNumero(n));
  const pct = Math.min(999, Math.round(r.pctUsado));
  return {
    ...r,
    pct,
    topeLine: mk(r.presupuestoMensual),
    gastadoLine: mk(r.gastosMesActual),
    disponibleLine: mk(r.disponible),
    excesoLine: mk(Math.max(0, -r.disponible)),
  };
}

/**
 * Recordatorios del tope mensual (mismas franjas Bogotá que pagos/TC).
 * Solo si hay presupuesto > 0; datos del mes calendario actual al sincronizar.
 */
export async function sincronizarNotificacionesLocalesPresupuesto(state) {
  if (Platform.OS === 'web') return;
  const ok = await asegurarCanalYPermisos();
  if (!ok) return;

  await cancelarProgramadasPresupuesto();

  const res = resumenPresupuestoMensualParaNotificacion(state, new Date());
  if (!res.activo) return;

  const ctx = ctxParaVariantesPresupuesto(res);
  const usarInforme = res.estadoKind === 'ok';
  const ahora = new Date();
  const { y, mo, d } = ymdBogota(ahora);
  const instantes = instantesAvisoDiaBogota(y, mo, d, ahora);

  for (let slotIx = 0; slotIx < instantes.length; slotIx += 1) {
    const triggerAt = instantes[slotIx];
    const disparador = construirDisparador(triggerAt, CANAL_PRESUPUESTO);
    if (disparador == null) continue;

    const { title: titulo, body: cuerpo } = usarInforme
      ? varianteNotifPresupuestoInforme(ctx)
      : varianteNotifPresupuestoAlerta(ctx);
    const tituloPers = tituloNotifConNombre(state, titulo);

    try {
      await Notifications.scheduleNotificationAsync({
        identifier: `mt-presup-${y}-${mo}-${d}-s${slotIx}`,
        content: {
          title: tituloPers,
          body: cuerpo,
          data: {
            tipo: TIPO_PRESUPUESTO,
            estadoKind: res.estadoKind,
            pct: ctx.pct,
          },
          sound: true,
          ...(Platform.OS === 'android' && {
            channelId: CANAL_PRESUPUESTO,
            priority: 'high',
          }),
        },
        trigger: disparador,
      });
    } catch (e) {
      console.warn('[MoneyTrack] scheduleNotificationAsync presupuesto:', e?.message || e);
    }
  }
}

/**
 * Programa notificaciones locales en franjas Bogotá (9:00, 12:00, 18:00; opcional ~5 min tras sync el mismo día si aplica) cada día en que falten 0–3 días para un pago activo.
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
      const { y, mo, d } = ymdBogotaMasDias(ahora, offset);
      const refMediodia = refMediodiaParaCalculoDias(y, mo, d);
      const dDays = diasHastaPagoProgramado(p, refMediodia);
      if (dDays == null || dDays < 0 || dDays > DIAS_MAX) continue;

      const instantes = instantesAvisoDiaBogota(y, mo, d, ahora);
      for (let slotIx = 0; slotIx < instantes.length; slotIx += 1) {
        const triggerAt = instantes[slotIx];
        const disparador = construirDisparador(triggerAt);
        if (disparador == null) continue;

        const concepto = String(p.concepto || 'Pago').trim() || 'Pago programado';
        const montoStr = formatearNumero(p.monto);
        const { title: titulo, body: cuerpo } = varianteNotifPagoProgramado(dDays, {
          concepto,
          montoStr,
          moneda,
        });
        const tituloPers = tituloNotifConNombre(state, titulo);
        const idBase = String(p.id != null ? p.id : 'sin-id');
        const idClave = `pp-loc-${idBase}-${y}-${mo}-${d}-s${slotIx}`;

        try {
          await Notifications.scheduleNotificationAsync({
            identifier: idClave,
            content: {
              title: tituloPers,
              body: cuerpo,
              data: { tipo: TIPO_DATA, pagoId: p.id, dias: dDays },
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
  }

  /**
   * Con `AVISOS_LOCALES_USAR_HORA_CERCANA`, si no hubo ninguna programación (vencimiento a más de 3 días),
   * una sola notificación de prueba a los 5 min (primer pago elegible) — no una por cada pago.
   */
  if (AVISOS_LOCALES_USAR_HORA_CERCANA && programadas === 0) {
    const hoyBog = ymdBogota(ahora);
    const refHoyMediodia = refMediodiaParaCalculoDias(hoyBog.y, hoyBog.mo, hoyBog.d);
    for (const p of pagos) {
      if (!p || p.activo === false) continue;
      const dHoy = diasHastaPagoProgramado(p, refHoyMediodia);
      if (dHoy == null || dHoy < 0 || dHoy <= DIAS_MAX || dHoy > 366) continue;

      const ideal = new Date(ahora.getTime() + MINUTOS_PRUEBA_LEJANO * 60_000);
      const disparador = construirDisparador(ideal);
      if (disparador == null) break;

      const concepto = String(p.concepto || 'Pago').trim() || 'Pago programado';
      const montoStr = formatearNumero(p.monto);
      const idBase = String(p.id != null ? p.id : 'sin-id');
      const y = hoyBog.y;
      const mo = hoyBog.mo;
      const da = hoyBog.d;
      const idClave = `pp-loc-prueba-lejano-${idBase}-${y}-${mo}-${da}`;
      const pruebaTxt = varianteNotifPagoPruebaLejano({ concepto, montoStr, moneda, dHoy });
      const tituloPruebaPers = tituloNotifConNombre(state, pruebaTxt.title);

      try {
        await Notifications.scheduleNotificationAsync({
          identifier: idClave,
          content: {
            title: tituloPruebaPers,
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
      break;
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
    const { y, mo, d } = ymdBogotaMasDias(ahora, offset);
    const refMediodia = refMediodiaParaCalculoDias(y, mo, d);
    const r = resumenAlertasTarjetasCredito(state, refMediodia);

    for (const t of r.tarjetas || []) {
      if (!t || t.pagoCorteMuestraAlDia) continue;

      const tRaw = tcs.find((x) => x && x.id === t.id);
      let pagoSug = '';
      if (tRaw) {
        try {
          const ex = construirExtractoBancarioTarjeta(tRaw, state, refMediodia);
          pagoSug = `${formatearNumero(montoPagoSugeridoDesdeExtracto(ex), 0)}${moneda ? ` ${moneda}` : ''}`;
        } catch {
          /* ignorar */
        }
      }

      const nom = String(t.nombreEntidad || 'Tarjeta').trim();
      const ctxTc = { nom, pagoSug };
      if (
        !t.alertaCorte &&
        !(t.alertaPagoUrgente && t.diasPago >= 0 && t.diasPago <= 3)
      ) {
        continue;
      }

      const instantes = instantesAvisoDiaBogota(y, mo, d, ahora);
      for (let slotIx = 0; slotIx < instantes.length; slotIx += 1) {
        const triggerAt = instantes[slotIx];
        const disparador = construirDisparador(triggerAt, CANAL_TARJETAS);
        if (disparador == null) continue;

        const programar = async (identifier, titulo, cuerpo, tipo, extra = {}) => {
          try {
            await Notifications.scheduleNotificationAsync({
              identifier,
              content: {
                title: tituloNotifConNombre(state, titulo),
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
            await programar(`tc-corte-hoy-${t.id}-${y}-${mo}-${d}-s${slotIx}`, v.title, v.body, TIPO_TC_CORTE, {
              aviso: 'corte_hoy',
            });
          } else if (dc === 1) {
            const v = varianteNotifTcCorteManana(ctxTc);
            await programar(`tc-corte-d1-${t.id}-${y}-${mo}-${d}-s${slotIx}`, v.title, v.body, TIPO_TC_CORTE, {
              aviso: 'corte_manana',
            });
          } else if (dc === 2) {
            const v = varianteNotifTcCorte2d(ctxTc);
            await programar(`tc-corte-d2-${t.id}-${y}-${mo}-${d}-s${slotIx}`, v.title, v.body, TIPO_TC_CORTE, {
              aviso: 'corte_2d',
            });
          }
        }

        if (t.alertaPagoUrgente && t.diasPago >= 0 && t.diasPago <= 3) {
          const dp = t.diasPago;
          const v = varianteNotifTcPago(dp, ctxTc);
          await programar(`tc-pago-d${dp}-${t.id}-${y}-${mo}-${d}-s${slotIx}`, v.title, v.body, TIPO_TC_PAGO, {
            diasPago: dp,
          });
        }
      }
    }
  }
}

/**
 * Metas con objetivo > 0: aviso local cuando falte avanzar, cuando estés ≥80 % o al cumplir el 100 %.
 * Un disparo al día por meta (primer slot Bogotá); si ya cumpliste, un único aviso en el próximo slot válido.
 */
export async function sincronizarNotificacionesLocalesMetas(state) {
  if (Platform.OS === 'web') return;
  const ok = await asegurarCanalYPermisos();
  if (!ok) return;

  await cancelarProgramadasMetasLocales();

  const moneda = String(state?.moneda || '').trim();
  const suf = moneda ? ` ${moneda}` : '';
  const metasRaw = state?.metas || [];
  const contrib = Array.isArray(state?.contribucionesMetas) ? state.contribucionesMetas : [];
  const metas = metasRaw.map(normalizarMeta).filter((m) => m.id && (parseFloat(m.objetivo) || 0) > 0);
  if (metas.length === 0) return;

  const ahora = new Date();

  for (const m of metas) {
    const obj = parseFloat(m.objetivo) || 0;
    const acum = contrib
      .filter((c) => c && c.metaId === m.id)
      .reduce((s, c) => s + (parseFloat(c.cantidad) || 0), 0);
    const pct = obj > 0 ? Math.min(100, (acum / obj) * 100) : 0;
    const nom = String(m.nombre || 'Meta').trim() || 'Meta';
    const acumTxt = formatearNumero(acum, 0);
    const objTxt = formatearNumero(obj, 0);

    if (pct >= 100) {
      const { y, mo, d } = ymdBogota(ahora);
      const instantes = instantesAvisoDiaBogota(y, mo, d, ahora);
      const triggerAt = instantes[0];
      if (!triggerAt) continue;
      const disparador = construirDisparador(triggerAt, CANAL_METAS);
      if (disparador == null) continue;
      const tituloRaw = `¡Meta cumplida: ${nom}! 🎉`;
      const cuerpo = `Llegaste a ${acumTxt}${suf} de ${objTxt}${suf}. Gran trabajo.`;
      try {
        await Notifications.scheduleNotificationAsync({
          identifier: `mt-meta-logro-${m.id}`,
          content: {
            title: tituloNotifConNombre(state, tituloRaw),
            body: cuerpo,
            data: { tipo: TIPO_META_LOCAL, metaId: m.id, aviso: 'cumplida', pct: 100 },
            sound: true,
            ...(Platform.OS === 'android' && {
              channelId: CANAL_METAS,
              priority: 'high',
            }),
          },
          trigger: disparador,
        });
      } catch (e) {
        console.warn('[MoneyTrack] scheduleNotificationAsync meta logro:', e?.message || e);
      }
      continue;
    }

    const rol = pct >= 80 ? 'cerca' : 'recordatorio';
    for (let offset = 0; offset <= HORIZONTE_METAS_DIAS; offset += 1) {
      const { y, mo, d } = ymdBogotaMasDias(ahora, offset);
      const instantes = instantesAvisoDiaBogota(y, mo, d, ahora);
      const triggerAt = instantes[0];
      if (!triggerAt) continue;
      const disparador = construirDisparador(triggerAt, CANAL_METAS);
      if (disparador == null) continue;

      let tituloRaw;
      let cuerpo;
      if (rol === 'cerca') {
        const pRound = Math.round(pct);
        tituloRaw = `Casi logras «${nom}» · ${pRound}% 🎯`;
        cuerpo = `Llevas ${acumTxt}${suf} de ${objTxt}${suf}. Un último empujón y la cierras.`;
      } else {
        tituloRaw = `Tu meta «${nom}» te recuerda 🌿`;
        cuerpo = `Vas en ${acumTxt}${suf} de ${objTxt}${suf}. Un aporte cuando puedas en Más → Metas.`;
      }

      const idClave = `mt-meta-${rol}-${m.id}-${y}-${mo}-${d}`;

      try {
        await Notifications.scheduleNotificationAsync({
          identifier: idClave,
          content: {
            title: tituloNotifConNombre(state, tituloRaw),
            body: cuerpo,
            data: { tipo: TIPO_META_LOCAL, metaId: m.id, aviso: rol, pct: Math.round(pct) },
            sound: true,
            ...(Platform.OS === 'android' && {
              channelId: CANAL_METAS,
              priority: 'high',
            }),
          },
          trigger: disparador,
        });
      } catch (e) {
        console.warn('[MoneyTrack] scheduleNotificationAsync meta:', e?.message || e);
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
        title: tituloNotifConNombre({}, prueba.title),
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
      pendientesPresupuesto: 0,
      pendientesMetas: 0,
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
  let pendientesPresupuesto = 0;
  let pendientesMetas = 0;
  let listProgramadas = [];
  try {
    listProgramadas = await Notifications.getAllScheduledNotificationsAsync();
    pendientes = listProgramadas.length;
    pendientesPagos = listProgramadas.filter((x) => x?.content?.data?.tipo === TIPO_DATA).length;
    pendientesPresupuesto = listProgramadas.filter((x) => x?.content?.data?.tipo === TIPO_PRESUPUESTO).length;
    pendientesMetas = listProgramadas.filter((x) => x?.content?.data?.tipo === TIPO_META_LOCAL).length;
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
      ? '• Android: sin «Alarmas y recordatorios» / energía sin restricciones, los avisos programados pueden retrasarse o no sonar con la app cerrada (Ajustes → Apps → MoneyTrack).'
      : null,
    `• Avisos pendientes en el sistema: ${pendientes}${
      pendientes >= 0
        ? ` (${pendientesPagos} pagos programados${pendientesPresupuesto ? `, ${pendientesPresupuesto} presupuesto` : ''}${
            pendientesMetas ? `, ${pendientesMetas} metas` : ''
          })`
        : ' (no se pudo leer)'
    }`,
    `• Pagos programados activos guardados: ${activos}`,
    '• Tras cambiar pagos, deja la app abierta unos segundos o vuelve a entrar para reprogramar.',
    ...(pendientes > 0 && listProgramadas.length > 0
      ? [
          '• Próximas alarmas (horario Bogotá 9:00 / 12:00 / 18:00; pueden ser días distintos):',
          ...lineasProximosDisparos(listProgramadas),
        ]
      : []),
  ].filter(Boolean);
  return {
    texto: lineas.join('\n'),
    permiso,
    pendientes,
    pendientesPagos,
    pendientesPresupuesto,
    pendientesMetas,
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
