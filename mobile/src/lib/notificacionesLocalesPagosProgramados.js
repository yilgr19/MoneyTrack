import { Platform } from 'react-native';
import * as Notifications from 'expo-notifications';
import { AndroidImportance } from 'expo-notifications';
import { diasHastaPagoProgramado } from './notificacionesApp';
import { formatearNumero } from './finance';

const CANAL_ANDROID = 'pagos-programados';
const TIPO_DATA = 'pagoProgramado';
/** Mismo rango que la campana: hoy (0) y 1–3 días antes del vencimiento. */
const DIAS_MAX = 3;
/** Días hacia adelante en que programamos avisos (incluye varios ciclos mensuales). */
const HORIZONTE_DIAS = 45;
const HORA_AVISO = 9;

let canalAndroidListo = false;

function diaHoraAviso(refDia) {
  return new Date(refDia.getFullYear(), refDia.getMonth(), refDia.getDate(), HORA_AVISO, 0, 0);
}

/** Si las 9:00 de ese día ya pasaron pero sigues en ventana 1–3 días, programa en breve (para que aparezca en la barra al cerrar la app). */
function ajustarDisparadorSiYaPasaronLasNueve(refDia, ahora, triggerAt) {
  const t = triggerAt.getTime();
  if (t > ahora.getTime()) return triggerAt;
  const finDiaRef = new Date(refDia.getFullYear(), refDia.getMonth(), refDia.getDate(), 23, 59, 59, 999);
  if (ahora.getTime() > finDiaRef.getTime()) return null;
  return new Date(ahora.getTime() + 60_000);
}

/** En Android, `TIME_INTERVAL` suele ser más fiable que `DATE` para alarmas próximas. */
function construirDisparador(ahora, triggerAt) {
  const ms = triggerAt.getTime() - ahora.getTime();
  const sec = Math.ceil(ms / 1000);
  if (sec < 15) return null;
  const maxSecIntervalo = 86400 * 60;
  if (Platform.OS === 'android' && sec <= maxSecIntervalo) {
    return {
      type: Notifications.SchedulableTriggerInputTypes.TIME_INTERVAL,
      seconds: sec,
      channelId: CANAL_ANDROID,
    };
  }
  return {
    type: Notifications.SchedulableTriggerInputTypes.DATE,
    date: triggerAt,
    ...(Platform.OS === 'android' && { channelId: CANAL_ANDROID }),
  };
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
  if (Platform.OS === 'android' && !canalAndroidListo) {
    await Notifications.setNotificationChannelAsync(CANAL_ANDROID, {
      name: 'Pagos programados',
      importance: AndroidImportance.HIGH,
      vibrationPattern: [0, 250, 250, 250],
      lightColor: '#4B246C',
    });
    canalAndroidListo = true;
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

      const ideal = diaHoraAviso(refDia);
      const triggerAt = ajustarDisparadorSiYaPasaronLasNueve(refDia, ahora, ideal);
      if (triggerAt == null) continue;

      const disparador = construirDisparador(ahora, triggerAt);
      if (disparador == null) continue;

      const concepto = String(p.concepto || 'Pago').trim() || 'Pago programado';
      const montoStr = formatearNumero(p.monto);
      const plazoLargo =
        d === 0 ? 'vence hoy' : d === 1 ? '1 día restante' : `${d} días restantes`;
      const titulo =
        d === 0
          ? 'Hoy vence un pago programado'
          : d === 1
            ? 'Falta 1 día de plazo para el pago'
            : `Faltan ${d} días de plazo para el pago`;
      const cuerpo =
        `«${concepto}» — te quedan ${plazoLargo} para realizarlo.` +
        (moneda
          ? ` Monto: ${montoStr} ${moneda}. Cumple a tiempo; luego anótalo en Gastos.`
          : ` Monto: ${montoStr}. Cumple a tiempo; luego anótalo en Gastos.`);
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

  const hayActivos = pagos.some((x) => x && x.activo !== false);
  if (hayActivos && programadas === 0) {
    console.warn(
      '[MoneyTrack] No hay notificaciones de pago en la barra: el vencimiento debe estar entre hoy y 3 días vista (Pagos programados). Otras avisos de la campana no usan la barra.'
    );
  }
}
