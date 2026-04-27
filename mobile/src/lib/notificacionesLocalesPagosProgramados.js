import { Platform } from 'react-native';
import * as Notifications from 'expo-notifications';
import { AndroidImportance } from 'expo-notifications';
import { diasHastaPagoProgramado } from './notificacionesApp';
import { formatearNumero, pagoDebeMostrarseParaPagar } from './finance';

const CANAL_ANDROID = 'pagos-programados';
const TIPO_DATA = 'pagoProgramado';
/** Mismo rango que la campana: 1, 2 o 3 días antes. */
const DIAS_VENTANA = 3;
/** Días hacia adelante en que programamos avisos (incluye varios ciclos mensuales). */
const HORIZONTE_DIAS = 45;
const HORA_AVISO = 9;

let permisosPedidos = false;
let canalAndroidListo = false;

function diaHoraAviso(refDia) {
  return new Date(refDia.getFullYear(), refDia.getMonth(), refDia.getDate(), HORA_AVISO, 0, 0);
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
  if (!permisosPedidos) {
    const { status: existente } = await Notifications.getPermissionsAsync();
    let st = existente;
    if (existente !== 'granted') {
      const r = await Notifications.requestPermissionsAsync();
      st = r.status;
    }
    permisosPedidos = true;
    if (st !== 'granted') return false;
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

  for (const p of pagos) {
    if (!p || p.activo === false) continue;

    for (let offset = 0; offset <= HORIZONTE_DIAS; offset += 1) {
      const refDia = new Date(ahora.getFullYear(), ahora.getMonth(), ahora.getDate() + offset, 12, 0, 0);
      if (pagoDebeMostrarseParaPagar(p, refDia)) continue;

      const d = diasHastaPagoProgramado(p, refDia);
      if (d == null || d < 1 || d > DIAS_VENTANA) continue;

      const triggerAt = diaHoraAviso(refDia);
      if (triggerAt.getTime() <= ahora.getTime()) continue;

      const concepto = String(p.concepto || 'Pago').trim() || 'Pago programado';
      const montoStr = formatearNumero(p.monto);
      const plazoLargo = d === 1 ? '1 día restante' : `${d} días restantes`;
      const titulo =
        d === 1
          ? 'Falta 1 día de plazo para el pago'
          : `Faltan ${d} días de plazo para el pago`;
      const cuerpo =
        `«${concepto}» — te quedan ${plazoLargo} para realizarlo.` +
        (moneda
          ? ` Monto: ${montoStr} ${moneda}. Cumple a tiempo; luego anótalo en Gastos.`
          : ` Monto: ${montoStr}. Cumple a tiempo; luego anótalo en Gastos.`);
      const idClave = `pp-loc-${p.id}-${refDia.getFullYear()}-${refDia.getMonth() + 1}-${refDia.getDate()}`;

      try {
        await Notifications.scheduleNotificationAsync({
          identifier: idClave,
          content: {
            title: titulo,
            body: cuerpo,
            data: { tipo: TIPO_DATA, pagoId: p.id, dias: d },
            sound: true,
          },
          trigger: {
            type: Notifications.SchedulableTriggerInputTypes.DATE,
            date: triggerAt,
            ...(Platform.OS === 'android' && { channelId: CANAL_ANDROID }),
          },
        });
      } catch {
        /* no bloquear la app si el SO rechaza una fecha concreta */
      }
    }
  }
}
