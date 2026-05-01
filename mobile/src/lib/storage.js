import AsyncStorage from '@react-native-async-storage/async-storage';
import { NOTIFICACIONES_LECTURA_KEY } from './notificacionesLectura';
import { CUENTAS } from './finance';

const KEY_ONBOARDING_COMPLETADO = 'onboardingCompletado';

const KEYS = [
  'moneda',
  'saldosCuentas',
  'bancosDetalle',
  'plataformasDetalle',
  'tarjetasCredito',
  'limiteTarjetaCredito',
  'presupuestoMensual',
  'presupuestoDesdeFecha',
  'ingresos',
  'gastos',
  'categorias',
  'metas',
  'contribucionesMetas',
  'pagosProgramados',
  'saldoEfectivo',
  'saldoBanco',
  'saldoInicialNota',
  'extractosTarjetasHistorial',
  'bolsillos',
  'recordatoriosPagoRegistrado',
  'intencionesCompra',
  'asistenteUmbral48h',
  'listaSuperCategoriaPreferida',
  'listaSuperArticulosExtra',
  'listaSuperCompraItems',
  'avisosGastosMovimiento',
];

export function emptySaldosCuentas() {
  return CUENTAS.reduce((acc, c) => {
    acc[c.id] = 0;
    return acc;
  }, {});
}

export async function loadAppState() {
  const pairs = await AsyncStorage.multiGet(KEYS);
  const map = Object.fromEntries(pairs);
  let saldosCuentas;
  try {
    saldosCuentas = map.saldosCuentas ? JSON.parse(map.saldosCuentas) : null;
  } catch {
    saldosCuentas = null;
  }
  const parseJson = (s, fallback) => {
    try {
      return s ? JSON.parse(s) : fallback;
    } catch {
      return fallback;
    }
  };
  return {
    moneda: map.moneda || '',
    saldosCuentas: saldosCuentas || undefined,
    saldoEfectivo: map.saldoEfectivo,
    saldoBanco: map.saldoBanco,
    limiteTarjetaCredito: map.limiteTarjetaCredito || '0',
    presupuestoMensual: map.presupuestoMensual || '0',
    presupuestoDesdeFecha: map.presupuestoDesdeFecha || '',
    ingresos: parseJson(map.ingresos, []),
    gastos: parseJson(map.gastos, []),
    categorias: parseJson(map.categorias, []),
    metas: parseJson(map.metas, []),
    contribucionesMetas: parseJson(map.contribucionesMetas, []),
    pagosProgramados: parseJson(map.pagosProgramados, []),
    saldoInicialNota: map.saldoInicialNota || '',
    bancosDetalle: parseJson(map.bancosDetalle, []),
    plataformasDetalle: parseJson(map.plataformasDetalle, []),
    tarjetasCredito: parseJson(map.tarjetasCredito, []),
    extractosTarjetasHistorial: parseJson(map.extractosTarjetasHistorial, []),
    bolsillos: parseJson(map.bolsillos, []),
    recordatoriosPagoRegistrado: parseJson(map.recordatoriosPagoRegistrado, []),
    intencionesCompra: parseJson(map.intencionesCompra, []),
    asistenteUmbral48h: map.asistenteUmbral48h,
    listaSuperCategoriaPreferida: map.listaSuperCategoriaPreferida || '',
    listaSuperArticulosExtra: parseJson(map.listaSuperArticulosExtra, []),
    listaSuperCompraItems: parseJson(map.listaSuperCompraItems, []),
    avisosGastosMovimiento: parseJson(map.avisosGastosMovimiento, []),
  };
}

export async function persistAppState(state) {
  const saldos = state.saldosCuentas || emptySaldosCuentas();
  const pairs = [
    ['moneda', state.moneda || ''],
    ['saldosCuentas', JSON.stringify(saldos)],
    ['limiteTarjetaCredito', String(parseFloat(state.limiteTarjetaCredito) || 0)],
    ['presupuestoMensual', String(parseFloat(state.presupuestoMensual) || 0)],
    ['presupuestoDesdeFecha', String(state.presupuestoDesdeFecha || '').trim().slice(0, 10)],
    ['ingresos', JSON.stringify(state.ingresos || [])],
    ['gastos', JSON.stringify(state.gastos || [])],
    ['categorias', JSON.stringify(state.categorias || [])],
    ['metas', JSON.stringify(state.metas || [])],
    ['contribucionesMetas', JSON.stringify(state.contribucionesMetas || [])],
    ['pagosProgramados', JSON.stringify(state.pagosProgramados || [])],
    ['saldoInicialNota', state.saldoInicialNota || ''],
    ['bancosDetalle', JSON.stringify(state.bancosDetalle || [])],
    ['plataformasDetalle', JSON.stringify(state.plataformasDetalle || [])],
    ['tarjetasCredito', JSON.stringify(state.tarjetasCredito || [])],
    ['extractosTarjetasHistorial', JSON.stringify(state.extractosTarjetasHistorial || [])],
    ['bolsillos', JSON.stringify(state.bolsillos || [])],
    ['recordatoriosPagoRegistrado', JSON.stringify(state.recordatoriosPagoRegistrado || [])],
    ['intencionesCompra', JSON.stringify(state.intencionesCompra || [])],
    ['asistenteUmbral48h', String(parseFloat(state.asistenteUmbral48h) || 50)],
    ['listaSuperCategoriaPreferida', state.listaSuperCategoriaPreferida || ''],
    ['listaSuperArticulosExtra', JSON.stringify(state.listaSuperArticulosExtra || [])],
    ['listaSuperCompraItems', JSON.stringify(state.listaSuperCompraItems || [])],
    ['avisosGastosMovimiento', JSON.stringify(state.avisosGastosMovimiento || [])],
  ];
  await AsyncStorage.multiSet(pairs);
}

export async function clearStoragePartial() {
  await AsyncStorage.multiRemove([
    'saldosCuentas',
    'bancosDetalle',
    'plataformasDetalle',
    'tarjetasCredito',
    'limiteTarjetaCredito',
    'presupuestoMensual',
    'presupuestoDesdeFecha',
    'saldoEfectivo',
    'saldoBanco',
    'saldoInicialNota',
  ]);
  await AsyncStorage.multiSet([
    ['gastos', '[]'],
    ['ingresos', '[]'],
    ['metas', '[]'],
    ['contribucionesMetas', '[]'],
  ]);
}

export async function clearStorageFull() {
  await clearStoragePartial();
  await AsyncStorage.multiRemove([
    'moneda',
    'pagosProgramados',
    'recordatoriosPagoRegistrado',
    KEY_ONBOARDING_COMPLETADO,
    NOTIFICACIONES_LECTURA_KEY,
  ]);
  await AsyncStorage.multiSet([
    ['gastos', '[]'],
    ['ingresos', '[]'],
    ['categorias', '[]'],
    ['metas', '[]'],
    ['contribucionesMetas', '[]'],
    ['extractosTarjetasHistorial', '[]'],
    ['bolsillos', '[]'],
    ['intencionesCompra', '[]'],
    ['listaSuperArticulosExtra', '[]'],
    ['listaSuperCompraItems', '[]'],
    ['avisosGastosMovimiento', '[]'],
    ['asistenteUmbral48h', '50'],
    ['listaSuperCategoriaPreferida', ''],
  ]);
}

export async function loadOnboardingCompletado() {
  try {
    const v = await AsyncStorage.getItem(KEY_ONBOARDING_COMPLETADO);
    return v === '1';
  } catch {
    return false;
  }
}

export async function setOnboardingCompletado() {
  await AsyncStorage.setItem(KEY_ONBOARDING_COMPLETADO, '1');
}

/** Tras importar un respaldo donde el tutorial aún no estaba completado. */
export async function clearOnboardingCompletado() {
  await AsyncStorage.removeItem(KEY_ONBOARDING_COMPLETADO);
}
