import AsyncStorage from '@react-native-async-storage/async-storage';
import { CUENTAS } from './finance';

const KEYS = [
  'moneda',
  'saldosCuentas',
  'limiteTarjetaCredito',
  'presupuestoMensual',
  'ingresos',
  'gastos',
  'categorias',
  'metas',
  'contribucionesMetas',
  'pagosProgramados',
  'saldoEfectivo',
  'saldoBanco',
  'saldoInicialNota',
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
    ingresos: parseJson(map.ingresos, []),
    gastos: parseJson(map.gastos, []),
    categorias: parseJson(map.categorias, []),
    metas: parseJson(map.metas, []),
    contribucionesMetas: parseJson(map.contribucionesMetas, []),
    pagosProgramados: parseJson(map.pagosProgramados, []),
    saldoInicialNota: map.saldoInicialNota || '',
  };
}

export async function persistAppState(state) {
  const saldos = state.saldosCuentas || emptySaldosCuentas();
  const pairs = [
    ['moneda', state.moneda || ''],
    ['saldosCuentas', JSON.stringify(saldos)],
    ['limiteTarjetaCredito', String(parseFloat(state.limiteTarjetaCredito) || 0)],
    ['presupuestoMensual', String(parseFloat(state.presupuestoMensual) || 0)],
    ['ingresos', JSON.stringify(state.ingresos || [])],
    ['gastos', JSON.stringify(state.gastos || [])],
    ['categorias', JSON.stringify(state.categorias || [])],
    ['metas', JSON.stringify(state.metas || [])],
    ['contribucionesMetas', JSON.stringify(state.contribucionesMetas || [])],
    ['pagosProgramados', JSON.stringify(state.pagosProgramados || [])],
    ['saldoInicialNota', state.saldoInicialNota || ''],
  ];
  await AsyncStorage.multiSet(pairs);
}

export async function clearStoragePartial() {
  await AsyncStorage.multiRemove([
    'saldosCuentas',
    'limiteTarjetaCredito',
    'presupuestoMensual',
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
  await AsyncStorage.multiRemove(['moneda', 'pagosProgramados']);
  await AsyncStorage.multiSet([
    ['gastos', '[]'],
    ['ingresos', '[]'],
    ['categorias', '[]'],
    ['metas', '[]'],
    ['contribucionesMetas', '[]'],
  ]);
}
