import React, { createContext, useCallback, useContext, useEffect, useMemo, useRef, useState } from 'react';
import { AppState } from 'react-native';
import {
  loadAppState,
  persistAppState,
  emptySaldosCuentas,
  clearStoragePartial,
  clearStorageFull,
  loadOnboardingCompletado,
  setOnboardingCompletado,
} from '../lib/storage';
import { reemplazarPagosRecordatorioTarjetas } from '../lib/finance';
import { notificacionesSistemaDisponibles } from '../lib/notificacionesLocalesEntorno';

function normalizeState(raw) {
  const saldos = raw.saldosCuentas
    ? { ...emptySaldosCuentas(), ...raw.saldosCuentas }
    : emptySaldosCuentas();
  if (!raw.saldosCuentas) {
    if (raw.saldoEfectivo) saldos.efectivo = parseFloat(raw.saldoEfectivo) || 0;
    if (raw.saldoBanco) saldos.banco = parseFloat(raw.saldoBanco) || 0;
  }
  let bancosDetalle = Array.isArray(raw.bancosDetalle) ? raw.bancosDetalle.filter((r) => r && typeof r === 'object') : [];
  if (!bancosDetalle.length && saldos.banco > 0) {
    bancosDetalle = [{ id: `mig-${Date.now()}`, nombre: 'Cuenta bancaria', saldo: saldos.banco }];
  }
  return {
    moneda: raw.moneda || '',
    saldosCuentas: saldos,
    bancosDetalle,
    limiteTarjetaCredito: parseFloat(raw.limiteTarjetaCredito) || 0,
    presupuestoMensual: parseFloat(raw.presupuestoMensual) || 0,
    ingresos: raw.ingresos || [],
    gastos: raw.gastos || [],
    categorias: raw.categorias || [],
    metas: raw.metas || [],
    contribucionesMetas: raw.contribucionesMetas || [],
    pagosProgramados: raw.pagosProgramados || [],
    saldoInicialNota: raw.saldoInicialNota || '',
    plataformasDetalle: Array.isArray(raw.plataformasDetalle)
      ? raw.plataformasDetalle.filter((r) => r && typeof r === 'object')
      : [],
    tarjetasCredito: Array.isArray(raw.tarjetasCredito)
      ? raw.tarjetasCredito.filter((r) => r && typeof r === 'object')
      : [],
    extractosTarjetasHistorial: Array.isArray(raw.extractosTarjetasHistorial)
      ? raw.extractosTarjetasHistorial.filter((r) => r && typeof r === 'object' && r.id)
      : [],
  };
}

/**
 * Indicios de que el usuario ya empezó (moneda, movimientos, metas, presupuesto, filas de banco, etc.).
 * Si la plata se acabó, sigue en true: el Inicio no vuelve a la pantalla mínima “Bienvenido”.
 */
export function tieneDatosPrevios(n) {
  if (!n || typeof n !== 'object') return false;
  if (String(n.moneda || '').trim()) return true;
  const saldos = n.saldosCuentas || {};
  if (Object.values(saldos).some((v) => (parseFloat(v) || 0) > 0)) return true;
  if (Array.isArray(n.gastos) && n.gastos.length > 0) return true;
  if (Array.isArray(n.ingresos) && n.ingresos.length > 0) return true;
  if (Array.isArray(n.categorias) && n.categorias.length > 0) return true;
  if (Array.isArray(n.metas) && n.metas.length > 0) return true;
  if (String(n.saldoInicialNota || '').trim()) return true;
  if (parseFloat(n.presupuestoMensual) > 0) return true;
  if (Array.isArray(n.bancosDetalle) && n.bancosDetalle.length > 0) return true;
  if (Array.isArray(n.plataformasDetalle) && n.plataformasDetalle.length > 0) return true;
  if (Array.isArray(n.tarjetasCredito) && n.tarjetasCredito.length > 0) return true;
  if (parseFloat(n.limiteTarjetaCredito) > 0) return true;
  return false;
}

const AppContext = createContext(null);

export function AppProvider({ children }) {
  const [state, setState] = useState(null);
  const stateRef = useRef(null);
  const [ready, setReady] = useState(false);
  const [mostrarOnboarding, setMostrarOnboarding] = useState(false);
  /** Tras el primer tutorial: abrir pestaña Saldo para que el usuario digite saldos iniciales allí. */
  const [postOnboardingIrASaldo, setPostOnboardingIrASaldo] = useState(false);

  useEffect(() => {
    stateRef.current = state;
  }, [state]);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      const [raw, flagOnboarding] = await Promise.all([loadAppState(), loadOnboardingCompletado()]);
      const normalized = normalizeState(raw);
      let onboardingHecho = flagOnboarding;
      if (!onboardingHecho && tieneDatosPrevios(normalized)) {
        await setOnboardingCompletado();
        onboardingHecho = true;
      }
      if (!cancelled) {
        setState(normalized);
        setMostrarOnboarding(!onboardingHecho);
        setReady(true);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  const replaceState = useCallback((updater) => {
    setState((prev) => {
      if (!prev) return prev;
      const next = typeof updater === 'function' ? updater(prev) : updater;
      persistAppState(next).catch(() => {});
      return next;
    });
  }, []);

  const lenG = state?.gastos?.length ?? 0;
  const nTc = state?.tarjetasCredito?.length ?? 0;
  const cupoHash =
    nTc > 0
      ? (state?.tarjetasCredito || [])
          .map((t) => `${t.id}:${t.cupoUtilizado || 0}:${t.cupoTotal || 0}`)
          .join('|')
      : '';
  /** Regenerar recordatorios TC si cambian fechas de corte/pago (antes solo cupo y no se volvían a crear). */
  const tcFechasHash =
    nTc > 0
      ? (state?.tarjetasCredito || [])
          .map((t) => `${String(t.fechaHoraCorte || '')}|${String(t.fechaHoraLimitePago || '')}`)
          .join('||')
      : '';
  useEffect(() => {
    if (!ready || !state) return;
    replaceState((s) => ({
      ...s,
      pagosProgramados: reemplazarPagosRecordatorioTarjetas(s.pagosProgramados || [], s, new Date()),
    }));
  }, [ready, lenG, nTc, cupoHash, tcFechasHash, replaceState, state?.moneda]);

  useEffect(() => {
    if (!ready || !state || mostrarOnboarding) return;
    if (!notificacionesSistemaDisponibles()) return;
    import('../lib/notificacionesLocalesPagosProgramados').then((m) =>
      m.sincronizarNotificacionesLocalesPagosProgramados(state).catch(() => {})
    );
  }, [ready, mostrarOnboarding, state]);

  useEffect(() => {
    if (!ready || mostrarOnboarding) return;
    if (!notificacionesSistemaDisponibles()) return;
    const sub = AppState.addEventListener('change', (next) => {
      if (next === 'active' && stateRef.current) {
        import('../lib/notificacionesLocalesPagosProgramados').then((m) =>
          m.sincronizarNotificacionesLocalesPagosProgramados(stateRef.current).catch(() => {})
        );
      }
    });
    return () => sub.remove();
  }, [ready, mostrarOnboarding]);

  const clearPostOnboardingIrASaldo = useCallback(() => {
    setPostOnboardingIrASaldo(false);
  }, []);

  const completarOnboarding = useCallback(async () => {
    await setOnboardingCompletado();
    setPostOnboardingIrASaldo(true);
    setMostrarOnboarding(false);
  }, []);

  const resetPartial = useCallback(async () => {
    await clearStoragePartial();
    const raw = await loadAppState();
    setState(normalizeState(raw));
  }, []);

  const resetFull = useCallback(async () => {
    await clearStorageFull();
    const raw = await loadAppState();
    setState(normalizeState(raw));
  }, []);

  const value = useMemo(
    () => ({
      state,
      ready,
      mostrarOnboarding,
      postOnboardingIrASaldo,
      clearPostOnboardingIrASaldo,
      completarOnboarding,
      replaceState,
      resetPartial,
      resetFull,
    }),
    [
      state,
      ready,
      mostrarOnboarding,
      postOnboardingIrASaldo,
      clearPostOnboardingIrASaldo,
      completarOnboarding,
      replaceState,
      resetPartial,
      resetFull,
    ]
  );

  return <AppContext.Provider value={value}>{children}</AppContext.Provider>;
}

export function useApp() {
  const ctx = useContext(AppContext);
  if (!ctx) throw new Error('useApp dentro de AppProvider');
  return ctx;
}
