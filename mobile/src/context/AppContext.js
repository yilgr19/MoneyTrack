import React, { createContext, useCallback, useContext, useEffect, useMemo, useState } from 'react';
import {
  loadAppState,
  persistAppState,
  emptySaldosCuentas,
  clearStoragePartial,
  clearStorageFull,
  loadOnboardingCompletado,
  setOnboardingCompletado,
} from '../lib/storage';

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
  };
}

/** Instalaciones anteriores al onboarding: si ya había actividad, no forzar el recorrido. */
function tieneDatosPrevios(n) {
  if (String(n.moneda || '').trim()) return true;
  const saldos = n.saldosCuentas || {};
  if (Object.values(saldos).some((v) => (parseFloat(v) || 0) > 0)) return true;
  if (Array.isArray(n.gastos) && n.gastos.length > 0) return true;
  if (Array.isArray(n.ingresos) && n.ingresos.length > 0) return true;
  if (Array.isArray(n.categorias) && n.categorias.length > 0) return true;
  if (Array.isArray(n.metas) && n.metas.length > 0) return true;
  if (String(n.saldoInicialNota || '').trim()) return true;
  return false;
}

const AppContext = createContext(null);

export function AppProvider({ children }) {
  const [state, setState] = useState(null);
  const [ready, setReady] = useState(false);
  const [mostrarOnboarding, setMostrarOnboarding] = useState(false);
  /** Tras el primer tutorial: abrir pestaña Saldo para que el usuario digite saldos iniciales allí. */
  const [postOnboardingIrASaldo, setPostOnboardingIrASaldo] = useState(false);

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

  const clearPostOnboardingIrASaldo = useCallback(() => {
    setPostOnboardingIrASaldo(false);
  }, []);

  const completarOnboarding = useCallback(async () => {
    await setOnboardingCompletado();
    setPostOnboardingIrASaldo(true);
    setMostrarOnboarding(false);
  }, []);

  const replaceState = useCallback((updater) => {
    setState((prev) => {
      if (!prev) return prev;
      const next = typeof updater === 'function' ? updater(prev) : updater;
      persistAppState(next).catch(() => {});
      return next;
    });
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
