import React, { createContext, useCallback, useContext, useEffect, useMemo, useState } from 'react';
import { loadAppState, persistAppState, emptySaldosCuentas, clearStoragePartial, clearStorageFull } from '../lib/storage';

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

const AppContext = createContext(null);

export function AppProvider({ children }) {
  const [state, setState] = useState(null);
  const [ready, setReady] = useState(false);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      const raw = await loadAppState();
      if (!cancelled) {
        setState(normalizeState(raw));
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
      replaceState,
      resetPartial,
      resetFull,
    }),
    [state, ready, replaceState, resetPartial, resetFull]
  );

  return <AppContext.Provider value={value}>{children}</AppContext.Provider>;
}

export function useApp() {
  const ctx = useContext(AppContext);
  if (!ctx) throw new Error('useApp dentro de AppProvider');
  return ctx;
}
