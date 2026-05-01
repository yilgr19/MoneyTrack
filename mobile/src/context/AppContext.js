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
  clearOnboardingCompletado,
  loadOnboardingReabrirActivo,
  setOnboardingReabrirActivo,
  clearOnboardingReabrirActivo,
} from '../lib/storage';
import { exportarRespaldoCompartir, importarRespaldoElegirArchivo } from '../lib/backupMoneyTrack';
import { reemplazarPagosRecordatorioTarjetas, generarIdGasto } from '../lib/finance';
import { normalizarIntencionCompraPersistida, normalizarLineaListaSuper } from '../lib/asistenteComprasLogic';
import { notificacionesSistemaDisponibles } from '../lib/notificacionesLocalesEntorno';

export function normalizeState(raw) {
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
    nombreUsuario: String(raw.nombreUsuario || '').trim().slice(0, 80),
    saldosCuentas: saldos,
    bancosDetalle,
    limiteTarjetaCredito: parseFloat(raw.limiteTarjetaCredito) || 0,
    presupuestoMensual: parseFloat(raw.presupuestoMensual) || 0,
    presupuestoDesdeFecha: String(raw.presupuestoDesdeFecha || '')
      .trim()
      .slice(0, 10),
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
      ? raw.tarjetasCredito
          .filter((r) => r && typeof r === 'object')
          .map((t) => {
            const { deudaInicialDesdeCorte: _omitDeudaIni, ...rest } = t;
            return {
              ...rest,
              cuotasDeudaInicial: Math.max(1, parseInt(String(t.cuotasDeudaInicial ?? '1'), 10) || 1),
            };
          })
      : [],
    extractosTarjetasHistorial: Array.isArray(raw.extractosTarjetasHistorial)
      ? raw.extractosTarjetasHistorial.filter((r) => r && typeof r === 'object' && r.id)
      : [],
    bolsillos: Array.isArray(raw.bolsillos) ? raw.bolsillos.filter((r) => r && typeof r === 'object' && r.id) : [],
    recordatoriosPagoRegistrado: Array.isArray(raw.recordatoriosPagoRegistrado)
      ? raw.recordatoriosPagoRegistrado.filter((k) => typeof k === 'string' && k.trim())
      : [],
    intencionesCompra: Array.isArray(raw.intencionesCompra)
      ? raw.intencionesCompra.map(normalizarIntencionCompraPersistida).filter((x) => x && x.estado === 'pendiente')
      : [],
    asistenteUmbral48h: parseFloat(raw.asistenteUmbral48h) >= 0 ? parseFloat(raw.asistenteUmbral48h) : 50,
    listaSuperCategoriaPreferida: String(raw.listaSuperCategoriaPreferida || ''),
    listaSuperArticulosExtra: Array.isArray(raw.listaSuperArticulosExtra)
      ? raw.listaSuperArticulosExtra.filter((s) => typeof s === 'string' && s.trim()).map((s) => s.trim())
      : [],
    listaSuperCompraItems: Array.isArray(raw.listaSuperCompraItems)
      ? raw.listaSuperCompraItems.map(normalizarLineaListaSuper).filter(Boolean)
      : [],
    avisosGastosMovimiento: Array.isArray(raw.avisosGastosMovimiento)
      ? raw.avisosGastosMovimiento.filter(
          (x) =>
            x &&
            typeof x === 'object' &&
            String(x.id || '').trim() &&
            (x.tipo === 'editado' || x.tipo === 'eliminado')
        )
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

function withTimeout(promise, ms) {
  return Promise.race([
    promise,
    new Promise((_, reject) => setTimeout(() => reject(new Error('timeout')), ms)),
  ]);
}

export function AppProvider({ children }) {
  const [state, setState] = useState(null);
  const stateRef = useRef(null);
  const [ready, setReady] = useState(false);
  const [mostrarOnboarding, setMostrarOnboarding] = useState(false);
  /** Tras el primer tutorial: abrir pestaña Saldo para que el usuario digite saldos iniciales allí. */
  const [postOnboardingIrASaldo, setPostOnboardingIrASaldo] = useState(false);
  /** Tras «ver tutorial otra vez»: al terminar no forzar pestaña Saldo (solo la primera vez). */
  const irASaldoTrasOnboardingRef = useRef(true);

  useEffect(() => {
    stateRef.current = state;
  }, [state]);

  useEffect(() => {
    let cancelled = false;
    let finished = false;
    /** Red de seguridad si todo lo demás falla (p. ej. almacenamiento que no resuelve). */
    const INIT_MS = 12000;

    const finish = (normalized, onboardingHecho) => {
      if (cancelled || finished) return;
      finished = true;
      setState(normalized);
      setMostrarOnboarding(!onboardingHecho);
      setReady(true);
    };

    const timeoutId = setTimeout(() => {
      if (cancelled || finished) return;
      if (typeof __DEV__ !== 'undefined' && __DEV__) {
        console.warn('[MoneyTrack] Carga inicial lenta o bloqueada; continuando con datos vacíos.');
      }
      finish(normalizeState({}), false);
    }, INIT_MS);

    (async () => {
      try {
        let raw;
        let flagOnboarding;
        let reabrirTutorialActivo = false;
        try {
          [raw, flagOnboarding, reabrirTutorialActivo] = await withTimeout(
            Promise.all([loadAppState(), loadOnboardingCompletado(), loadOnboardingReabrirActivo()]),
            7000
          );
        } catch {
          raw = {};
          flagOnboarding = false;
          reabrirTutorialActivo = false;
        }
        const normalized = normalizeState(raw);
        let onboardingHecho = flagOnboarding;
        if (!onboardingHecho && tieneDatosPrevios(normalized) && !reabrirTutorialActivo) {
          try {
            await withTimeout(setOnboardingCompletado(), 4000);
          } catch {
            /* no bloquear arranque si setItem se cuelga */
          }
          onboardingHecho = true;
        }
        if (!cancelled) {
          if (reabrirTutorialActivo && !onboardingHecho) {
            irASaldoTrasOnboardingRef.current = false;
          }
          clearTimeout(timeoutId);
          finish(normalized, onboardingHecho);
        }
      } catch (e) {
        if (typeof __DEV__ !== 'undefined' && __DEV__) {
          console.error('[MoneyTrack] Error al cargar almacenamiento:', e);
        }
        if (!cancelled) {
          clearTimeout(timeoutId);
          finish(normalizeState({}), false);
        }
      }
    })();

    return () => {
      cancelled = true;
      clearTimeout(timeoutId);
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

  /** Gastos antiguos sin `id`: asignar uno y persistir en el siguiente guardado. */
  useEffect(() => {
    if (!ready || !state) return;
    const arr = state.gastos;
    if (!Array.isArray(arr) || !arr.some((g) => g && !String(g.id || '').trim())) return;
    replaceState((s) => ({
      ...s,
      gastos: (s.gastos || []).map((g) => {
        if (!g) return g;
        if (String(g.id || '').trim()) return g;
        return { ...g, id: generarIdGasto() };
      }),
    }));
  }, [ready, state, replaceState]);

  /**
   * Solo para efectos: cambia si cambia la lista de pagos programados (misma ref si solo moviste gastos/ingresos).
   * Evita re-sincronizar notificaciones en cada actualización de estado y posibles cierres/atascos en Expo Go.
   */
  const claveParaNotificacionesPagos = useMemo(
    () => `${String(state?.moneda ?? '')}\n${JSON.stringify(state?.pagosProgramados || [])}`,
    [state?.moneda, state?.pagosProgramados]
  );

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
  /** Ítems de lista de compras en urgente: al cambiar, reprogramar avisos locales. */
  const listaSuperUrgHash = useMemo(() => {
    const arr = (state?.listaSuperCompraItems || []).filter((x) => x && x.urgencia === 'urgente');
    return arr
      .map((x) => `${x.id}:${String(x.nombre || '').trim()}`)
      .sort()
      .join('|');
  }, [state?.listaSuperCompraItems]);
  /** Digest de gastos para reprogramar avisos de presupuesto al editar montos/fechas. */
  const gastosDigestPresupuesto = useMemo(() => {
    const arr = state?.gastos || [];
    return arr
      .filter(Boolean)
      .map((g) => `${String(g.id ?? '')}:${String(g.cantidad ?? '')}:${String(g.fecha || '').slice(0, 10)}`)
      .sort()
      .join('\t');
  }, [state?.gastos]);
  const metasDigestLocales = useMemo(() => {
    const arr = state?.metas || [];
    return arr
      .filter(Boolean)
      .map((m) => `${String(m.id ?? '')}:${String(parseFloat(m.objetivo) || 0)}`)
      .join('|');
  }, [state?.metas]);
  const contribMetasDigestLocales = useMemo(() => {
    try {
      return JSON.stringify(state?.contribucionesMetas || []);
    } catch {
      return '';
    }
  }, [state?.contribucionesMetas]);
  /** Pagos programados + TC + gastos (pago al día) + lista súper urgente + presupuesto + metas + nombre para reprogramar avisos. */
  const claveParaNotificacionesLocales = useMemo(
    () =>
      `${claveParaNotificacionesPagos}|tc:${tcFechasHash}|n:${nTc}|g:${lenG}|lsu:${listaSuperUrgHash}|pres:${String(state?.presupuestoMensual ?? '')}|pdf:${String(state?.presupuestoDesdeFecha ?? '')}|gdp:${gastosDigestPresupuesto}|md:${metasDigestLocales}|cm:${contribMetasDigestLocales}|nu:${String(state?.nombreUsuario ?? '')}`,
    [
      claveParaNotificacionesPagos,
      tcFechasHash,
      nTc,
      lenG,
      listaSuperUrgHash,
      state?.presupuestoMensual,
      state?.presupuestoDesdeFecha,
      gastosDigestPresupuesto,
      metasDigestLocales,
      contribMetasDigestLocales,
      state?.nombreUsuario,
    ]
  );
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
    import('../lib/notificacionesLocalesPagosProgramados').then((m) => {
      const log = (e, etiqueta) => {
        if (typeof __DEV__ !== 'undefined' && __DEV__) {
          console.warn(`[MoneyTrack] ${etiqueta}`, e?.message || e);
        }
      };
      m.sincronizarNotificacionesLocalesPagosProgramados(state).catch((e) => log(e, 'sync notif pagos'));
      m.sincronizarNotificacionesLocalesTarjetasCredito(state).catch((e) => log(e, 'sync notif TC'));
      m.sincronizarNotificacionesLocalesPresupuesto(state).catch((e) => log(e, 'sync notif presupuesto'));
      m.sincronizarNotificacionesLocalesMetas(state).catch((e) => log(e, 'sync notif metas'));
    });
    import('../lib/notificacionesLocalesListaSuper').then((m) => {
      m.sincronizarNotificacionesListaSuperUrgente(state).catch((e) => {
        if (typeof __DEV__ !== 'undefined' && __DEV__) {
          console.warn('[MoneyTrack] sync notif lista súper urgente', e?.message || e);
        }
      });
    });
  }, [ready, mostrarOnboarding, claveParaNotificacionesLocales]);

  useEffect(() => {
    if (!ready || mostrarOnboarding) return;
    if (!notificacionesSistemaDisponibles()) return;
    const sub = AppState.addEventListener('change', (next) => {
      if (next === 'active' && stateRef.current) {
        import('../lib/notificacionesLocalesPagosProgramados').then((m) => {
          const cur = stateRef.current;
          if (!cur) return;
          m.sincronizarNotificacionesLocalesPagosProgramados(cur).catch((e) => {
            if (typeof __DEV__ !== 'undefined' && __DEV__) console.warn('[MoneyTrack] sync notif pagos (active)', e);
          });
          m.sincronizarNotificacionesLocalesTarjetasCredito(cur).catch((e) => {
            if (typeof __DEV__ !== 'undefined' && __DEV__) console.warn('[MoneyTrack] sync notif TC (active)', e);
          });
          m.sincronizarNotificacionesLocalesPresupuesto(cur).catch((e) => {
            if (typeof __DEV__ !== 'undefined' && __DEV__) {
              console.warn('[MoneyTrack] sync notif presupuesto (active)', e);
            }
          });
          m.sincronizarNotificacionesLocalesMetas(cur).catch((e) => {
            if (typeof __DEV__ !== 'undefined' && __DEV__) {
              console.warn('[MoneyTrack] sync notif metas (active)', e);
            }
          });
        });
        import('../lib/notificacionesLocalesListaSuper').then((m) => {
          const cur = stateRef.current;
          if (!cur) return;
          m.sincronizarNotificacionesListaSuperUrgente(cur).catch((e) => {
            if (typeof __DEV__ !== 'undefined' && __DEV__) {
              console.warn('[MoneyTrack] sync notif lista súper (active)', e);
            }
          });
        });
      }
    });
    return () => sub.remove();
  }, [ready, mostrarOnboarding]);

  const clearPostOnboardingIrASaldo = useCallback(() => {
    setPostOnboardingIrASaldo(false);
  }, []);

  const reabrirTutorialOnboarding = useCallback(async () => {
    irASaldoTrasOnboardingRef.current = false;
    try {
      await setOnboardingReabrirActivo();
    } catch {
      /* continuar */
    }
    await clearOnboardingCompletado();
    setPostOnboardingIrASaldo(false);
    setMostrarOnboarding(true);
  }, []);

  const completarOnboarding = useCallback(async (nombreUsuario) => {
    const n = String(nombreUsuario || '').trim().slice(0, 80);
    const irASaldo = irASaldoTrasOnboardingRef.current;
    irASaldoTrasOnboardingRef.current = true;
    replaceState((s) => ({ ...s, nombreUsuario: n }));
    try {
      await clearOnboardingReabrirActivo();
    } catch {
      /* continuar */
    }
    await setOnboardingCompletado();
    setPostOnboardingIrASaldo(irASaldo);
    setMostrarOnboarding(false);
  }, [replaceState]);

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

  const exportarDatosRespaldo = useCallback(async () => {
    if (!state) {
      return { ok: false, mensaje: 'Aún no hay datos cargados.' };
    }
    const onboardingHecho = await loadOnboardingCompletado();
    return exportarRespaldoCompartir(state, onboardingHecho);
  }, [state]);

  const importarDatosRespaldo = useCallback(async () => {
    const r = await importarRespaldoElegirArchivo();
    if (r.cancelado) {
      return { ok: false, cancelado: true, mensaje: '' };
    }
    if (!r.ok) {
      return { ok: false, mensaje: r.error || 'No se pudo importar.' };
    }
    const normalized = normalizeState(r.data);
    await persistAppState(normalized);
    if (r.onboardingCompletado) {
      await setOnboardingCompletado();
      try {
        await clearOnboardingReabrirActivo();
      } catch {
        /* ignorar */
      }
      setMostrarOnboarding(false);
    } else {
      await clearOnboardingCompletado();
      setMostrarOnboarding(true);
    }
    setPostOnboardingIrASaldo(false);
    setState(normalized);
    const fecha = r.exportedAt ? `\nArchivo exportado: ${r.exportedAt}` : '';
    return {
      ok: true,
      mensaje: `Datos restaurados. Reinicia la app si algo no se ve al instante.${fecha}`,
    };
  }, []);

  const value = useMemo(
    () => ({
      state,
      ready,
      mostrarOnboarding,
      postOnboardingIrASaldo,
      clearPostOnboardingIrASaldo,
      completarOnboarding,
      reabrirTutorialOnboarding,
      replaceState,
      resetPartial,
      resetFull,
      exportarDatosRespaldo,
      importarDatosRespaldo,
    }),
    [
      state,
      ready,
      mostrarOnboarding,
      postOnboardingIrASaldo,
      clearPostOnboardingIrASaldo,
      completarOnboarding,
      reabrirTutorialOnboarding,
      replaceState,
      resetPartial,
      resetFull,
      exportarDatosRespaldo,
      importarDatosRespaldo,
    ]
  );

  return <AppContext.Provider value={value}>{children}</AppContext.Provider>;
}

export function useApp() {
  const ctx = useContext(AppContext);
  if (!ctx) throw new Error('useApp dentro de AppProvider');
  return ctx;
}
