import React, { createContext, useCallback, useContext, useEffect, useMemo, useState } from 'react';
import { useApp } from './AppContext';
import { reunirNotificacionesApp } from '../lib/notificacionesApp';
import {
  loadFirmasLectura,
  marcarAvisosActualesComoVistos,
  saveFirmasLectura,
} from '../lib/notificacionesLectura';

const Ctx = createContext(null);

/**
 * Misma “firma de lectura” en toda la app: un clic en la campana en una pantalla
 * marca visto para todas (tabs y otras instancias de la campanita).
 */
export function NotificacionLecturaProvider({ children }) {
  const { state } = useApp();
  const [firmasLeidas, setFirmasLeidas] = useState(null);

  useEffect(() => {
    let active = true;
    loadFirmasLectura().then((m) => {
      if (active) setFirmasLeidas(m || {});
    });
    return () => {
      active = false;
    };
  }, []);

  const marcarVistosAhora = useCallback(() => {
    if (!state) return;
    const cur = reunirNotificacionesApp(state, new Date()).items;
    setFirmasLeidas((prev) => {
      const base = prev == null ? {} : prev;
      const next = marcarAvisosActualesComoVistos(cur, base);
      saveFirmasLectura(next).catch(() => {});
      return next;
    });
  }, [state]);

  const value = useMemo(() => ({ firmasLeidas, marcarVistosAhora }), [firmasLeidas, marcarVistosAhora]);

  return <Ctx.Provider value={value}>{children}</Ctx.Provider>;
}

export function useNotificacionLectura() {
  const v = useContext(Ctx);
  if (!v) throw new Error('useNotificacionLectura dentro de NotificacionLecturaProvider');
  return v;
}
