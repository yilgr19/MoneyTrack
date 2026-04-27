import {
  calcularSaldosPorCuenta,
  diasCalendarioHasta,
  diasHastaProximoDiaCalendario,
  formatearNumero,
  limiteTotalTarjetasCredito,
  montoGastoAfectaSaldo,
  normalizarCategoria,
  obtenerMesAño,
  parseFechaHoraLocal,
  pagoDebeMostrarseParaPagar,
  proximaOcurrenciaMensual,
  resumenAlertasTarjetasCredito,
} from './finance';

const DIAS_PAGO_CERCA = 7;

function diasHastaPagoMensual(pago, ref = new Date()) {
  const dia = Math.min(28, parseInt(pago.diaPago, 10) || 1);
  return diasHastaProximoDiaCalendario(dia, ref);
}

function proximoQuincenal(ref) {
  const r = new Date(ref.getFullYear(), ref.getMonth(), ref.getDate(), 12, 0, 0);
  const cands = [];
  for (const d of [1, 15]) {
    const pat = new Date(2020, 0, d, 12, 0, 0);
    const p = proximaOcurrenciaMensual(pat, new Date(r.getTime() - 86400000));
    if (p) cands.push(p);
  }
  if (cands.length === 0) return null;
  return cands.reduce((a, b) => (a < b ? a : b));
}

function diasHastaPagoQuincenal(pago, ref) {
  const p = proximoQuincenal(ref);
  if (!p) return null;
  return diasCalendarioHasta(p, ref);
}

function diasHastaPagoSemanal(pago, ref) {
  if (!pago.fechaInicio) return null;
  const start = parseFechaHoraLocal(pago.fechaInicio);
  if (!start) return null;
  const wantDow = start.getDay();
  const base = new Date(ref.getFullYear(), ref.getMonth(), ref.getDate());
  for (let add = 0; add < 7; add++) {
    const d = new Date(base);
    d.setDate(base.getDate() + add);
    if (d.getDay() === wantDow) {
      return diasCalendarioHasta(d, ref);
    }
  }
  return 7;
}

/**
 * Días hasta el próximo vencimiento (0 = hoy, negativo = retraso respecto a la fecha de calendario esperada).
 */
function diasHastaPagoProgramado(pago, ref = new Date()) {
  if (!pago || pago.activo === false) return null;

  if (pago.frecuencia === 'unico') {
    const f = parseFechaHoraLocal(pago.fechaInicio);
    if (!f) return null;
    return diasCalendarioHasta(
      new Date(f.getFullYear(), f.getMonth(), f.getDate(), 12, 0, 0),
      new Date(ref.getFullYear(), ref.getMonth(), ref.getDate(), 12, 0, 0)
    );
  }
  if (pago.frecuencia === 'mensual') {
    return diasHastaPagoMensual(pago, ref);
  }
  if (pago.frecuencia === 'quincenal') {
    return diasHastaPagoQuincenal(pago, ref);
  }
  if (pago.frecuencia === 'semanal') {
    return diasHastaPagoSemanal(pago, ref);
  }
  return null;
}

/**
 * Avisos de pagos programados: “listo para pagar” (misma regla que Gastos) o vence en ≤7 días.
 */
function notificacionesPagos(state, ref) {
  const out = [];
  const pagos = state.pagosProgramados || [];
  for (const p of pagos) {
    if (p.activo === false) continue;
    const concepto = String(p.concepto || 'Pago').trim() || 'Pago programado';
    const d = diasHastaPagoProgramado(p, ref);
    const listoGastos = pagoDebeMostrarseParaPagar(p, ref);

    if (listoGastos) {
      out.push({
        id: `pp-${p.id}-pagar`,
        tipo: 'pago',
        severidad: 'danger',
        titulo: `Por pagar: ${concepto}`,
        detalle: 'Toca en Gastos (bloque de pagos programados) para rellenar el formulario o regístralo allí.',
      });
      continue;
    }

    if (d == null) continue;
    if (d < 0 && p.frecuencia === 'unico') {
      out.push({
        id: `pp-${p.id}-unico-pasado`,
        tipo: 'pago',
        severidad: 'warning',
        titulo: `Pendiente: ${concepto}`,
        detalle: 'Fecha de pago única ya pasó. Confirma o elimina en Pagos programados.',
      });
      continue;
    }
    if (d === 0) {
      out.push({
        id: `pp-${p.id}-hoy`,
        tipo: 'pago',
        severidad: 'warning',
        titulo: concepto,
        detalle: 'Toca hoy. Revisa en Gastos o Pagos programados.',
      });
    } else if (d > 0 && d <= DIAS_PAGO_CERCA) {
      out.push({
        id: `pp-${p.id}-cerca`,
        tipo: 'pago',
        severidad: d <= 2 ? 'warning' : 'info',
        titulo: `Próximo: ${concepto}`,
        detalle: `Vence en ${d} día${d === 1 ? '' : 's'}.`,
      });
    }
  }
  return out;
}

function notificacionesCategorias(state, ref) {
  const out = [];
  const ahora = ref instanceof Date ? ref : new Date();
  const m = ahora.getMonth();
  const y = ahora.getFullYear();
  const gastos = state.gastos || [];
  const cats = (state.categorias || []).map(normalizarCategoria);
  for (const cat of cats) {
    if (!cat.limite) continue;
    const lim = parseFloat(cat.limite);
    if (Number.isNaN(lim) || lim <= 0) continue;
    const gastado = gastos
      .filter((g) => {
        if (g.categoria !== cat.nombre) return false;
        const { mes, año } = obtenerMesAño(g.fecha);
        return mes === m && año === y;
      })
      .reduce((s, g) => s + montoGastoAfectaSaldo(g), 0);
    if (gastado > lim) {
      out.push({
        id: `cat-${cat.nombre}`,
        tipo: 'categoria',
        severidad: 'warning',
        titulo: `Límite superado: ${cat.icono} ${cat.nombre}`,
        detalle: `Gastos del mes: ${formatearNumero(gastado)} (límite ${formatearNumero(lim)}). Revisa en Categorías.`,
      });
    }
  }
  return out;
}

/**
 * Saldo “disponible” en cupo vía app + alertas de resumen (pago/corte, uso alto).
 */
function notificacionesTarjetas(state, ref) {
  const out = [];
  const r = resumenAlertasTarjetasCredito(state, ref);
  const limite = limiteTotalTarjetasCredito(state);
  const saldos = calcularSaldosPorCuenta(state);
  const disp = saldos.tarjetaCredito ?? 0;

  if (limite > 0 && disp <= 0.0001) {
    out.push({
      id: 'tc-sin-cupo',
      tipo: 'tc',
      severidad: 'danger',
      titulo: 'Tarjeta de crédito: sin cupo disponible',
      detalle:
        'El saldo de cupo visto en la app es 0. Revisa en Saldo → Tarjetas o registra el cupo y lo utilizado.',
    });
  } else if (limite > 0 && r.global.porcentaje >= 90 && (r.tarjetas || []).length === 0) {
    out.push({
      id: 'tc-cupo-alto',
      tipo: 'tc',
      severidad: 'warning',
      titulo: 'Uso muy alto del cupo (global)',
      detalle: `${r.global.porcentaje.toFixed(0)}% del límite según lo registrado en la app. Controla en Saldo.`,
    });
  }

  for (const t of r.tarjetas || []) {
    if (t.cupoTotal <= 0) continue;
    const libre = t.cupoTotal - t.cupoUtilizado;
    if (t.alertaPagoUrgente) {
      out.push({
        id: `tc-pago-${t.id}`,
        tipo: 'tc',
        severidad: 'danger',
        titulo: `Pago próximo: ${t.nombreEntidad}`,
        detalle: `Fecha límite en ${t.diasPago} d. · ${t.etiquetaProxPago || '—'}`,
      });
    } else if (t.alertaCorte) {
      out.push({
        id: `tc-corte-${t.id}`,
        tipo: 'tc',
        severidad: 'warning',
        titulo: `Corte próximo: ${t.nombreEntidad}`,
        detalle: `Corte en ${t.diasCorte} d. · ${t.etiquetaProxCorte || '—'}`,
      });
    } else if (t.alertaUtil) {
      out.push({
        id: `tc-uso-${t.id}`,
        tipo: 'tc',
        severidad: 'info',
        titulo: `Uso alto de cupo: ${t.nombreEntidad}`,
        detalle: `${t.utilPct.toFixed(0)}% utilizado. Ver Saldo → Tarjeta.`,
      });
    } else if (libre > 0 && libre < t.cupoTotal * 0.05) {
      out.push({
        id: `tc-libre-${t.id}`,
        tipo: 'tc',
        severidad: 'info',
        titulo: `Poco cupo libre: ${t.nombreEntidad}`,
        detalle: `Queda disponible aprox. ${libre.toFixed(0)} según registro de la app.`,
      });
    }
  }
  return out;
}

const ORDEN_SEVER = { danger: 0, warning: 1, info: 2 };

/**
 * Lista unificada de notificaciones para el centro de campana.
 * @returns {{ items: Array<{id,titulo,detalle,tipo,severidad}>, total: number }}
 */
export function reunirNotificacionesApp(state, ref = new Date()) {
  const a = notificacionesPagos(state, ref);
  const b = notificacionesCategorias(state, ref);
  const c = notificacionesTarjetas(state, ref);
  const items = [...a, ...b, ...c].sort((x, y) => ORDEN_SEVER[x.severidad] - ORDEN_SEVER[y.severidad]);
  const seen = new Set();
  const deduped = items.filter((it) => {
    if (seen.has(it.id)) return false;
    seen.add(it.id);
    return true;
  });
  return { items: deduped, total: deduped.length };
}
