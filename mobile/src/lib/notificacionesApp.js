import {
  calcularSaldosPorCuenta,
  construirExtractoBancarioTarjeta,
  diasCalendarioHasta,
  diasHastaProximoDiaCalendario,
  formatearNumero,
  limiteTotalTarjetasCredito,
  montoGastoAfectaSaldoEnMes,
  montoPagoSugeridoDesdeExtracto,
  normalizarCategoria,
  obtenerSaldosIniciales,
  parseFechaHoraLocal,
  pagoDebeMostrarseParaPagar,
  proximaOcurrenciaMensual,
  resumenAlertasTarjetasCredito,
  totalSaldoLiquido,
} from './finance';

const DIAS_PAGO_CERCA = 7;
/** Bajo esto, aviso de “poco en efectivo/cuentas” (sin contar cupo de tarjeta). */
const LIQ_BAJO_UMBRAL = 100_000;

/** Misma notificación, mismo mensaje: evita “parpadeos” al marcar leída; añade variedad fija por id. */
function rotarFrase(semilla, frases) {
  if (!frases || frases.length === 0) return '';
  const s = String(semilla);
  let h = 0;
  for (let i = 0; i < s.length; i += 1) h = (h * 31 + s.charCodeAt(i)) % 10007;
  return frases[h % frases.length];
}

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
      const id = `pp-${p.id}-pagar`;
      out.push({
        id,
        tipo: 'pago',
        severidad: 'danger',
        puntuacionOrden: 1_000_000,
        titulo: rotarFrase(id, [
          `${concepto}: pendiente de anotar en el mes (Gastos).`,
          `Falta registrar ${concepto} en el mes. Entra a Gastos.`,
        ]),
        detalle: rotarFrase(id + 'b', [
          'Bloque de pagos programados o formulario. Tú eliges el momento.',
          'Un registro y el resumen acompaña a lo real.',
        ]),
      });
      continue;
    }

    if (d == null) continue;
    if (d < 0 && p.frecuencia === 'unico') {
      const id = `pp-${p.id}-unico-pasado`;
      out.push({
        id,
        tipo: 'pago',
        severidad: 'warning',
        puntuacionOrden: 860_000,
        titulo: `Pago único ${concepto}: la fecha ya pasó.`,
        detalle: rotarFrase(id, [
          'Más → Pagos programados: fecha o borrado. Tú eliges.',
          'Ajusta o quita en Pagos programados.',
        ]),
      });
      continue;
    }
    if (d === 0) {
      const id = `pp-${p.id}-hoy`;
      out.push({
        id,
        tipo: 'pago',
        severidad: 'warning',
        puntuacionOrden: 960_000,
        titulo: rotarFrase(id, [
          `Hoy: ${concepto} (Gastos).`,
          `Recuerda hoy: ${concepto}.`,
        ]),
        detalle: rotarFrase(id + 'b', [
          'Si pagaste, anótalo en Gastos.',
          'Gastos: marcar o repasar en un minuto.',
        ]),
      });
    } else if (d > 0 && d <= DIAS_PAGO_CERCA) {
      const id = `pp-${p.id}-cerca`;
      const dTxt = d === 1 ? 'mañana' : `en ${d} días`;
      // Más cerca = más arriba
      out.push({
        id,
        tipo: 'pago',
        severidad: d <= 2 ? 'warning' : 'info',
        puntuacionOrden: 950_000 - d * 12_000,
        titulo: rotarFrase(id, [
          `${concepto} ${dTxt} · revisa en Gastos.`,
          `${dTxt} cae ${concepto}. No lo pierdas de vista.`,
        ]),
        detalle: rotarFrase(String(d) + id, [
          'Gastos o Más → Pagos programados.',
          'Dale un vistazo a lo guardado.',
        ]),
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
      .filter((g) => g.categoria === cat.nombre)
      .reduce((s, g) => s + montoGastoAfectaSaldoEnMes(g, state, m, y), 0);
    if (gastado > lim) {
      const id = `cat-${cat.nombre}`;
      const gTxt = formatearNumero(gastado, 0);
      const lTxt = formatearNumero(lim, 0);
      out.push({
        id,
        tipo: 'categoria',
        severidad: 'warning',
        puntuacionOrden: 500_000,
        titulo: rotarFrase(id, [
          `${cat.nombre} ${cat.icono}: gastos ${gTxt} · tope ${lTxt}. Un poco por encima.`,
          `${cat.nombre} ${cat.icono} · ${gTxt} de ${lTxt} tope. Pasaste poco.`,
        ]),
        detalle: rotarFrase(id + 'x', [
          'Toca subir tope o bajar gasto. Categorías. Tú eliges.',
          'Ajusta en Categorías o el mes apretando.',
        ]),
      });
    }
  }
  return out;
}

function tuvoMovimientosOConfigLiquido(state) {
  if ((state.gastos || []).length + (state.ingresos || []).length > 0) return true;
  const ini = obtenerSaldosIniciales(state);
  for (const id of ['efectivo', 'banco', 'nequi', 'daviplata', 'billeteras']) {
    if (Math.abs(parseFloat(ini[id]) || 0) > 0.0001) return true;
  }
  return false;
}

/**
 * Aviso cuando en efectivo + banco + billeteras (sin cupo de tarjeta) el total va en cero o bajo.
 */
function notificacionesSaldo(state) {
  if (!tuvoMovimientosOConfigLiquido(state)) return [];
  const liquido = totalSaldoLiquido(state);
  const q = formatearNumero(liquido, 0);
  if (liquido <= 0) {
    return [
      {
        id: 'saldo-liquido-critico',
        tipo: 'saldo',
        severidad: 'danger',
        puntuacionOrden: 920_000,
        titulo: rotarFrase('sal-cri', [
          `Efectivo + bancos + apps: ${q} (sin cupo de tarjeta en esta suma).`,
          `Plata al día: ${q}. Revisar cajas reales en Saldo.`,
        ]),
        detalle: rotarFrase('sal-cri2', [
          'Saldo: ingresos, gasto o ajuste de saldos. Tú diriges.',
          'Que cuadre con el bolsillo, no solo con el cupo.',
        ]),
      },
    ];
  }
  if (liquido > 0 && liquido < LIQ_BAJO_UMBRAL) {
    return [
      {
        id: 'saldo-liquido-bajo',
        tipo: 'saldo',
        severidad: 'warning',
        puntuacionOrden: 450_000,
        titulo: rotarFrase('sal-baj', [
          `Poco en efectivo/cuentas: unos ${q} (cupo de tarjeta aparte).`,
          `Caja al día: ~${q} · cuida el margen.`,
        ]),
        detalle: rotarFrase('sal-baj2', [
          'Saldo o Ingreso si hace falta. Tú eliges.',
          'Antes de ajustar, mira en Saldo.',
        ]),
      },
    ];
  }
  return [];
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
    const limTxt = formatearNumero(limite, 0);
    out.push({
      id: 'tc-sin-cupo',
      tipo: 'tc',
      severidad: 'danger',
      puntuacionOrden: 910_000,
      titulo: rotarFrase('tcsin', [
        `Cupo libre 0 (app) · tope anotado ${limTxt} · afinar en Saldo.`,
        `TC: 0 libre según registro. Revisa límite y usado (Saldo).`,
      ]),
      detalle: rotarFrase('tcsin2', [
        'Saldo → Tarjetas, que coincida con el banco.',
        'Ajusta números en Saldo si no cuadra.',
      ]),
    });
  } else if (limite > 0 && r.global.porcentaje >= 90 && (r.tarjetas || []).length === 0) {
    const pg = r.global.porcentaje.toFixed(0);
    out.push({
      id: 'tc-cupo-alto',
      tipo: 'tc',
      severidad: 'warning',
      puntuacionOrden: 420_000,
      titulo: rotarFrase('tcglo', [
        `Cupo global ~${pg}% usado. Ojo a Saldo y números.`,
        `Casi lleno el ${pg}% del cupo (global). Revisa.`,
      ]),
      detalle: rotarFrase('tcglo2', [
        'Ajusta límite o ritmo de gasto en Saldo. Tú eliges.',
        'Saldo: bajar carga o el tope que diste.',
      ]),
    });
  }

  const mon = (state.moneda && String(state.moneda).trim()) || '';
  const tcs = state.tarjetasCredito || [];
  for (const t of r.tarjetas || []) {
    if (t.cupoTotal <= 0) continue;
    const libre = t.cupoTotal - t.cupoUtilizado;
    const tRaw = t.id ? tcs.find((x) => x && x.id === t.id) : null;
    if (t.corteHoy || t.diasCorte === 0) {
      const idEx = `tc-extracto-${t.id}`;
      let pagoSug = '';
      let cierreTxt = '';
      let intTxt = '';
      if (tRaw) {
        const ex = construirExtractoBancarioTarjeta(tRaw, state, ref);
        const s = montoPagoSugeridoDesdeExtracto(ex);
        pagoSug = `${formatearNumero(s, 0)}${mon ? ` ${mon}` : ''}`;
        cierreTxt = `${formatearNumero(ex.capitalCierreLineas, 0)}${mon ? ` ${mon}` : ''}`;
        intTxt = `${formatearNumero(ex.intereses, 0)}${mon ? ` ${mon}` : ''}`;
      }
      out.push({
        id: idEx,
        tipo: 'tc',
        severidad: 'warning',
        tarjetaId: t.id,
        puntuacionOrden: 1_100_000,
        titulo: rotarFrase(idEx, [
          pagoSug
            ? `Corte ${t.nombreEntidad}: pago sugerido ~${pagoSug} (cuotas + deuda, int. aprox.).`
            : `Día de corte · abre el extracto (${t.nombreEntidad}).`,
          pagoSug
            ? `Corte: ~${pagoSug} · ${t.nombreEntidad} · revisa Más → Pagos programados.`
            : `Corte hoy: movimientos y extracto · ${t.nombreEntidad}.`,
        ]),
        detalle: pagoSug
          ? rotarFrase(idEx + 'd', [
              `Capital cierre: ${cierreTxt || '—'}. Int. est. periodo: ${intTxt || '0'}. Hay recordatorio con monto en Pagos programados (TC). Saldo: extracto de la tarjeta.`,
              'Montos aprox. con tasa E.A. y tramos a cuota en Gastos. Confirma en banco. Saldo: ver detalle.',
            ])
          : rotarFrase(idEx + 'd', [
              'Toca para ver cupo, detalle y proyección a 3/6 cuotas.',
              'Extracto con deuda, cuotas y ahorro estimado. Toca.',
            ]),
      });
    }
    if (t.alertaPagoUrgente) {
      const id = `tc-pago-${t.id}`;
      const dP = t.diasPago;
      let pagoL = '';
      if (tRaw) {
        const exL = construirExtractoBancarioTarjeta(tRaw, state, ref);
        pagoL = `${formatearNumero(montoPagoSugeridoDesdeExtracto(exL), 0)}${mon ? ` ${mon}` : ''}`;
      }
      out.push({
        id,
        tipo: 'tc',
        severidad: 'danger',
        puntuacionOrden: 1_000_000 - dP * 18_000,
        titulo: rotarFrase(id, [
          pagoL
            ? `${t.nombreEntidad} · pago ~${pagoL} en ${dP} d. (vence ${t.etiquetaProxPago || 'Saldo'}).`
            : `${t.nombreEntidad} · pago en ${dP} día${dP === 1 ? '' : 's'} (${t.etiquetaProxPago || 'Saldo'}).`,
          pagoL
            ? `Vence: ~${pagoL} · ${t.nombreEntidad} · ${dP} d.`
            : `Cerca el pago ${t.nombreEntidad} · ${dP} d. · ${t.etiquetaProxPago || 'ver en Saldo'}.`,
        ]),
        detalle: rotarFrase(id + 'p', [
          pagoL
            ? 'Incluye tramos a cuota e int. aprox. Pagos programados (TC) y Gastos. Corrobora con el banco.'
            : 'Banco o Saldo, sin pasarte de la fecha.',
          pagoL ? 'Que cuadre con el extracto real; evita intereses o mora.' : 'Anótalo donde lo veas: evita intereses o mora.',
        ]),
      });
    } else if (t.alertaCorte && t.diasCorte > 0 && t.diasCorte <= 2 && !t.corteHoy) {
      const id = `tc-corte-${t.id}`;
      let corteSug = '';
      if (tRaw) {
        const ex2 = construirExtractoBancarioTarjeta(tRaw, state, ref);
        corteSug = `${formatearNumero(montoPagoSugeridoDesdeExtracto(ex2), 0)}${mon ? ` ${mon}` : ''}`;
      }
      out.push({
        id,
        tipo: 'tc',
        severidad: 'warning',
        puntuacionOrden: 800_000 - t.diasCorte * 15_000,
        titulo: rotarFrase(id, [
          corteSug
            ? `Corte en ${t.diasCorte} d. · ${t.nombreEntidad} · sugerido ~${corteSug}.`
            : `Corte ${t.nombreEntidad} en unos ${t.diasCorte} d. (${t.etiquetaProxCorte || 'Saldo'}).`,
          `${t.nombreEntidad} · corte cerca, ${t.diasCorte} d.`,
        ]),
        detalle: rotarFrase(id + 'c', [
          'Revisa compras y tramos a cuota en Gastos. Recordatorio y monto anotados en Pagos programados. Saldo: extracto.',
          'Cierre con cifra clara: mira en Saldo y confirma en banco.',
        ]),
      });
    } else if (t.alertaUtil) {
      const idU = `tc-uso-${t.id}`;
      const pU = t.utilPct.toFixed(0);
      out.push({
        id: idU,
        tipo: 'tc',
        severidad: 'info',
        puntuacionOrden: 320_000,
        titulo: rotarFrase(idU, [
          `${t.nombreEntidad} · ~${pU}% del cupo usado.`,
          `${pU}% de cupo en ${t.nombreEntidad} · ojo a Saldo.`,
        ]),
        detalle: rotarFrase(idU + 'u', [
          'Ajusta ritmo o límite en Saldo. Tú eliges.',
          'Si ajusta, baja carga o el tope.',
        ]),
      });
    } else if (libre > 0 && libre < t.cupoTotal * 0.05) {
      const idL = `tc-libre-${t.id}`;
      const lTxt = formatearNumero(libre, 0);
      const totTxt = formatearNumero(t.cupoTotal, 0);
      out.push({
        id: idL,
        tipo: 'tc',
        severidad: 'info',
        puntuacionOrden: 300_000,
        titulo: rotarFrase(idL, [
          `${t.nombreEntidad} · poco libre: ${lTxt} / ${totTxt}.`,
          `Margen bajo en ${t.nombreEntidad} · ~${lTxt} libre.`,
        ]),
        detalle: rotarFrase(idL + 'f', [
          'Alinea con el banco en Saldo si no cuadra.',
          'Saldo: que el dato acompañe a la real.',
        ]),
      });
    }
  }
  return out;
}

const ORDEN_SEVER = { danger: 0, warning: 1, info: 2 };

/** Más puntuación = arriba: primero lo más “próximo/urgente” (pagos, tarjeta) y luego el resto. */
function puntuacionOrdenDefecto(sem) {
  const s = ORDEN_SEVER[sem] ?? 2;
  if (s === 0) return 600_000;
  if (s === 1) return 300_000;
  return 100_000;
}

/**
 * Lista unificada de notificaciones para el centro de campana.
 * @returns {{ items: Array<{id,titulo,detalle,tipo,severidad}>, total: number }}
 */
export function reunirNotificacionesApp(state, ref = new Date()) {
  const a = notificacionesPagos(state, ref);
  const b = notificacionesCategorias(state, ref);
  const s = notificacionesSaldo(state);
  const c = notificacionesTarjetas(state, ref);
  const items = [...a, ...b, ...s, ...c]
    .map((it) => ({
      ...it,
      puntuacionOrden: it.puntuacionOrden != null ? it.puntuacionOrden : puntuacionOrdenDefecto(it.severidad),
    }))
    .sort((x, y) => {
      const diff = (y.puntuacionOrden ?? 0) - (x.puntuacionOrden ?? 0);
      if (diff !== 0) return diff;
      return ORDEN_SEVER[x.severidad] - ORDEN_SEVER[y.severidad];
    });
  const seen = new Set();
  const deduped = items.filter((it) => {
    if (seen.has(it.id)) return false;
    seen.add(it.id);
    return true;
  });
  // No exponer puntuación al UI; firma y lista usan el resto
  return {
    items: deduped.map(({ puntuacionOrden: _p, ...rest }) => rest),
    total: deduped.length,
  };
}
