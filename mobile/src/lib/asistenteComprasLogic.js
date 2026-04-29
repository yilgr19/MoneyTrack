import {
  normalizarCategoria,
  obtenerMesAño,
  montoGastoAfectaSaldoEnMes,
  formatearNumero,
  obtenerCuentasOrigenGastoElegible,
} from './finance';

export function generarIdIntencionCompra() {
  return `int_${Date.now()}_${Math.random().toString(36).slice(2, 10)}`;
}

export function normalizarIntencionCompraPersistida(raw) {
  if (!raw || typeof raw !== 'object') return null;
  const estado = raw.estado === 'completada' || raw.estado === 'cancelada' ? raw.estado : 'pendiente';
  return {
    id: String(raw.id || generarIdIntencionCompra()),
    nombre: String(raw.nombre || '').trim(),
    precioEstimado: Math.max(0, parseFloat(raw.precioEstimado) || 0),
    nombreCategoria: String(raw.nombreCategoria || '').trim(),
    vecesPorSemana: Math.max(0.01, parseFloat(raw.vecesPorSemana) || 3),
    minutosPorSesion: Math.max(1, parseFloat(raw.minutosPorSesion) || 60),
    añosUso: Math.max(0.25, parseFloat(raw.añosUso) || 3),
    creadoEn: typeof raw.creadoEn === 'number' ? raw.creadoEn : Date.now(),
    aplicabaCooldown:
      raw.aplicabaCooldown === true || (raw.cooldownHasta != null && parseFloat(raw.cooldownHasta) > 0),
    cooldownHasta: raw.cooldownHasta != null ? parseFloat(raw.cooldownHasta) || null : null,
    estado,
  };
}

/** Sesiones totales estimadas en la vida útil del producto */
export function totalSesionesEstimadas(vecesPorSemana, añosUso) {
  const v = Math.max(0.01, parseFloat(vecesPorSemana) || 1);
  const a = Math.max(0.25, parseFloat(añosUso) || 1);
  return v * 52 * a;
}

/** Costo por uso / sesión */
export function costoPorSesion(precio, vecesPorSemana, añosUso) {
  const total = totalSesionesEstimadas(vecesPorSemana, añosUso);
  const p = Math.max(0, parseFloat(precio) || 0);
  return total > 0 ? p / total : p;
}

const PRECIO_REF_CINE_DEFAULT = 10;

export function mensajeCostoPorUso({
  nombreProducto,
  precio,
  costoSesion,
  precioReferenciaSesion = PRECIO_REF_CINE_DEFAULT,
  etiquetaReferencia = 'una salida al cine',
}) {
  const nombre = String(nombreProducto || 'Este artículo').trim();
  const cs = Math.max(0, costoSesion || 0);
  const ref = Math.max(0.01, precioReferenciaSesion);
  const ratio = cs / ref;
  const csFmt = formatearNumero(cs);
  const refFmt = formatearNumero(ref);
  const precioFmt = formatearNumero(parseFloat(precio) || 0);

  if (ratio <= 1.15) {
    return `${nombre} sale ${csFmt} por uso estimado frente a ~${refFmt} en ${etiquetaReferencia}. Si lo usarás tanto como piensas, tiene sentido frente al ocio fuera de casa.`;
  }
  if (ratio <= 2.5) {
    return `${nombre} (${precioFmt} total): cada uso sería ~${csFmt}; comparado con ~${refFmt} (${etiquetaReferencia}), piensa si lo usarás lo suficiente para amortizarlo.`;
  }
  return `${nombre}: cada uso estimado ~${csFmt}, bastante más caro que ~${refFmt} (${etiquetaReferencia}). Si solo lo usarás un par de veces al año, el costo por uso es alto: ¿merece la pena?`;
}

/** Gastado en categoría este mes (misma regla que Gastos). */
export function gastadoEnCategoriaMes(state, nombreCategoria, mes, año) {
  const gastos = state?.gastos || [];
  const nom = String(nombreCategoria || '').trim();
  let s = 0;
  for (let i = 0; i < gastos.length; i++) {
    const g = gastos[i];
    if (!g || String(g.categoria || '').trim() !== nom) continue;
    s += montoGastoAfectaSaldoEnMes(g, state, mes, año);
  }
  return s;
}

export function datosTermometroCategoria(state, nombreCategoria, mes, año, precioPropuesto) {
  const categorias = (state?.categorias || []).map(normalizarCategoria);
  const nom = String(nombreCategoria || '').trim();
  const cat = categorias.find((c) => String(c.nombre || '').trim() === nom);
  const limiteCat = cat && cat.limite != null ? parseFloat(cat.limite) || 0 : 0;
  const gastado = gastadoEnCategoriaMes(state, nom, mes, año);
  const propuesto = Math.max(0, parseFloat(precioPropuesto) || 0);

  const presupuestoGlobal = parseFloat(state?.presupuestoMensual) || 0;

  let barraPct = 0;
  let sombraHastaPct = 0;
  let hayLimite = limiteCat > 0;

  if (hayLimite) {
    barraPct = Math.min(100, (gastado / limiteCat) * 100);
    sombraHastaPct = Math.min(100, ((gastado + propuesto) / limiteCat) * 100);
  } else if (presupuestoGlobal > 0) {
    hayLimite = true;
    barraPct = Math.min(100, (gastado / presupuestoGlobal) * 100);
    sombraHastaPct = Math.min(100, ((gastado + propuesto) / presupuestoGlobal) * 100);
  }

  const restante = hayLimite
    ? limiteCat > 0
      ? Math.max(0, limiteCat - gastado - propuesto)
      : Math.max(0, presupuestoGlobal - gastado - propuesto)
    : null;

  let alertaTexto = null;
  if (limiteCat > 0 && gastado + propuesto > limiteCat) {
    alertaTexto = `Si compras esto ahora, pasarías el tope mensual en «${nom}».`;
  } else if (limiteCat > 0 && restante !== null && gastado + propuesto <= limiteCat) {
    alertaTexto = `Si compras esto ahora, te quedarían solo ~${formatearNumero(restante)} para el resto del mes en «${nom}».`;
  } else if (limiteCat <= 0 && presupuestoGlobal > 0 && restante !== null) {
    alertaTexto = `Vista orientativa con tu presupuesto global del mes (no hay límite por categoría definido). Quedarían ~${formatearNumero(restante)} en el tope mensual tras esta compra.`;
  }

  return {
    hayLimite,
    usaPresupuestoGlobal: limiteCat <= 0 && presupuestoGlobal > 0,
    limiteMostrado: limiteCat > 0 ? limiteCat : presupuestoGlobal,
    etiquetaLimite:
      limiteCat > 0
        ? `Límite «${nom}»`
        : presupuestoGlobal > 0
          ? 'Presupuesto mensual (referencia)'
          : 'Sin tope definido',
    gastado,
    propuesto,
    barraPct,
    sombraHastaPct,
    restanteTrasCompra: restante,
    alertaTexto,
  };
}

/** Promedio por ticket de compra en categoría en un mes calendario dado */
function promedioTicketCategoriaEnMes(state, nombreCategoria, mes, año) {
  const gastos = state?.gastos || [];
  const nom = String(nombreCategoria || '').trim();
  const montos = [];
  for (let i = 0; i < gastos.length; i++) {
    const g = gastos[i];
    if (!g || String(g.categoria || '').trim() !== nom) continue;
    const { mes: m, año: y } = obtenerMesAño(g.fecha);
    if (m !== mes || y !== año) continue;
    const monto = montoGastoAfectaSaldoEnMes(g, state, mes, año);
    if (monto > 0) montos.push(monto);
  }
  if (!montos.length) return null;
  const suma = montos.reduce((a, b) => a + b, 0);
  return { promedio: suma / montos.length, visitas: montos.length, total: suma };
}

/**
 * Estimación de la bolsa según registros históricos del mes indicado (por defecto mes anterior al actual).
 */
export function estimarListaSuperDesdeHistorial(state, nombreCategoria, numItemsMarcados, refMesOffset = 1) {
  const ref = new Date();
  ref.setMonth(ref.getMonth() - refMesOffset);
  const mes = ref.getMonth();
  const año = ref.getFullYear();
  const stats = promedioTicketCategoriaEnMes(state, nombreCategoria, mes, año);
  const n = Math.max(0, parseInt(numItemsMarcados, 10) || 0);
  const nombresMes = [
    'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
    'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre',
  ];
  const etiquetaMes = `${nombresMes[mes]} ${año}`;

  if (!stats || stats.visitas < 1) {
    return {
      estimado: n * 5,
      confianza: 'baja',
      mensaje: `No hay bastantes registros en «${nombreCategoria}» para ${etiquetaMes}. Estimación orientativa: ~${n * 5} unidades (supuesto ~5 c/u).`,
      etiquetaMes,
    };
  }

  const factor = Math.min(1.4, 0.35 + n * 0.12);
  const estimado = stats.promedio * factor;
  return {
    estimado,
    confianza: 'media',
    mensaje: `Según ${stats.visitas} registro(s) en «${nombreCategoria}» (${etiquetaMes}), un ticket típico rondaba ${formatearNumero(stats.promedio)}. Con ${n} artículos marcados, una bolsa similar podría acercarse a ~${formatearNumero(estimado)} (orientativo).`,
    etiquetaMes,
    promedioTicket: stats.promedio,
  };
}

export function elegirCategoriaSuperPorDefecto(state) {
  const pref = String(state?.listaSuperCategoriaPreferida || '').trim();
  if (pref) return pref;
  const cats = (state?.categorias || []).map(normalizarCategoria);
  const hit = cats.find((c) => /super|merc|abarro|despensa|comida/i.test(String(c.nombre || '')));
  if (hit) return String(hit.nombre).trim();
  return cats[0]?.nombre ? String(cats[0].nombre).trim() : '';
}

/** Cuenta por defecto para registrar gasto (primera con saldo suficiente entre opciones válidas). */
export function primeraCuentaParaGasto(state, monto) {
  const m = Math.max(0, parseFloat(monto) || 0);
  const cuentas = obtenerCuentasOrigenGastoElegible(state || {}, m, m, {});
  if (cuentas.length > 0) return cuentas[0].value;
  return 'efectivo';
}

export function generarIdListaSuperLinea() {
  return `ls_${Date.now()}_${Math.random().toString(36).slice(2, 10)}`;
}

const RANK_URGENCIA = { urgente: 0, normal: 1, puede_esperar: 2 };

export function normalizarLineaListaSuper(raw) {
  if (!raw || typeof raw !== 'object') return null;
  const nom = String(raw.nombre || '').trim();
  if (!nom) return null;
  let u = raw.urgencia;
  if (u !== 'urgente' && u !== 'puede_esperar' && u !== 'normal') u = 'normal';
  return {
    id: String(raw.id || generarIdListaSuperLinea()),
    nombre: nom,
    urgencia: u,
  };
}

/** Orden: urgente primero, luego normal, puede esperar; dentro, por nombre */
export function ordenarLineasListaSuper(lines) {
  const arr = [...(lines || [])];
  arr.sort((a, b) => {
    const ra = RANK_URGENCIA[a.urgencia] ?? 9;
    const rb = RANK_URGENCIA[b.urgencia] ?? 9;
    if (ra !== rb) return ra - rb;
    return String(a.nombre).localeCompare(String(b.nombre), 'es');
  });
  return arr;
}

export const URGENCIA_LISTA_SUPER = [
  { id: 'urgente', label: 'Urgente' },
  { id: 'normal', label: 'Normal' },
  { id: 'puede_esperar', label: 'Puede esperar' },
];

/** Regla 48 h: si aplica cooldown, hasta `cooldownHasta` no se puede registrar compra. */
export function puedeRegistrarCompraPorRegla48h(intencion, ahoraMs) {
  if (!intencion || intencion.estado !== 'pendiente') return false;
  if (!intencion.aplicabaCooldown) return true;
  const hasta = intencion.cooldownHasta;
  if (hasta == null) return true;
  return ahoraMs >= hasta;
}

function pad2(n) {
  return String(n).padStart(2, '0');
}

/** Cuenta atrás HH:MM:SS para UI de intenciones. */
export function formatCountdownMs(ms) {
  if (ms <= 0) return '00:00:00';
  const sTotal = Math.floor(ms / 1000);
  const h = Math.floor(sTotal / 3600);
  const m = Math.floor((sTotal % 3600) / 60);
  const s = sTotal % 60;
  return `${pad2(h)}:${pad2(m)}:${pad2(s)}`;
}

export { PRECIO_REF_CINE_DEFAULT };
