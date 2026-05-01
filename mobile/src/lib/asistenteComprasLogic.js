import {
  normalizarCategoria,
  obtenerMesAño,
  montoGastoAfectaSaldoEnMes,
  montoGastoCuentaParaPresupuestoEnMes,
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

/** Mínimo de tickets en la ventana para comparar con tu historial real (evita ruido). */
const MIN_TICKETS_HISTORIAL_CONFIABLE = 3;
const MESES_VENTANA_HISTORIAL = 3;

/**
 * Tickets de gasto en la categoría en los últimos `numMeses` meses calendario (incluye el mes en curso).
 */
export function benchmarkTicketsCategoriaUltimosMeses(state, nombreCategoria, numMeses = MESES_VENTANA_HISTORIAL) {
  const nom = String(nombreCategoria || '').trim();
  if (!nom) {
    return { confiable: false, visitas: 0, promedio: 0, mediana: 0 };
  }
  const montos = [];
  const d = new Date();
  for (let k = 0; k < numMeses; k++) {
    const ref = new Date(d.getFullYear(), d.getMonth() - k, 15);
    const mes = ref.getMonth();
    const año = ref.getFullYear();
    const gastos = state?.gastos || [];
    for (let i = 0; i < gastos.length; i++) {
      const g = gastos[i];
      if (!g || String(g.categoria || '').trim() !== nom) continue;
      const ma = obtenerMesAño(g.fecha);
      if (ma.mes !== mes || ma.año !== año) continue;
      const monto = montoGastoAfectaSaldoEnMes(g, state, mes, año);
      if (monto > 0) montos.push(monto);
    }
  }
  if (montos.length === 0) {
    return { confiable: false, visitas: 0, promedio: 0, mediana: 0 };
  }
  const suma = montos.reduce((a, b) => a + b, 0);
  const promedio = suma / montos.length;
  const sorted = [...montos].sort((a, b) => a - b);
  const mediana = sorted[Math.floor(sorted.length / 2)];
  return {
    confiable: montos.length >= MIN_TICKETS_HISTORIAL_CONFIABLE,
    visitas: montos.length,
    promedio,
    mediana,
  };
}

function parrafoHistorialConfiable(bench, cat, precio) {
  if (!bench || !bench.confiable) return '';
  const nom = String(cat || '').trim();
  const p = formatearNumero(bench.promedio);
  const m = formatearNumero(bench.mediana);
  const n = bench.visitas;
  if (precio <= bench.promedio * 0.88) {
    return `Según tus ${n} gastos registrados recientes en «${nom}», sueles mover cerca de ${p} por compra (mediana ~${m}). Lo que evalúas va por debajo de ese ritmo: los datos no apuntan a un exceso claro.`;
  }
  if (precio <= bench.promedio * 1.18) {
    return `Con ${n} compras recientes en «${nom}», tu ticket típico ronda ${p} (mediana ~${m}). El precio que miras va en la línea de lo que ya registraste; el patrón es reconocible.`;
  }
  return `Tus últimos registros en «${nom}» (${n} compras) sitúan un ticket habitual cerca de ${p} (mediana ~${m}). Estás por encima de ese patrón: merece la pena asegurarse del motivo antes de gastar.`;
}

function fnv1aHash32(str) {
  let h = 2166136261;
  const s = String(str || '');
  for (let i = 0; i < s.length; i++) h = Math.imul(h ^ s.charCodeAt(i), 16777619);
  return h >>> 0;
}

/**
 * 20 avisos rotativos (tono cercano). Banda según sesiones totales estimadas (datos del formulario), sin referencias externas ficticias.
 */
const PLANTILLAS_ANALISIS_INTENCION = [
  (c) =>
    `${c.nombre} encaja bien si el uso es real: repartes ${c.precioFmt} en unas ${c.sesionesFmt} sesiones estimadas y cada uso queda en ~${c.csFmt}.`,
  (c) =>
    `Con tantas sesiones previstas, el costo por uso (~${c.csFmt}) se mantiene razonable frente al total (${c.precioFmt}). Solo evita duplicar algo que ya tienes.`,
  (c) =>
    `La amortización acompaña: muchos usos diluyen el precio. Si cumples ese ritmo, el gasto se defiende con números.`,
  (c) =>
    `Por cómo lo contaste, cada vez que lo uses “cuesta” ~${c.csFmt}; para el total ${c.precioFmt} eso suele ser señal de compra pensada, no impulsiva.`,
  (c) =>
    `El “precio por uso” sale contenido. Eso invita a decir sí solo si de verdad lo vas a integrar al día a día.`,
  (c) =>
    `Los supuestos de uso que pusiste hacen que el número por sesión sea amable. Falta que encaje con tu mes y con tus prioridades, no con la calculadora.`,
  (c) =>
    `En términos de uso frecuente, el reparto del ${c.precioFmt} tiene sentido. La duda ya no es la división, sino si lo sacarás tanto del armario.`,
  (c) =>
    `Zona intermedia: cada uso ~${c.csFmt} con ~${c.sesionesFmt} sesiones estimadas. Si el uso real es menor del que anotaste, el costo por vez sube; vale la pena ser honesto.`,
  (c) =>
    `No es alarmante, pero ya pide constancia. Si al final lo usarás pocas veces, el ~${c.csFmt} por uso se va a sentir caro.`,
  (c) =>
    `Aquí el hábito marca la diferencia: con las sesiones que estimaste, el gasto se entiende; sin ese hábito, conviene frenar.`,
  (c) =>
    `El total (${c.precioFmt}) aún se puede defender si las sesiones son reales. Revisa veces por semana y años de uso: un pequeño cambio mueve mucho el costo por uso.`,
  (c) =>
    `Ni luz verde chillona ni roja fuerte: el ~${c.csFmt} por uso pide que confirmes si lo necesitas o si hay alternativa más barata.`,
  (c) =>
    `Un “tal vez” sincero: si puedes esperar oferta, comparar otra marca o segunda mano, este tramo intermedio suele premiar la paciencia.`,
  (c) =>
    `Vuelve a mirar el formulario: minutos por sesión y veces por semana definen esas ${c.sesionesFmt} sesiones. Si dudas de los datos, duda también del veredicto.`,
  (c) =>
    `Si la vida real te aleja de los usos que pusiste, el costo por uso real será mayor que ~${c.csFmt}. Ajusta cifras o asume el riesgo.`,
  (c) =>
    `Pocas sesiones estimadas: cada uso ~${c.csFmt} pesa mucho frente al ${c.precioFmt}. Sin uso intensivo, el dinero suele encontrar mejor sitio.`,
  (c) =>
    `La amortización va justa: salvo que sea imprescindible y lo uses muy seguido, es fácil arrepentirse del precio por cada vez.`,
  (c) =>
    `Los números piden pausa: repartir ${c.precioFmt} en tan pocas sesiones encarece cada uso. Valora esperar o buscar sustituto.`,
  (c) =>
    `Por uso, sale caro. Tiene sentido solo si el valor práctico o emocional es muy alto para ti.`,
  (c) =>
    `No es un gasto que se diluya solo con un uso esporádico: sin constancia, ~${c.csFmt} por vez duele; prioriza necesidad real sobre impulso.`,
];

/**
 * Texto principal del análisis de intención + si hubo comparación con historial real.
 */
export function construirAnalisisMensajeIntencion(state, editando, costoSesion) {
  const nombre = String(editando?.nombre || 'Esto').trim() || 'Esto';
  const precio = Math.max(0, parseFloat(editando?.precioEstimado) || 0);
  const cat = String(editando?.nombreCategoria || '').trim();
  const cs = Math.max(0, costoSesion || 0);
  const sesTot = Math.max(0, totalSesionesEstimadas(editando.vecesPorSemana, editando.añosUso));
  /** Solo datos del formulario: más sesiones → mejor amortización del precio. */
  const bucket = sesTot >= 45 ? 'alta' : sesTot >= 12 ? 'media' : 'baja';
  const bench = benchmarkTicketsCategoriaUltimosMeses(state, cat, MESES_VENTANA_HISTORIAL);
  let hist = parrafoHistorialConfiable(bench, cat, precio);
  if (
    !hist &&
    cat &&
    bench.visitas > 0 &&
    bench.visitas < MIN_TICKETS_HISTORIAL_CONFIABLE
  ) {
    hist = `Por ahora hay ${bench.visitas} gasto(s) reciente(s) en «${cat}»; con al menos ${MIN_TICKETS_HISTORIAL_CONFIABLE} registros en esa categoría el contraste con tu ritmo real será más fiable.`;
  }
  const ctx = {
    nombre,
    precioFmt: formatearNumero(precio),
    csFmt: formatearNumero(cs),
    sesionesFmt: formatearNumero(Math.round(sesTot)),
  };
  const pools = {
    alta: [0, 1, 2, 3, 4, 5, 6],
    media: [7, 8, 9, 10, 11, 12, 13, 14],
    baja: [15, 16, 17, 18, 19],
  };
  const pool = pools[bucket];
  const seed = `${editando?.id || nombre}|${bucket}|${bench.confiable ? 1 : 0}`;
  const pick = pool[fnv1aHash32(seed) % pool.length];
  const cuerpo = PLANTILLAS_ANALISIS_INTENCION[pick](ctx);
  const msgValor = hist ? `${hist}\n\n${cuerpo}` : cuerpo;
  return { msgValor, historialConfiable: bench.confiable, ticketsHistorial: bench.visitas };
}

/** Gastado en categoría este mes (misma regla que Gastos). */
export function gastadoEnCategoriaMes(state, nombreCategoria, mes, año) {
  const gastos = state?.gastos || [];
  const nom = String(nombreCategoria || '').trim();
  let s = 0;
  for (let i = 0; i < gastos.length; i++) {
    const g = gastos[i];
    if (!g || String(g.categoria || '').trim() !== nom) continue;
    s += montoGastoCuentaParaPresupuestoEnMes(g, state, mes, año);
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
