import {
  calcularSaldosPorCuenta,
  construirExtractoBancarioTarjeta,
  diasCalendarioHasta,
  diasHastaProximoDiaCalendario,
  formatearNumero,
  limiteTotalTarjetasCredito,
  montoGastoAfectaSaldoEnMes,
  montoGastoCuentaParaPresupuestoEnMes,
  montoPagoSugeridoDesdeExtracto,
  normalizarCategoria,
  normalizarMeta,
  obtenerSaldosIniciales,
  parseFechaHoraLocal,
  pagoDebeMostrarseParaPagar,
  proximaOcurrenciaMensual,
  resumenAlertasTarjetasCredito,
  totalSaldoLiquido,
} from './finance';
import { ordenarLineasListaSuper } from './asistenteComprasLogic';
import { varianteGastoEditadoCampana, varianteGastoEliminadoCampana } from './notificacionesVariantesAmigables';
import { tituloNotifConNombre } from './notificacionesPersonalizacion';

/** Cuántos avisos de editar/quitar gasto guardamos en estado (campana). */
export const MAX_AVISOS_GASTOS_CAMPANA = 25;

/**
 * @param {'editado'|'eliminado'} tipo
 * @param {{ nombre?: string, montoLine?: string }} datos
 */
export function nuevaEntradaAvisoGastoMovimiento(tipo, datos) {
  if (tipo !== 'editado' && tipo !== 'eliminado') return null;
  const nombre = String(datos?.nombre || '').trim().slice(0, 72) || 'Gasto';
  const montoLine = String(datos?.montoLine || '').trim().slice(0, 48);
  return {
    id: `gmov-${tipo}-${Date.now()}-${Math.random().toString(36).slice(2, 9)}`,
    ts: Date.now(),
    tipo,
    nombre,
    montoLine,
  };
}

/** Añade un aviso a `avisosGastosMovimiento` y recorta la cola. */
export function withAvisoGastoMovimiento(prevState, tipo, datos) {
  const ent = nuevaEntradaAvisoGastoMovimiento(tipo, datos);
  if (!ent || !prevState) return prevState;
  const prev = Array.isArray(prevState.avisosGastosMovimiento) ? prevState.avisosGastosMovimiento : [];
  return {
    ...prevState,
    avisosGastosMovimiento: [...prev, ent].slice(-MAX_AVISOS_GASTOS_CAMPANA),
  };
}

/**
 * Tras ver el centro de avisos: quita del estado los ítems `gasto_movimiento` que estaban en la lista
 * (así no vuelven aunque la firma de lectura cambie).
 */
export function stateSinAvisosGastoMovimientoEnLista(prevState, itemsCampana) {
  if (!prevState || !Array.isArray(itemsCampana)) return prevState;
  const quitar = new Set(
    itemsCampana.filter((it) => it && it.tipo === 'gasto_movimiento').map((it) => String(it.id))
  );
  if (quitar.size === 0) return prevState;
  const prev = Array.isArray(prevState.avisosGastosMovimiento) ? prevState.avisosGastosMovimiento : [];
  const next = prev.filter((e) => e && !quitar.has(String(e.id)));
  if (next.length === prev.length) return prevState;
  return { ...prevState, avisosGastosMovimiento: next };
}

function notificacionesGastosMovimiento(state, ref = new Date()) {
  const arr = state.avisosGastosMovimiento;
  if (!Array.isArray(arr) || arr.length === 0) return [];
  const now = ref.getTime();
  const sorted = [...arr].sort((a, b) => (b.ts || 0) - (a.ts || 0));
  const slice = sorted.slice(0, MAX_AVISOS_GASTOS_CAMPANA);
  return slice.map((e) => {
    const ctx = {
      nombre: e.nombre || 'Gasto',
      montoLine: e.montoLine || '',
    };
    const v = e.tipo === 'eliminado' ? varianteGastoEliminadoCampana(ctx) : varianteGastoEditadoCampana(ctx);
    const ageMin = Math.max(0, (now - (e.ts || 0)) / 60000);
    const puntuacionOrden = 713_000 - Math.min(5000, Math.floor(ageMin * 35));
    return {
      id: String(e.id),
      tipo: 'gasto_movimiento',
      severidad: e.tipo === 'eliminado' ? 'warning' : 'info',
      puntuacionOrden,
      titulo: v.title,
      detalle: v.body,
    };
  });
}

/** Aviso en campana en los últimos días: id por día = recordatorio “diario” al abrir. */
const DIAS_RECORDATORIO_CAMPANA = 3;
/** Bajo esto, aviso de “poco en efectivo/cuentas” (sin contar cupo de tarjeta). */
const LIQ_BAJO_UMBRAL = 100_000;
/** Aviso “cupo al filo”: queda poco libre respecto al tope (≤32%). Incluye 5%, 10%, 20%…; sin piso mínimo para no omitir casos como 50.000 libres con tope alto. */
const TC_CUPO_LIBRE_AVISO_MAX = 0.32;

/**
 * Elige 1 de N textos de forma estable (pseudo-azar): misma semilla → misma frase hasta que cambie la semilla.
 * Así hay variedad entre avisos y días sin romper “leído/no leído” ni parpadear al re-renderizar.
 */
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
 * Exportado para sincronizar notificaciones locales del sistema.
 */
export function diasHastaPagoProgramado(pago, ref = new Date()) {
  if (!pago || pago.activo === false) return null;

  /** Solo la fecha elegida: avisos/campana no se repiten mes a mes (ni quincena/semana). */
  if (pago.recordarCadaMes === false) {
    const f = parseFechaHoraLocal(pago.fechaInicio);
    if (!f) return null;
    return diasCalendarioHasta(
      new Date(f.getFullYear(), f.getMonth(), f.getDate(), 12, 0, 0),
      new Date(ref.getFullYear(), ref.getMonth(), ref.getDate(), 12, 0, 0)
    );
  }

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
 * Avisos de pagos programados: “listo para pagar” (Gastos), hoy, o recordatorio diario en campana
 * los últimos 3 días antes del vencimiento (id distinto por día).
 */
function notificacionesPagos(state, ref) {
  const ymd = `${ref.getFullYear()}-${String(ref.getMonth() + 1).padStart(2, '0')}-${String(ref.getDate()).padStart(2, '0')}`;
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
        titulo: rotarFrase(`${id}-t-${ymd}`, [
          `${concepto} te espera en Gastos este mes 📒✨`,
          `Falta anotar ${concepto} · un minuto y quedas al día ⏱️💛`,
          `¿Ya pagaste ${concepto}? Regístralo en Gastos 🙌`,
          `${concepto} sin registrar · tu resumen lo agradecerá 📊😊`,
          `Pequeño recordatorio: ${concepto} pendiente en Gastos 💪`,
          `Vamos: anota ${concepto} en Gastos y respiras tranquilo 😌`,
          `${concepto} · márcalo en Gastos cuando puedas 📝💚`,
          `Tu app quiere cuadrar: falta ${concepto} en Gastos ✨`,
          `${concepto} · un registro rápido en Gastos y listo 🚀`,
          `Casi perfecto: solo falta ${concepto} en Gastos ⭐`,
        ]),
        detalle: rotarFrase(`${id}-d-${ymd}`, [
          'Puedes hacerlo desde Gastos o Pagos programados · tú mandas 😊',
          'Un solo registro y todo sigue clarito 📌✨',
          'Sin estrés: Gastos y en un momento queda listo 💛',
          'Así tu mes queda ordenado y sin sorpresas 🗓️😌',
          'Tú eliges cuándo; esto es solo un recordatorio con cariño 💬',
          'Pásate por Gastos cuando te venga bien · aquí estamos 🙏',
          'Registrar es la forma más fácil de cuidar tu plata 💰✨',
          'Pequeño paso, gran claridad en tu resumen 📊💚',
          'Anótalo y sigue con tu día en paz ☀️',
          'Nadie juzga aquí: solo un empujoncito amable 🤗',
        ]),
      });
      continue;
    }

    if (d == null) continue;
    if (d < 0 && (p.frecuencia === 'unico' || p.recordarCadaMes === false)) {
      const id = `pp-${p.id}-unico-pasado`;
      out.push({
        id,
        tipo: 'pago',
        severidad: 'warning',
        puntuacionOrden: 860_000,
        titulo: rotarFrase(`${id}-t-${ymd}`, [
          `La fecha de ${concepto} ya pasó · revisa con calma 📅🤍`,
          `${concepto} (único) quedó atrás en el calendario 🌿`,
          `Ojo con ${concepto}: el plazo único ya fue 🔔`,
          `${concepto} · si ya no aplica, actualízalo sin drama ✨`,
          `Quizá toque ajustar ${concepto} en Pagos programados 📝`,
          `La fecha de ${concepto} ya no es futura · mira si sigue vigente 🙌`,
          `${concepto} · un vistazo rápido y lo dejas fino 👀💛`,
          `Nada grave: solo alinea ${concepto} con tu realidad hoy 😊`,
          `${concepto} pasó de fecha · tú decides el siguiente paso 💪`,
          `Tu calendario cambió; ${concepto} puede necesitar un toque 🗓️`,
        ]),
        detalle: rotarFrase(`${id}-d-${ymd}`, [
          'Más → Pagos programados · edita o borra en un momento 📱',
          'Si ya no toca, quítalo y sigue ligero 🌤️',
          'Cambiar la fecha es rápido y te quita ruido de la cabeza 🧠💚',
          'La app es tuya: ajústala sin miedo 🤗',
          'Un cambio pequeño y todo vuelve a tener sentido ✨',
          'Revisa el concepto y deja solo lo que te sirve 💛',
          'Puedes dejarlo perfecto en segundos ⏱️😊',
          'Nada complicado · entra, corrige y listo 👍',
          'Tu “yo del futuro” te lo agradece cuando está ordenado 🙏',
          'Recordatorio amable: esto es solo para ayudarte 💬',
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
        titulo: rotarFrase(`${id}-t-${ymd}`, [
          `Hoy es el día de ${concepto} · ¡tú puedes! 💪✨`,
          `Última vuelta para ${concepto} hoy 🎯💛`,
          `${concepto} vence hoy · date tu mejor versión 😊`,
          `Respira y organiza ${concepto} antes de que termine el día 🌙`,
          `Hoy toca ${concepto} · un pasito y lo bajas de la lista ✅`,
          `El plazo de ${concepto} es hoy · aquí vamos contigo 🤝`,
          `${concepto} · hoy es un buen día para cumplir 🌟`,
          `No estás solo: ${concepto} te recuerda con cariño 🔔💚`,
          `Hoy cerramos ${concepto} con buena energía ⚡😌`,
          `${concepto} espera por ti hoy · tú sabes cómo 💛`,
        ]),
        detalle: rotarFrase(`${id}-d-${ymd}`, [
          'Paga con calma y anótalo en Gastos · así todo cuadra 📒✨',
          'Si ya salió el dinero, regístralo y celebra el orden 🎉',
          'Un registro hoy = resumen claro mañana 📊😊',
          'Tu esfuerzo cuenta: deja el movimiento en Gastos 💪',
          'Así tu app refleja lo que realmente pasó 💚',
          'Es rápido: monto, cuenta y listo · tú puedes ⚡',
          'Cuidar el detalle te da paz con tu plata 🙌',
          'Hoy es buen día para estar al día contigo mismo 🤍',
          'Pequeño hábito, gran diferencia · anótalo 📝',
          'Lo importante es avanzar: un paso basta 🌿',
        ]),
      });
    } else if (d > 0 && d <= DIAS_RECORDATORIO_CAMPANA) {
      const idCerca = `pp-${p.id}-cerca-${ymd}`;
      const plazoTxt = d === 1 ? '1 día de plazo restante' : `${d} días de plazo restante`;
      const faltanTxt = d === 1 ? 'Falta 1 día' : `Faltan ${d} días`;
      out.push({
        id: idCerca,
        tipo: 'pago',
        severidad: d <= 2 ? 'warning' : 'info',
        puntuacionOrden: 950_000 - d * 12_000,
        titulo: rotarFrase(`${idCerca}-t`, [
          `${faltanTxt} para ${concepto} · vamos con calma 🐢💚`,
          `«${concepto}» · en ${d} día${d === 1 ? '' : 's'} te toca estar listo 📅✨`,
          `Respira: ${concepto} está a ${d} día${d === 1 ? '' : 's'} 🌤️`,
          `Aún hay tiempo para ${concepto} · úsalo a tu favor ⏳💛`,
          `${concepto} se acerca · piénsalo con tranquilidad ☕😊`,
          `Aviso suave: ${concepto} en ${d} día${d === 1 ? '' : 's'} 🔔`,
          `Organiza ${concepto} con calma; el plazo es pronto 🗓️🤍`,
          `${faltanTxt} · ${concepto} te espera sin drama 💬`,
          `Vas bien: solo falta afinar ${concepto} ✨`,
          `${concepto} · cuenta regresiva amable: ${d} día${d === 1 ? '' : 's'} ⏱️💚`,
        ]),
        detalle: rotarFrase(`${idCerca}-d`, [
          `${plazoTxt} · luego anótalo en Gastos y listo 📒😌`,
          'Paga tranquilo y deja el registro; tu resumen lo agradecerá ✨',
          'Si ya pagaste, mismo monto en Gastos y a celebrar el orden 🎉',
          'Un minuto en Gastos y todo queda alineado 💚',
          'Así no se te pasa nada por alto 👀💛',
          'Tú marcas el ritmo; esto solo te acompaña 🤗',
          'Registrar es cuidarte · vale la pena 🙏',
          'Paso a paso: primero pago, luego registro tranquilo 🌿',
          'Buena organización, cero presión fea 💪',
          'Recordatorio con abrazo virtual 🤗✨',
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
      .reduce((s, g) => s + montoGastoCuentaParaPresupuestoEnMes(g, state, m, y), 0);
    if (gastado > lim) {
      const id = `cat-${cat.nombre}`;
      const gTxt = formatearNumero(gastado, 0);
      const lTxt = formatearNumero(lim, 0);
      /** 10+10 frases; semilla con fecha: misma variedad todo el día, distinta otro día. */
      const ymd = `${y}-${String(m + 1).padStart(2, '0')}-${String(ahora.getDate()).padStart(2, '0')}`;
      out.push({
        id,
        tipo: 'categoria',
        severidad: 'warning',
        puntuacionOrden: 500_000,
        titulo: rotarFrase(`${id}-t-${ymd}`, [
          `${cat.nombre} ${cat.icono} · llevas ${gTxt} y el tope era ${lTxt} 🌿`,
          `Ojo amable: ${cat.nombre} ${cat.icono} pasó el límite (${gTxt} / ${lTxt}) 💛`,
          `En ${cat.nombre} ${cat.icono} te pasaste un poquito · ${gTxt} vs ${lTxt} 📊`,
          `${cat.nombre} ${cat.icono}: más gasto (${gTxt}) que lo que marcaste (${lTxt}) ✨`,
          `Tu guía en ${cat.nombre} ${cat.icono} era ${lTxt} · vas en ${gTxt} 🎯`,
          `Momento de revisar ${cat.nombre} ${cat.icono} · ${gTxt} sobre ${lTxt} 😊`,
          `${cat.nombre} ${cat.icono} pide un ajuste · gasto ${gTxt}, tope ${lTxt} 💪`,
          `Nada catastrófico: ${cat.nombre} ${cat.icono} superó el tope (${gTxt}) 🤝`,
          `${cat.nombre} ${cat.icono} · el límite ${lTxt} quedó corto con ${gTxt} ☕`,
          `Pequeña señal: ${cat.nombre} ${cat.icono} ${gTxt} > ${lTxt} · tú mandas 🔔`,
        ]),
        detalle: rotarFrase(`${id}-d-${ymd}`, [
          'Puedes subir el tope o bajar gasto · Más → Categorías 💚',
          'Ajustar el límite también es cuidarte · sin culpas 🤗',
          'Un cambio en Categorías y retomas el control ✨',
          'Si fue un mes especial, sube el tope con tranquilidad 📌',
          'Si prefieres ajustar hábitos, el tope te ayuda a mirar con claridad 👀',
          'Tú decides: más margen o más foco en gasto 💛',
          'La app solo refleja lo que pasó; mañana es otra oportunidad 🌅',
          'Pequeño desvío, gran aprendizaje · sigues yendo bien 🌟',
          'En Categorías lo dejas como te sienta mejor 😌',
          'Organizar es un acto de cariño hacia ti · un tap y listo 💪',
        ]),
      });
    }
  }
  return out;
}

/**
 * Aviso cuando el gasto del mes supera el presupuesto mensual.
 * Títulos: solo gastado y disponible (negativo = exceso); no se muestra la cifra del límite/tope.
 * Diez títulos y diez detalles en rotación diaria por semilla.
 */
function notificacionesPresupuestoMensual(state, ref) {
  const ahora = ref instanceof Date ? ref : new Date();
  const mes = ahora.getMonth();
  const anio = ahora.getFullYear();
  const tope = parseFloat(state?.presupuestoMensual) || 0;
  if (tope <= 0) return [];
  const gastos = state.gastos || [];
  const gastosMes = gastos.reduce(
    (s, g) => s + montoGastoCuentaParaPresupuestoEnMes(g, state, mes, anio),
    0
  );
  if (gastosMes <= tope) return [];
  const ymd = `${anio}-${String(mes + 1).padStart(2, '0')}-${String(ahora.getDate()).padStart(2, '0')}`;
  const gTxt = formatearNumero(gastosMes, 0);
  const exceso = gastosMes - tope;
  const exT = formatearNumero(exceso, 0);
  const mon = (state.moneda && String(state.moneda).trim()) || '';
  const mS = mon ? ` ${mon}` : '';

  return [
    {
      id: 'presupuesto-supera-tope',
      tipo: 'presupuesto',
      severidad: 'danger',
      puntuacionOrden: 875_000,
      titulo: rotarFrase(`presup-tit-${ymd}`, [
        `Este mes llevas ${gTxt}${mS} gastados · disponible −${exT}${mS} 📉`,
        `Tu guía del mes se movió: gastado ${gTxt}${mS}, en rojo −${exT}${mS} 💛`,
        `Respira: ${gTxt}${mS} en salidas · margen −${exT}${mS} 🌿`,
        `Llevas ${gTxt}${mS} · el “disponible” quedó en −${exT}${mS} hoy 📊`,
        `Mes intenso: ${gTxt}${mS} gastado · −${exT}${mS} bajo tu guía ✨`,
        `Números claros: ${gTxt}${mS} fuera · −${exT}${mS} de colchón 🎯`,
        `Ojo al ritmo: ${gTxt}${mS} en gastos · −${exT}${mS} disponible 😊`,
        `Tu mes suma ${gTxt}${mS} en gastos · margen −${exT}${mS} 💪`,
        `Aquí va con cariño: ${gTxt}${mS} gastados, −${exT}${mS} libre 🤝`,
        `Línea roja suave: ${gTxt}${mS} vs guía · −${exT}${mS} 📌`,
      ]),
      detalle: rotarFrase(`presup-det-${ymd}`, [
        'No es fracaso: es información para decidir mejor mañana 🌅',
        'Revisa Inicio o Gastos · ahí ves el desglose con calma 📒',
        'El límite del mes lo ajustas en Saldo cuando quieras 💚',
        'Un ajuste pequeño hoy evita estrés después ✨',
        'Puedes frenar un gasto o subir la guía · tú eliges 💛',
        'Lo importante es que sigas mirando tus números 👀',
        'Cada mes es práctica; vas aprendiendo tu ritmo 🌟',
        'Tu app te acompaña, no te juzga 🤗',
        'Si el mes fue especial, sube la guía sin culpa ☕',
        'Paso a paso: un número honesto vale más que mil excusas 💬',
      ]),
    },
  ];
}

/**
 * Recordatorio amable para revisar cómo va el presupuesto del mes (sin superar el tope).
 * Un aviso por día; 10 textos de título y 10 de detalle vía rotarFrase.
 */
function notificacionesPresupuestoRevisar(state, ref) {
  const ahora = ref instanceof Date ? ref : new Date();
  const mes = ahora.getMonth();
  const anio = ahora.getFullYear();
  const tope = parseFloat(state?.presupuestoMensual) || 0;
  if (tope <= 0) return [];
  const gastos = state.gastos || [];
  const gastosMes = gastos.reduce(
    (s, g) => s + montoGastoCuentaParaPresupuestoEnMes(g, state, mes, anio),
    0
  );
  if (gastosMes > tope) return [];
  const ymd = `${anio}-${String(mes + 1).padStart(2, '0')}-${String(ahora.getDate()).padStart(2, '0')}`;
  const gTxt = formatearNumero(gastosMes, 0);
  const disp = Math.max(0, tope - gastosMes);
  const dispTxt = formatearNumero(disp, 0);
  const topeTxt = formatearNumero(tope, 0);
  const mon = (state.moneda && String(state.moneda).trim()) || '';
  const mS = mon ? ` ${mon}` : '';
  const pct = tope > 0 ? Math.min(100, Math.round((gastosMes / tope) * 100)) : 0;

  return [
    {
      id: `presupuesto-revisar-${ymd}`,
      tipo: 'presupuesto',
      severidad: 'info',
      puntuacionOrden: 265_000,
      titulo: rotarFrase(`presup-rev-tit-${ymd}`, [
        `¿Cómo va tu mes? Llevas ${gTxt}${mS} gastados · guía ${topeTxt}${mS} 📊✨`,
        `Momento ideal para mirar tu presupuesto: ${gTxt}${mS} de ${topeTxt}${mS} · ~${pct}% 🎯`,
        `Tu guía del mes: ${topeTxt}${mS} · gastado ${gTxt}${mS} · aún ${dispTxt}${mS} tranquilos 💛`,
        `Pequeño check-in: vas en ${gTxt}${mS} · te quedan ~${dispTxt}${mS} bajo el tope 😊`,
        `Respira y revisa: presupuesto ${topeTxt}${mS} · llevas ${gTxt}${mS} 🌿`,
        `Llevas el ritmo anotado: ${gTxt}${mS} · margen ~${dispTxt}${mS} hasta la guía 💪`,
        `Inicio y Gastos tienen el detalle · ${gTxt}${mS} de ${topeTxt}${mS} hasta ahora 📒`,
        `Cuidar el presupuesto es cuidarte · ${gTxt}${mS} gastados · ${dispTxt}${mS} disponibles 🤝`,
        `Un vistazo rápido: ~${pct}% de tu guía usada · ${gTxt}${mS} en salidas ☕`,
        `Tu mes en números amables: gastado ${gTxt}${mS} · tope ${topeTxt}${mS} ✨`,
      ]),
      detalle: rotarFrase(`presup-rev-det-${ymd}`, [
        'Mira Inicio o Gastos con calma; nadie corre contra ti 🌅',
        'Si el ritmo te gusta, sigue; si no, ajustas en Saldo sin culpas 💚',
        'Revisar hoy evita sorpresas al final del mes · tú mandas 📌',
        'La guía es tuya: puedes editarla en Saldo cuando lo necesites ✨',
        'Cada mirada a tus números es un punto para tu tranquilidad 🤗',
        'Pequeño hábito, gran claridad: un minuto y listo ⏱️',
        'Tu app celebra que quieras estar al tanto · sigue así 🌟',
        'Compara con tus metas del mes y celebra lo que ya avanzaste 🎉',
        'Organizar no es castigo; es cariño hacia tu futuro yo 💛',
        'Paso a paso: presupuesto revisado = mente más liviana 😌',
      ]),
    },
  ];
}

/**
 * Recordatorio cuando hay ítems en la lista del súper.
 * Tono según urgencia: con al menos un “urgente” → cercano y con prisa amable; solo “normal” → amigable;
 * todo “puede esperar” → relajado. Id fijo; rotarFrase con semilla por día y conteos.
 */
function notificacionesListaSuperCompras(state, ref) {
  const lineas = ordenarLineasListaSuper(state?.listaSuperCompraItems || []);
  const n = lineas.length;
  if (n === 0) return [];
  const nUrg = lineas.filter((l) => l.urgencia === 'urgente').length;
  const nEsp = lineas.filter((l) => l.urgencia === 'puede_esperar').length;
  const ahora = ref instanceof Date ? ref : new Date();
  const ymd = `${ahora.getFullYear()}-${String(ahora.getMonth() + 1).padStart(2, '0')}-${String(ahora.getDate()).padStart(2, '0')}`;
  const itm = n === 1 ? 'ítem' : 'ítems';
  /** Mezcla con urgente → tono “con prisa”; solo puede_esperar → relajado; si no, normal amigable. */
  const tono = nUrg > 0 ? 'urgente' : nEsp === n ? 'relajado' : 'normal';
  const semTit = `lista-sup-tit-${tono}-${ymd}-${n}-${nUrg}-${nEsp}`;
  const semDet = `lista-sup-det-${tono}-${ymd}-${n}-${nUrg}-${nEsp}`;

  const titulosUrgente = [
    `¡Uy! En tu lista hay ${n} ${itm} y ${nUrg === 1 ? '1 es urgente' : `${nUrg} son urgentes`} · mira ya 🛒💛`,
    `La lista te reclama con cariño: ${nUrg} cosas no pueden esperar mucho (de ${n} en total) ⏰`,
    `Sin drama feo, pero sí: hay ${nUrg} urgentes entre ${n} ${itm} · date una vuelta ⚡`,
    `Antes de que se te pase: ${nUrg} ítems urgentes en tu lista (${n} en total) 🔔`,
    `Tu súper te necesita: ${n} ${itm} y ${nUrg} piden prioridad hoy 💪`,
    `Pequeño empujón amable: ${nUrg} urgentes · lista de ${n} ${itm} ✨`,
    `Hoy conviene abrir la lista: ${nUrg} cosas urgentes de ${n} 👀`,
    `Corre suave pero no lo dejes: ${nUrg} urgentes en ${n} ${itm} 😅`,
    `Tus pendientes llaman: ${nUrg} urgentes de ${n} ${itm} en la lista 🎯`,
    `Atención suave: ${nUrg} con prisa entre ${n} ${itm} · vamos 🙏`,
  ];
  const detallesUrgente = [
    'Abre el carrito o Más → Asistente y tacha lo urgente primero 🛒',
    'Marcar comprado alivia; después puedes registrar en Gastos sin estrés 💚',
    'Un minuto ahora = menos culpas después · tú puedes ⏱️',
    'Nadie te regaña; la lista solo te recuerda con cariño 🤗',
    'Las urgentes van arriba en la app · fácil de ver 👆',
    'Si ya compraste algo urgente, marca comprado y respiras 😌',
    'Ir con la lista clara te ahorra vueltas y olvidos 🙌',
    'Paso a paso: mira, compra lo urgente, anota · listo ✨',
    'El mercado puede esperar un poco; lo urgente de la lista, menos ☕',
    'Organizarte un poquito antes de salir ya es ganar 💛',
  ];

  const titulosNormal = [
    `Tu lista del súper tiene ${n} ${itm} · échale un vistazo cuando puedas 🛒✨`,
    `Hay ${n} pendiente${n === 1 ? '' : 's'} por comprar · te va a ayudar al salir 💛`,
    `Recuerda lo que anotaste: ${n} ${itm} en la lista 📝`,
    `Antes de ir al mercado, revisa tus ${n} ${itm} 👀`,
    `Llevas ${n} ${itm} en la lista · todo en orden 😊`,
    `Recordatorio suave: ${n} cosa${n === 1 ? '' : 's'} en la lista del súper 🌿`,
    `Tu lista tiene ${n} ${itm} · organízate con calma 🧺`,
    `${n} ${itm} esperándote · un toque y los ves 📱`,
    `Buen momento para avanzar con ${n} ${itm} de la lista 🎯`,
    `Lista con ${n} ${itm} pendiente${n === 1 ? '' : 's'} · cuando quieras 🙌`,
  ];
  const detallesNormal = [
    'Abre el ícono del carrito o Más → Asistente de compras 🛒',
    'Marcar comprado quita el ítem; después puedes anotar en Gastos 💚',
    'Ir preparado al mercado ahorra tiempo y plata 🤝',
    'Tachar la lista es un mini logro; celebra cada ítem ✨',
    'Si ya no necesitas algo, bórralo sin culpa; la lista es tuya 🌟',
    'Un minuto de revisión = menos vueltas y menos olvidos ⏱️',
    'Comprar con idea clara también es cuidarte 💛',
    'Paso a paso: revisa, compra tranquilo, anota en Gastos 😌',
    'Tu yo del futuro agradece cuando llevas la lista clara 🌅',
    'Nada de presión; solo un recordatorio amable 💬',
  ];

  const titulosRelajado = [
    `Tu lista tiene ${n} ${itm} · pueden esperar; cuando quieras las ves 🌿`,
    `Sin apuro: ${n} cosa${n === 1 ? '' : 's'} en la lista para más adelante ☕`,
    `La lista va tranquila: ${n} ${itm} sin prisa 😌`,
    `Cuando quieras, revisa ${n} ${itm} del súper · todo puede esperar 🛋️`,
    `${n} ${itm} guardados en la lista · sin fecha apretada ✨`,
    `Todo relajado: ${n} pendiente${n === 1 ? '' : 's'} en la lista 🌤️`,
    `Tu súper puede esperar; ${n} ${itm} ahí nomás 💛`,
    `Lista tranquila con ${n} ${itm} · mírala cuando quieras 👋`,
    `Nada urgente: ${n} ${itm} anotados por si acaso 📝`,
    `A tu ritmo: ${n} ${itm} en la lista del súper 🐢`,
  ];
  const detallesRelajado = [
    'Abre el carrito o el Asistente cuando quieras; no hay apuro 🛒',
    'Si cambia algo, edita o borra sin culpa 🌟',
    'Cuando compres, marca comprado y listo 💚',
    'El mercado no huye; tú eliges el día 😊',
    'Revisar con calma también es orden ✨',
    'Gastos puede esperar al volver; sin estrés 💛',
    'La lista es tuya; nadie te corre 🤗',
    'Un café y después miras si quieres ☕',
    'Paso a paso, sin presión 🌿',
    'Asistente de compras te espera en Más cuando quieras 📱',
  ];

  const titulos =
    tono === 'urgente' ? titulosUrgente : tono === 'relajado' ? titulosRelajado : titulosNormal;
  const detalles =
    tono === 'urgente' ? detallesUrgente : tono === 'relajado' ? detallesRelajado : detallesNormal;

  return [
    {
      id: 'lista-super-pendiente',
      tipo: 'listaSuper',
      severidad: tono === 'urgente' ? 'warning' : 'info',
      puntuacionOrden: tono === 'urgente' ? 268_000 : 262_000,
      titulo: rotarFrase(semTit, titulos),
      detalle: rotarFrase(semDet, detalles),
    },
  ];
}

/**
 * Metas con objetivo > 0: recordatorio si falta avance, aviso si ≥80 %, celebración si 100 %.
 */
function notificacionesMetas(state, ref) {
  const metasRaw = state?.metas || [];
  const contrib = Array.isArray(state?.contribucionesMetas) ? state.contribucionesMetas : [];
  const metas = metasRaw.map(normalizarMeta).filter((m) => m.id && (parseFloat(m.objetivo) || 0) > 0);
  if (metas.length === 0) return [];

  const ahora = ref instanceof Date ? ref : new Date();
  const ymd = `${ahora.getFullYear()}-${String(ahora.getMonth() + 1).padStart(2, '0')}-${String(ahora.getDate()).padStart(2, '0')}`;
  const moneda = String(state?.moneda || '').trim();
  const suf = moneda ? ` ${moneda}` : '';
  const out = [];

  for (const m of metas) {
    const obj = parseFloat(m.objetivo) || 0;
    const acum = contrib
      .filter((c) => c && c.metaId === m.id)
      .reduce((s, c) => s + (parseFloat(c.cantidad) || 0), 0);
    const pct = obj > 0 ? Math.min(100, (acum / obj) * 100) : 0;
    const nom = String(m.nombre || 'Meta').trim() || 'Meta';
    const acumTxt = formatearNumero(acum, 0);
    const objTxt = formatearNumero(obj, 0);
    const pRound = Math.round(pct);

    if (pct >= 100) {
      const id = `meta-${m.id}-logro`;
      out.push({
        id,
        tipo: 'meta',
        severidad: 'info',
        puntuacionOrden: 718_000,
        titulo: rotarFrase(`${id}-t-${ymd}`, [
          `¡Cumpliste la meta «${nom}»! 🎉`,
          `«${nom}» al 100 % · lo lograste ✨`,
          `Meta «${nom}» completada · bravo 💪`,
          `Tope alcanzado en «${nom}» · celebra un momento 🌟`,
          `«${nom}» cerrada con éxito · gran trabajo 🤝`,
        ]),
        detalle: rotarFrase(`${id}-d-${ymd}`, [
          `Llevas ${acumTxt}${suf} de ${objTxt}${suf}. Sigue así en Más → Metas.`,
          `${acumTxt}${suf} sobre ${objTxt}${suf} · tu esfuerzo se nota 💛`,
          `Objetivo cumplido: ${acumTxt}${suf}. ¿Otra meta cuando quieras?`,
          `Números redondos: ${acumTxt}${suf} / ${objTxt}${suf} · disfrútalo 😊`,
          `Un logro más en el bolsillo · ${acumTxt}${suf} logrados 🙌`,
        ]),
      });
      continue;
    }

    if (pct >= 80) {
      const id = `meta-${m.id}-cerca`;
      out.push({
        id,
        tipo: 'meta',
        severidad: 'warning',
        puntuacionOrden: 712_000,
        titulo: rotarFrase(`${id}-t-${ymd}`, [
          `Casi «${nom}» · vas al ${pRound}% 🎯`,
          `«${nom}» a un paso · ${pRound}% listo ⚡`,
          `Falta poquito para «${nom}» · ${pRound}% 🌿`,
          `Meta «${nom}» casi cerrada · ${pRound}% 💛`,
          `Último tramo de «${nom}» · ${pRound}% ✨`,
        ]),
        detalle: rotarFrase(`${id}-d-${ymd}`, [
          `Llevas ${acumTxt}${suf} de ${objTxt}${suf}. Un aporte más y la cierras.`,
          `${acumTxt}${suf} sobre ${objTxt}${suf} · sigue en Más → Metas cuando puedas.`,
          `Estás muy cerca: ${acumTxt}${suf} de ${objTxt}${suf} · ánimo 💪`,
          `Un empujón amable y «${nom}» queda lista 😊`,
          `Tu avance se ve: ${acumTxt}${suf} / ${objTxt}${suf} · casi 🎉`,
        ]),
      });
      continue;
    }

    const id = `meta-${m.id}-avance`;
    out.push({
      id,
      tipo: 'meta',
      severidad: 'info',
      puntuacionOrden: 395_000,
      titulo: rotarFrase(`${id}-t-${ymd}`, [
        `Tu meta «${nom}» te recuerda con cariño 🌿`,
        `¿Un momento para «${nom}»? 💛`,
        `«${nom}» sigue abierta · avanza cuando puedas ✨`,
        `Pequeño recordatorio: meta «${nom}» 📌`,
        `No te olvides de «${nom}» · va en ${pRound}% 🎯`,
      ]),
      detalle: rotarFrase(`${id}-d-${ymd}`, [
        `Vas en ${acumTxt}${suf} de ${objTxt}${suf}. Aporta desde Más → Metas o al registrar movimientos.`,
        `${acumTxt}${suf} / ${objTxt}${suf} · cada aporte suma 💪`,
        `Llevas ${pRound}% · un pasito hoy cuenta 😊`,
        `Revisa «${nom}» en Metas y suma lo que puedas 🤝`,
        `Las metas son tuyas; la app solo te acompaña 💬`,
      ]),
    });
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
  const ahora = new Date();
  const ymd = `${ahora.getFullYear()}-${String(ahora.getMonth() + 1).padStart(2, '0')}-${String(ahora.getDate()).padStart(2, '0')}`;
  const liquido = totalSaldoLiquido(state);
  const q = formatearNumero(liquido, 0);
  if (liquido <= 0) {
    return [
      {
        id: 'saldo-liquido-critico',
        tipo: 'saldo',
        severidad: 'danger',
        puntuacionOrden: 920_000,
        titulo: rotarFrase(`sal-cri-t-${ymd}`, [
          `Tu efectivo y cuentas (sin cupo de tarjeta) van en ${q} 🏦`,
          `Caja al día muy justa: ${q} en líquido · respira y revisa 🤍`,
          `Líquido total ${q} · es buen momento para mirar Saldo con calma 🌿`,
          `Efectivo + banco + apps: ${q} · el cupo de la TC va aparte 💳`,
          `Número honesto: ${q} disponible en líquido hoy 📊`,
          `Ojo cariñoso: el líquido quedó en ${q} 💛`,
          `Tu “plata al día” marca ${q} · revisemos juntos en Saldo 🤝`,
          `Momento de alinear: líquido ${q} según la app ✨`,
          `Aquí vamos con apoyo: ${q} en caja real (sin cupo TC) 💪`,
          `Pequeña alerta amable: líquido en ${q} 🔔`,
        ]),
        detalle: rotarFrase(`sal-cri-d-${ymd}`, [
          'Entra a Saldo: ingreso, gasto o ajuste · tú tienes el control 💚',
          'Que coincida con tu bolsillo real · sin miedo a corregir 😊',
          'La tarjeta en cupo es otro capítulo · aquí va el efectivo 🤍',
          'Un ajuste hoy te da claridad mañana 🌅',
          'Si hay error de cifra, corregir es un acto de orden ✨',
          'Paso a paso: revisa efectivo, banco y apps que usas 📱',
          'No estás solo en esto; la app solo muestra lo que anotaste 🤗',
          'Pequeño chequeo y vuelves a la tranquilidad 😌',
          'Cuidar el líquido es cuidar tu paz 💛',
          'Saldo te espera con todo lo que necesitas para cuadrar 📌',
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
        titulo: rotarFrase(`sal-baj-t-${ymd}`, [
          `Tienes poco colchón en líquido: ~${q} · vamos con calma 🌤️`,
          `Caja algo justa (~${q}) · buen momento para revisar Saldo 💛`,
          `~${q} en efectivo y cuentas · el cupo de la TC es aparte 💳`,
          `Margen fino: ~${q} líquido · sin drama, solo atención suave 🔔`,
          `Tu “plata al día” está en ~${q} · un vistazo y listo 👀`,
          `Líquido ~${q} · pequeño recordatorio para cuidarte 🌿`,
          `Vas con ~${q} en caja real · organiza con cariño 🤍`,
          `Número a vigilar: ~${q} en líquido 📊`,
          `Hay ~${q} antes de tocar ahorros fuertes · piénsalo tranquilo ☕`,
          `Colchón bajito (~${q}) · tú decides el siguiente paso 💪`,
        ]),
        detalle: rotarFrase(`sal-baj-d-${ymd}`, [
          'Saldo e Ingresos te ayudan a ver el panorama completo 📒',
          'Un ingreso anotado o un gasto menos cambia el ánimo ✨',
          'Ajustar saldos iniciales también cuenta · sin culpas 😊',
          'Pequeños cambios hoy = más paz mañana 🌅',
          'La TC en cupo no entra aquí · revisa su ficha aparte 💚',
          'Tú marcas el ritmo; esto es solo una luz amarilla 🚦',
          'Si puedes, aparta algo antes del próximo gasto 💛',
          'Mirar los números ya es un gran paso 👏',
          'La app te respalda para ordenar con claridad 🤗',
          'Un minuto en Saldo y retomas el control 📌',
        ]),
      },
    ];
  }
  return [];
}

/**
 * Saldo “disponible” en cupo vía app + alertas de resumen (pago/corte, uso alto, ~20% libre).
 */
function notificacionesTarjetas(state, ref) {
  const ymd = `${ref.getFullYear()}-${String(ref.getMonth() + 1).padStart(2, '0')}-${String(ref.getDate()).padStart(2, '0')}`;
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
      titulo: rotarFrase(`tcsin-t-${ymd}`, [
        `Cupo libre en cero según la app · tope ${limTxt} 💳`,
        `La TC no muestra cupo disponible ahora · límite ${limTxt} 🔔`,
        `Sin cupo libre en el registro · ${limTxt} de tope 😊`,
        `Momento de revisar la tarjeta: cupo 0 · ${limTxt} anotado ✨`,
        `Tu TC en la app dice “sin aire” · tope ${limTxt} 💛`,
        `Ojo amable: cupo agotado en la app · ${limTxt} 🌿`,
        `Aquí va con calma: 0 libre · límite ${limTxt} 🤝`,
        `Revisa números con el banco · tope app ${limTxt} 📊`,
        `La ficha pide un chequeo · ${limTxt} de límite 💪`,
        `Sin margen en cupo (app) · ${limTxt} 🎯`,
      ]),
      detalle: rotarFrase(`tcsin-d-${ymd}`, [
        'Abre Saldo → Tarjeta y compáralo con el banco · sin estrés 💚',
        'Si el banco muestra otra cifra, aquí la corriges en un momento 📱',
        'Alinear app y banco es un acto de tranquilidad 😌',
        'Límite y deuda bien anotados = menos sorpresas ✨',
        'Tú puedes dejarlo perfecto; la app te guía 🤗',
        'Un ajuste pequeño y el cupo vuelve a tener sentido 💛',
        'No pasa nada: los números a veces necesitan un cariñito ☕',
        'Revisa compras y abonos anotados · todo encaja 🧩',
        'Paso a paso en Más → Saldo → Tarjeta 📌',
        'Cuidar la TC es cuidar tu futuro yo 💪',
      ]),
    });
  } else if (limite > 0 && r.global.porcentaje >= 90 && (r.tarjetas || []).length === 0) {
    const pg = r.global.porcentaje.toFixed(0);
    out.push({
      id: 'tc-cupo-alto',
      tipo: 'tc',
      severidad: 'warning',
      puntuacionOrden: 420_000,
      titulo: rotarFrase(`tcglo-t-${ymd}`, [
        `Casi todo el cupo global está en uso (~${pg}%) · respira y revisa 💳`,
        `Tu TC global va ~${pg}% · buen momento para mirar Saldo 😊`,
        `Queda poquito aire en el cupo total (~${pg}%) ✨`,
        `~${pg}% usado en conjunto · nada grave, solo atención 💛`,
        `El cupo global pide calma: ~${pg}% ocupado 🌿`,
        `Vas ~${pg}% en tarjetas · un vistazo y decides tranquilo 👀`,
        `Línea amarilla suave: ~${pg}% del cupo total 🚦`,
        `Tu plástico global está cargado (~${pg}%) 💪`,
        `Momento ideal para planear el siguiente pago (~${pg}%) 📊`,
        `~${pg}% · pequeño recordatorio antes de otra compra 🔔`,
      ]),
      detalle: rotarFrase(`tcglo-d-${ymd}`, [
        'Saldo te muestra el panorama · tú eliges abonar o esperar 💚',
        'Un abono cariñoso baja la tensión del cierre 🙌',
        'Si el tope en app no cuadra, ajústalo sin culpas ✨',
        'Antes de comprar: mira bolsillo y cupo con honestidad 🤍',
        'Organizar la TC es un regalo para tu mes que viene 🎁',
        'Puedes bajar ritmo o subir pago · ambas son válidas 💛',
        'La app te acompaña; el banco manda en la cifra final 📌',
        'Pequeño plan hoy = menos estrés mañana 🌅',
        'Tú mandas en tu dinero; esto es solo información 🤗',
        'Un paso en Saldo y retomas claridad 😌',
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
        titulo: pagoSug
          ? rotarFrase(`${idEx}-t-${ymd}`, [
              `Hoy es corte con ${t.nombreEntidad} · guía ~${pagoSug} 💳✨`,
              `Llegó el corte de ${t.nombreEntidad} · idea de pago ~${pagoSug} 📅`,
              `${t.nombreEntidad} cierra ciclo hoy · ~${pagoSug} como referencia 💛`,
              `Día de corte · ${t.nombreEntidad} · monto guía ~${pagoSug} 🎯`,
              `Tu TC ${t.nombreEntidad} pide atención hoy · ~${pagoSug} 🤝`,
              `Corte amable: ${t.nombreEntidad} · sugerido ~${pagoSug} ☕`,
              `Hoy alinea ${t.nombreEntidad} con el banco · ~${pagoSug} 📊`,
              `Momento de revisar ${t.nombreEntidad} · guía ~${pagoSug} 💪`,
              `Cierre del periodo · ${t.nombreEntidad} · ~${pagoSug} 🌿`,
              `Buen día para ordenar ${t.nombreEntidad} · ~${pagoSug} 🔔`,
            ])
          : rotarFrase(`${idEx}-t0-${ymd}`, [
              `Hoy es corte con ${t.nombreEntidad} · mira el extracto 💳`,
              `${t.nombreEntidad} · día de cerrar ciclo · abre Saldo ✨`,
              `Llegó el corte · ${t.nombreEntidad} te espera en la app 😊`,
              `Tu tarjeta ${t.nombreEntidad} hoy pide un vistazo 👀`,
              `Corte hoy · ${t.nombreEntidad} · sin estrés, con calma 🌿`,
              `Momento de alinear ${t.nombreEntidad} con el banco 🤝`,
              `${t.nombreEntidad} · revisa cupo y movimientos 💛`,
              `Hoy toca ser amigo de tu extracto · ${t.nombreEntidad} 📊`,
              `Pequeño recordatorio de corte · ${t.nombreEntidad} 🔔`,
              `Abre la ficha y celebra llevar el control 💪`,
            ]),
        detalle: pagoSug
          ? rotarFrase(`${idEx}-d-${ymd}`, [
              `Capital al cierre ~${cierreTxt || '—'} · intereses ~${intTxt || '0'} · confirma en banco 🏦`,
              'La app estima; el banco tiene la palabra final · revisa con cariño 💚',
              'Hay recordatorio en Pagos programados si lo configuraste 📝',
              'Si ya pagaste, anótalo y respira tranquilo 😌',
              'Cuadrar números hoy = menos sorpresas mañana ✨',
              'Un minuto en Saldo y todo cobra sentido 📌',
              'Los intereses son aproximados · valida con tu extracto 📊',
              'Tú puedes dejarlo perfecto paso a paso 🤗',
              'Registrar el pago mantiene viva tu claridad 💛',
              'Pequeño esfuerzo, gran paz mental 🌅',
            ])
          : rotarFrase(`${idEx}-d0-${ymd}`, [
              'En Saldo ves cupo, cuotas y el detalle completo 💳',
              'Abre el extracto y alinea con lo que dice el banco 🤝',
              'Hoy es buen día para sentirte dueño de tus números 💪',
              'Revisa movimientos con calma; no hay prisa fea ☕',
              'La app está para ayudarte, no para presionarte 😊',
              'Un vistazo al extracto y sigues con tu día ✨',
              'Organizar la TC es un acto de autocuidado 💚',
              'Paso a paso: cupo, deuda, próximo pago 📊',
              'Aquí vamos contigo · mira la ficha de la tarjeta 🌿',
              'Claridad hoy, tranquilidad mañana 🌅',
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
        titulo: pagoL
          ? rotarFrase(`${id}-t-${ymd}`, [
              `¡Vamos! ${t.nombreEntidad} · pago ~${pagoL} en ${dP} día${dP === 1 ? '' : 's'} 💳✨`,
              `Cuenta regresiva amable: ${dP} día${dP === 1 ? '' : 's'} · ~${pagoL} · ${t.nombreEntidad} ⏳`,
              `Pronto toca pagar ${t.nombreEntidad} · guía ~${pagoL} 💛`,
              `${t.nombreEntidad} te espera en ${dP} día${dP === 1 ? '' : 's'} · ~${pagoL} 🎯`,
              `Marca el calendario: ${t.nombreEntidad} · ~${pagoL} 📅`,
              `Un empujón cariñoso: pago ~${pagoL} · ${t.nombreEntidad} 🤝`,
              `Vence ${t.etiquetaProxPago || 'pronto'} · ~${pagoL} · ${t.nombreEntidad} 😊`,
              `Tú puedes organizarlo: ${dP} día${dP === 1 ? '' : 's'} · ${t.nombreEntidad} 💪`,
              `Idea de monto ~${pagoL} · confirma con el banco 🏦`,
              `Respira y planea: ${t.nombreEntidad} · ~${pagoL} 🌿`,
            ])
          : rotarFrase(`${id}-t0-${ymd}`, [
              `En ${dP} día${dP === 1 ? '' : 's'} toca pagar ${t.nombreEntidad} · mira Saldo 💳`,
              `Se acerca el pago a ${t.nombreEntidad} · vamos con calma ⏳`,
              `${t.nombreEntidad} · ${dP} día${dP === 1 ? '' : 's'} para organizarte 💛`,
              `Cuenta atrás: ${t.nombreEntidad} · ${dP} día${dP === 1 ? '' : 's'} 📅`,
              `Buen momento para abrir el extracto · ${t.nombreEntidad} 👀`,
              `Tu TC pide atención en ${dP} día${dP === 1 ? '' : 's'} · ${t.nombreEntidad} ✨`,
              `Nada de drama: solo planear ${t.nombreEntidad} 🤗`,
              `Pago próximo · ${t.nombreEntidad} · revisa fechas 😊`,
              `Un paso hoy evita estrés mañana · ${t.nombreEntidad} 💪`,
              `Aquí vamos contigo · Saldo tiene el detalle 🌿`,
            ]),
        detalle: pagoL
          ? rotarFrase(`${id}-d-${ymd}`, [
              'La app sugiere; el banco confirma · ambos son equipo contigo 🤝',
              'Cuadra con tu extracto y sonríe al anotar el pago 😊',
              'Evitar la fecha es cuidarte a ti mismo 💚',
              'Registrar en Gastos mantiene tu historia clara 📒',
              'Monto guía ~ referencia amable, no regla dura 💛',
              'Un pago a tiempo es un regalo para tu mes que viene 🎁',
              'Si ya pagaste, anótalo y celebra el orden 🎉',
              'Pequeño hábito, gran tranquilidad ✨',
              'Tú mandas; esto solo ilumina el camino 🔦',
              'Saldo y Pagos programados te acompañan 📌',
            ])
          : rotarFrase(`${id}-d0-${ymd}`, [
              'Abre Saldo y mira fechas sin presión fea 😌',
              'El extracto es tu mejor amigo hoy 📊',
              'Pagar a tiempo es un acto de paz 🌅',
              'La ficha de la TC tiene todo lo que necesitas 💳',
              'Evitar mora también es autocuidado 💪',
              'Un minuto de revisión vale horas de calma ⏱️',
              'Vas bien; solo falta alinear con el banco 🤝',
              'Aquí nadie juzga · solo te recordamos con cariño 💬',
              'Organizar hoy es dormir mejor mañana 😴💚',
              'Paso a paso, sin culpas ✨',
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
        titulo: corteSug
          ? rotarFrase(`${id}-t-${ymd}`, [
              `En ${t.diasCorte} día${t.diasCorte === 1 ? '' : 's'} llega el corte · ${t.nombreEntidad} · guía ~${corteSug} 📅💳`,
              `Corte a la vuelta · ${t.nombreEntidad} · idea ~${corteSug} ✨`,
              `${t.nombreEntidad} cierra pronto · ~${corteSug} como referencia 💛`,
              `Prepárate con calma: corte en ${t.diasCorte} d. · ${t.nombreEntidad} 🤝`,
              `Faltan ${t.diasCorte} d. · ${t.nombreEntidad} · monto guía ~${corteSug} 🎯`,
              `Tu TC avisa con tiempo · ${t.nombreEntidad} · ~${corteSug} 🌿`,
              `Buen momento para mirar movimientos · ${t.nombreEntidad} 👀`,
              `Organiza antes del corte · ~${corteSug} · ${t.nombreEntidad} 💪`,
              `Pequeña alerta amable · ${t.diasCorte} d. · ${t.nombreEntidad} 🔔`,
              `Planear hoy = tranquilidad después · ~${corteSug} ☕`,
            ])
          : rotarFrase(`${id}-t0-${ymd}`, [
              `En ${t.diasCorte} día${t.diasCorte === 1 ? '' : 's'} es corte con ${t.nombreEntidad} 💳`,
              `${t.nombreEntidad} · el calendario avisa con cariño 📅`,
              `Corte cerca · ${t.nombreEntidad} · revisa Saldo 😊`,
              `Aún hay tiempo · ${t.nombreEntidad} · mira el extracto ✨`,
              `${t.nombreEntidad} · ${t.diasCorte} d. para organizarte 💛`,
              `Tu tarjeta te recuerda: corte pronto · ${t.nombreEntidad} 🔔`,
              `Sin estrés: abre la ficha y mira números 📊`,
              `Pequeño recordatorio · ${t.nombreEntidad} · ${t.etiquetaProxCorte || 'Saldo'} 🌿`,
              `Vas bien; solo alinea con el banco 🤝`,
              `Claridad hoy, mejor cierre mañana 🌅`,
            ]),
        detalle: rotarFrase(`${id}-c-${ymd}`, [
          'Gastos, Saldo y Pagos programados son tu trío de apoyo 📒💚',
          'Confirma la cifra con el banco y sonríe 😊',
          'Las cuotas en Gastos cuentan la historia real ✨',
          'Si cambia la fecha, edita el recordatorio sin drama 📝',
          'El extracto en Saldo es tu brújula amable 🧭',
          'Un minuto de orden vale horas de paz ⏱️',
          'Tú decides el ritmo; la app ilumina el camino 🔦',
          'Registrar compras a cuota ayuda al cierre 🤝',
          'Pequeños datos bien puestos = gran tranquilidad 💛',
          'Aquí vamos contigo, paso a paso 🌿',
        ]),
      });
    } else if (t.alertaUtil) {
      const ratioLibre = t.cupoTotal > 0 ? libre / t.cupoTotal : 0;
      const pU = t.utilPct.toFixed(0);
      const pLibre = (ratioLibre * 100).toFixed(0);
      if (libre > 0 && ratioLibre <= TC_CUPO_LIBRE_AVISO_MAX) {
        const id20 = `tc-cupo-20-${t.id}`;
        const lTxt = formatearNumero(libre, 0);
        const totTxt = formatearNumero(t.cupoTotal, 0);
        out.push({
          id: id20,
          tipo: 'tc',
          severidad: 'warning',
          tarjetaId: t.id,
          puntuacionOrden: 335_000,
          titulo: rotarFrase(`${id20}-t-${ymd}`, [
            `${t.nombreEntidad} · te queda ~${pLibre}% libre (${lTxt} de ${totTxt}) 💳`,
            `Cupo casi lleno en ${t.nombreEntidad} · aún hay ${lTxt} libres 💛`,
            `Respira: en ${t.nombreEntidad} el aire es ~${pLibre}% · ${lTxt} sueltos 🌿`,
            `Momento de decidir con calma · ${t.nombreEntidad} · ${lTxt} libre 😊`,
            `Tu TC ${t.nombreEntidad} va ajustada · ${lTxt} de ${totTxt} ✨`,
            `Pequeña luz amarilla · ${t.nombreEntidad} · poco cupo 🚦`,
            `Vas con ~${pLibre}% de colchón · ${t.nombreEntidad} 💪`,
            `${lTxt} libres en ${t.nombreEntidad} · buen día para planear 📊`,
            `Nada grave: solo atención suave a ${t.nombreEntidad} 🤝`,
            `El cupo pide cariño · ${t.nombreEntidad} · ${lTxt} libre ☕`,
          ]),
          detalle: rotarFrase(`${id20}-d-${ymd}`, [
            'Puedes pausar compras, abonar o subir el tope en Saldo · tú eliges 💚',
            'Un abono aunque sea pequeño baja la tensión 🙌',
            'Si el banco muestra otro disponible, alinea la ficha sin culpas ✨',
            'Antes del corte, mira el extracto con un café ☕',
            'Planear es poder; aquí tienes la info para hacerlo 💛',
            'La app no juzga; solo te acompaña con claridad 🤗',
            'Menos compra hoy = más paz al cierre 🌅',
            'Subir el tope en app si tu banco ya lo hizo también vale 📌',
            'Cuidar el cupo es cuidar tu mes que viene 💪',
            'Paso a paso en Saldo y todo mejora 😌',
          ]),
        });
      } else {
        const idU = `tc-uso-${t.id}`;
        out.push({
          id: idU,
          tipo: 'tc',
          severidad: 'info',
          puntuacionOrden: 320_000,
          titulo: rotarFrase(`${idU}-t-${ymd}`, [
            `${t.nombreEntidad} · llevas ~${pU}% del cupo usado · vas en ritmo 📊`,
            `Tu TC ${t.nombreEntidad} marca ~${pU}% · buen momento para mirar Saldo 😊`,
            `~${pU}% ocupado en ${t.nombreEntidad} · información, no regaña 💛`,
            `El cupo de ${t.nombreEntidad} va ~${pU}% · sigue así de consciente ✨`,
            `Línea informativa: ~${pU}% con ${t.nombreEntidad} 💳`,
            `${t.nombreEntidad} · ~${pU}% · pequeño recordatorio amable 🔔`,
            `Vas al día con tus números: ~${pU}% en ${t.nombreEntidad} 💪`,
            `Tu plástico habla: ~${pU}% usado · ${t.nombreEntidad} 🌿`,
            `Nada de miedo: ~${pU}% · solo claridad 🤝`,
            `Momento de celebrar que miras tu TC: ~${pU}% · ${t.nombreEntidad} 🎯`,
          ]),
          detalle: rotarFrase(`${idU}-d-${ymd}`, [
            'Saldo te muestra ritmo y límites · tú ajustas con calma 💚',
            'Puedes abonar, bajar compras o editar el tope ✨',
            'Menos cuotas nuevas = más aire al cierre 🌤️',
            'Dejar colchón también es inteligencia emocional con plata 💛',
            'La ficha de la tarjeta es tu mejor vista 👀',
            'Organizar no es castigo; es cariño hacia ti 🤗',
            'Un vistazo hoy evita sustos mañana 📅',
            'Tú mandas el ritmo de gasto; la app solo refleja 💬',
            'Pequeños ajustes hoy, gran diferencia al mes siguiente 🌅',
            'Sigue así de atento; eso ya es un hábito ganador 🏆',
          ]),
        });
      }
    } else if (libre > 0 && libre < t.cupoTotal * 0.05) {
      const idL = `tc-libre-${t.id}`;
      const lTxt = formatearNumero(libre, 0);
      const totTxt = formatearNumero(t.cupoTotal, 0);
      out.push({
        id: idL,
        tipo: 'tc',
        severidad: 'info',
        puntuacionOrden: 300_000,
        titulo: rotarFrase(`${idL}-t-${ymd}`, [
          `${t.nombreEntidad} · queda poquito libre: ${lTxt} de ${totTxt} 💳`,
          `Margen fino en ${t.nombreEntidad} · ${lTxt} sueltos 💛`,
          `Tu TC ${t.nombreEntidad} casi llena · ${lTxt} libre ✨`,
          `Ojo cariñoso: ${lTxt} libre en ${t.nombreEntidad} (tope ${totTxt}) 🔔`,
          `${t.nombreEntidad} pide calma al gastar · ${lTxt} disponibles 🌿`,
          `Pequeño colchón: ${lTxt} sobre ${totTxt} · ${t.nombreEntidad} 😊`,
          `Vas ajustado pero bien informado · ${t.nombreEntidad} 📊`,
          `${lTxt} libres · ${t.nombreEntidad} · decide con tranquilidad ☕`,
          `La app avisa suave: poco cupo en ${t.nombreEntidad} 💪`,
          `Aún hay ${lTxt} · úsalos con intención · ${t.nombreEntidad} 🎯`,
        ]),
        detalle: rotarFrase(`${idL}-d-${ymd}`, [
          'Compara con el banco en Saldo · sin prisa fea 🏦',
          'El extracto y la app deben ser amigos 🤝',
          'El “disponible” del banco a veces tarda · paciencia 💚',
          'Poco colchón · abona o evita un gasto extra si puedes 💛',
          'Corrige tope o saldo usado si ves diferencia ✨',
          'Cuidar el cupo es cuidar tu tranquilidad 😌',
          'Un minuto en la ficha y todo cobra sentido 📌',
          'Tú mandas; esto es solo una luz suave 🔦',
          'Organizar hoy es dormir mejor mañana 🌅',
          'Aquí vamos contigo · Saldo tiene el detalle 🤗',
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
  const p = notificacionesPresupuestoMensual(state, ref);
  const pr = notificacionesPresupuestoRevisar(state, ref);
  const ls = notificacionesListaSuperCompras(state, ref);
  const nm = notificacionesMetas(state, ref);
  const gm = notificacionesGastosMovimiento(state, ref);
  const items = [...a, ...b, ...s, ...c, ...p, ...pr, ...ls, ...nm, ...gm]
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
    items: deduped.map(({ puntuacionOrden: _p, ...rest }) => ({
      ...rest,
      titulo: tituloNotifConNombre(state, rest.titulo),
    })),
    total: deduped.length,
  };
}
