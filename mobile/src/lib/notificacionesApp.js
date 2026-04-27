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

/** Aviso en campana en los últimos días: id por día = recordatorio “diario” al abrir. */
const DIAS_RECORDATORIO_CAMPANA = 3;
/** Bajo esto, aviso de “poco en efectivo/cuentas” (sin contar cupo de tarjeta). */
const LIQ_BAJO_UMBRAL = 100_000;
/** Aviso “cupo al filo”: queda poco libre respecto al tope (≤32%). Incluye 5%, 10%, 20%…; sin piso mínimo para no omitir casos como 50.000 libres con tope alto. */
const TC_CUPO_LIBRE_AVISO_MAX = 0.32;

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
 * Exportado para sincronizar notificaciones locales del sistema.
 */
export function diasHastaPagoProgramado(pago, ref = new Date()) {
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
          `${concepto}: pendiente de anotar en el mes (Gastos).`,
          `Falta registrar ${concepto} en el mes. Entra a Gastos.`,
          `${concepto}: aún no lo anotas en el mes. Revisa Gastos.`,
          `Tienes ${concepto} sin registrar en Gastos para este mes.`,
          `En Gastos falta ${concepto} de este mes. Regístralo cuando pagues.`,
        ]),
        detalle: rotarFrase(`${id}-d-${ymd}`, [
          'Bloque de pagos programados o formulario. Tú eliges el momento.',
          'Un registro y el resumen acompaña a lo real.',
          'Pagos programados o pantalla de gasto: el que te vaya mejor.',
          'Anótalo en Gastos y el resumen queda alineado.',
          'Desde Pagos programados o Gastos, en el momento que elijas.',
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
        titulo: rotarFrase(`${id}-t-${ymd}`, [
          `Pago único ${concepto}: la fecha ya pasó.`,
          `La fecha de ${concepto} (pago único) ya quedó atrás.`,
          `${concepto}: el pago único venció en el calendario.`,
          `Pasó la fecha del pago único «${concepto}».`,
          `${concepto} (único): revisa la fecha en Pagos programados.`,
        ]),
        detalle: rotarFrase(`${id}-d-${ymd}`, [
          'Más → Pagos programados: fecha o borrado. Tú eliges.',
          'Ajusta o quita en Pagos programados.',
          'Corrige la fecha o bórralo si ya no aplica.',
          'Entra a Pagos programados y actualiza o quita el aviso.',
          'En Pagos programados lo dejas al día con un par de toques.',
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
          `Hoy vence: «${concepto}» (queda 0 días de plazo).`,
          `Último día de plazo: ${concepto}. Regístralo hoy en Gastos.`,
          `Hoy es el día: ${concepto}. Plazo 0.`,
          `Vence hoy ${concepto}. Anótalo o paga y registra en Gastos.`,
          `Plazo hoy: «${concepto}». Gastos te espera para el registro.`,
        ]),
        detalle: rotarFrase(`${id}-d-${ymd}`, [
          'Tiempo: vence hoy. Si ya pagaste, anótalo en Gastos. Si aún no, hoy es el día para realizarlo.',
          'Si pagaste, regístralo hoy en Gastos; si no, paga hoy y luego anota.',
          'Hoy cierra el plazo: luego monto y cuenta en Gastos.',
          'Mismo día: pago o registro en Gastos para no perder el hilo.',
          'Vence hoy: cumple y anótalo en Gastos con el detalle que toque.',
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
          `${faltanTxt}: «${concepto}»`,
          `«${concepto}» · ${d} día${d === 1 ? '' : 's'} para el vencimiento.`,
          `Quedan ${d} día${d === 1 ? '' : 's'} para ${concepto}.`,
          `En ${d} día${d === 1 ? '' : 's'} toca: ${concepto}.`,
          `Cerca el plazo: ${concepto} · ${d} día${d === 1 ? '' : 's'}.`,
        ]),
        detalle: rotarFrase(`${idCerca}-d`, [
          `Tiempo: ${plazoTxt} para el pago. Cumple y luego anótalo en Gastos. Si ya pagaste, regístralo con mismo monto y el programado baja solo.`,
          `Tienes ${plazoTxt}. Paga a tiempo, regístralo en Gastos y alinea cuenta y concepto.`,
          `${plazoTxt} para actuar. Después, Gastos. Si pagaste, un registro y listo el programado cuando coincida.`,
          `Plazo: ${plazoTxt}. Que no se pase; en Gastos queda trazado igual que en la vida real.`,
          `Reloj: ${plazoTxt}. Pago, registro; si pagaste, mismos datos en Gastos y el aviso acompaña.`,
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
      /** 5+5 frases; la fecha en la semilla rota el “aviso al azar” por día, estable el mismo día. */
      const ymd = `${y}-${String(m + 1).padStart(2, '0')}-${String(ahora.getDate()).padStart(2, '0')}`;
      out.push({
        id,
        tipo: 'categoria',
        severidad: 'warning',
        puntuacionOrden: 500_000,
        titulo: rotarFrase(`${id}-t-${ymd}`, [
          `Uff, ${cat.nombre} ${cat.icono}: te pasaste del límite (gastos ${gTxt}, tope ${lTxt})\n😱 😉`,
          `Uff, ${cat.nombre} ${cat.icono}: pasaste el tope, gastos ${gTxt} y el límite era ${lTxt}\n😱 😉`,
          `Uff, en ${cat.nombre} ${cat.icono} te pasaste: ${gTxt} y el tope es ${lTxt}\n😱 😉`,
          `Uff, ojo: ${cat.nombre} ${cat.icono} ya pasó el tope del mes (${gTxt} / ${lTxt})\n😱 😉`,
          `Uff, ${cat.nombre} ${cat.icono}: ${gTxt} de gasto y el tope es ${lTxt}, te pasaste\n😱 😉`,
        ]),
        detalle: rotarFrase(`${id}-d-${ymd}`, [
          `Cambia el límite en Más → Categorías o baja un poco los gastos de este mes\n😱 😉`,
          `Puedes subir el tope en Categorías o gastar menos en lo que queda del mes\n😱 😉`,
          `Cambia el tope en Más → Categorías o cuida un poco lo que sigues gastando en el mes\n😱 😉`,
          `Más → Categorías: sube el tope o baja lo que aún quieres gastar este mes\n😱 😉`,
          `Ajusta el tope en Categorías o baja un poco el gasto; en Más → Categorías\n😱 😉`,
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
          `Efectivo + bancos + apps: ${q} (sin cupo de tarjeta en esta suma).`,
          `Plata al día: ${q}. Revisar cajas reales en Saldo.`,
          `Efectivo y cuentas: ${q} (sin contar el cupo de la TC).`,
          `Total líquido (sin TC anotada en cupo): ${q}. Mira en Saldo.`,
          `Cajas a la mano: ${q}. Tarjeta en cupo, aparte, en el resumen de Saldo.`,
        ]),
        detalle: rotarFrase(`sal-cri-d-${ymd}`, [
          'Saldo: ingresos, gasto o ajuste de saldos. Tú diriges.',
          'Que cuadre con el bolsillo, no solo con el cupo de tarjeta.',
          'Ajusta con ingresos, gasto o cifras iniciales en Saldo si hace falta.',
          'Entra a Saldo y alinea cajas reales; suma billeteras y banco.',
          'Poco o nada: revisa Efectivo, banco y apps que uses en la app.',
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
          `Poco en efectivo/cuentas: unos ${q} (cupo de tarjeta aparte).`,
          `Caja al día: ~${q} · cuida el margen.`,
          `Queda poco líquido: alrededor de ${q} (TC por aparte en Saldo).`,
          `Ojo: ${q} en efectivo y cuentas. Margen ajustado.`,
          `Aún hay ${q} sin contar el cupo de la tarjeta. Repasa Saldo.`,
        ]),
        detalle: rotarFrase(`sal-baj-d-${ymd}`, [
          'Saldo o Ingreso si hace falta. Tú eliges.',
          'Antes de ajustar, mira en Saldo dónde más puedes tocar.',
          'Un ingreso o menos gasto hoy: Saldo e Ingresos te ayudan a verlo.',
          'Más → Ingreso o ajusta saldos: que no te agarre con la caja baja.',
          'Revisa billeteras y banco en Saldo; el cupo de la TC es otra cifra.',
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
        `Cupo libre 0 (app) · tope anotado ${limTxt} · afinar en Saldo.`,
        `TC: 0 libre según registro. Revisa límite y usado (Saldo).`,
        `Según la app, no te queda cupo. Tope anotado ${limTxt} · mira en Saldo.`,
        `Cupo 0: cuadra cifra con banco. Límite en app ${limTxt}.`,
        `Sin cupo libre ahora. Revisa tope y deuda anotada (${limTxt}) en Saldo.`,
      ]),
      detalle: rotarFrase(`tcsin-d-${ymd}`, [
        'Saldo → Tarjetas, que coincida con el banco.',
        'Ajusta números en Saldo si no cuadra.',
        'Banco real vs app: alinea límite y saldo usado en Saldo.',
        'Límite, compras y abonos: en Saldo lo dejas claro con el banco.',
        'Si el banco dice otra cifra, corrige en Más → Saldo → Tarjeta.',
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
        `Cupo global ~${pg}% usado. Ojo a Saldo y números.`,
        `Casi lleno el ${pg}% del cupo (global). Revisa.`,
        `Casi lleno: ~${pg}% del cupo total en el registro global.`,
        `Te queda poco colchón: ~${pg}% de cupo ya salió.`,
        `Global: ~${pg}% de cupo usado. Mira banco y app en Saldo.`,
      ]),
      detalle: rotarFrase(`tcglo-d-${ymd}`, [
        'Ajusta límite o ritmo de gasto en Saldo. Tú eliges.',
        'Saldo: bajar carga o el tope que diste en la ficha de TC.',
        'Menos rotación o tope distinto: edítalo en Saldo y sigue al banco.',
        'Abona o ajusta el límite anotado para que el % baje con la realidad.',
        'Cupo fino: pasa por Saldo antes de otra compra.',
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
              `Corte ${t.nombreEntidad}: pago sugerido ~${pagoSug} (cuotas + deuda, int. aprox.).`,
              `Corte: ~${pagoSug} · ${t.nombreEntidad} · revisa Más → Pagos programados.`,
              `Hoy corte en ${t.nombreEntidad}. Sugerido ~${pagoSug} (aprox. según tramos e intereses).`,
              `Corte hoy, ${t.nombreEntidad}. Monto aprox. ~${pagoSug} en Pagos programados o Saldo.`,
              `Es día de corte: ${t.nombreEntidad}. Guía de pago ~${pagoSug}.`,
            ])
          : rotarFrase(`${idEx}-t0-${ymd}`, [
              `Día de corte · abre el extracto (${t.nombreEntidad}).`,
              `Corte hoy: movimientos y extracto · ${t.nombreEntidad}.`,
              `Hoy toca corte con ${t.nombreEntidad}. Mira el extracto en la app.`,
              `Corte en curso. ${t.nombreEntidad}: abre el detalle en Saldo.`,
              `Corte hoy, ${t.nombreEntidad}. Revisa deuda y cuotas anotadas.`,
            ]),
        detalle: pagoSug
          ? rotarFrase(`${idEx}-d-${ymd}`, [
              `Capital cierre: ${cierreTxt || '—'}. Int. est. periodo: ${intTxt || '0'}. Hay recordatorio con monto en Pagos programados (TC). Saldo: extracto de la tarjeta.`,
              'Montos aprox. con tasa E.A. y tramos a cuota en Gastos. Confirma en banco. Saldo: ver detalle.',
              'Tasa y cuotas en Gastos; el cierre banco puede variar. Saldo: extracto. Pagos programados: recordatorio.',
              'Intereses aprox. según tasa. Corrobora corte e intereses reales con el banco.',
              'Si ya bajas el pago, regístralo: Pagos programados o Gastos, como lo tengas anotado.',
            ])
          : rotarFrase(`${idEx}-d0-${ymd}`, [
              'Toca para ver cupo, detalle y proyección a 3/6 cuotas.',
              'Extracto con deuda, cuotas y ahorro estimado. Toca.',
              'En Saldo está el corte, cupo y lo que hace la app con tus compras.',
              'Anota pago o revisa: el extracto hoy es la brújula con el banco.',
              'Abre el extracto en la app: ahí alineas con lo del banco.',
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
              `${t.nombreEntidad} · pago ~${pagoL} en ${dP} d. (vence ${t.etiquetaProxPago || 'Saldo'}).`,
              `Vence: ~${pagoL} · ${t.nombreEntidad} · ${dP} d.`,
              `Pago aprox. ${pagoL} · ${dP} día${dP === 1 ? '' : 's'} · ${t.nombreEntidad}.`,
              `Cuenta: ${t.nombreEntidad} · ~${pagoL} en ${dP} d. (${t.etiquetaProxPago || 'Saldo'}).`,
              `Faltan ${dP} d. · ${t.nombreEntidad} · pago alrededor de ${pagoL}.`,
            ])
          : rotarFrase(`${id}-t0-${ymd}`, [
              `${t.nombreEntidad} · pago en ${dP} día${dP === 1 ? '' : 's'} (${t.etiquetaProxPago || 'Saldo'}).`,
              `Cerca el pago ${t.nombreEntidad} · ${dP} d. · ${t.etiquetaProxPago || 'ver en Saldo'}.`,
              `En ${dP} d. vence pago a ${t.nombreEntidad}. Revisa en Saldo.`,
              `Cuenta regresiva: ${dP} d. para pago a ${t.nombreEntidad}.`,
              `Pago próximo a ${t.nombreEntidad} · plazo ${dP} d.`,
            ]),
        detalle: pagoL
          ? rotarFrase(`${id}-d-${ymd}`, [
              'Incluye tramos a cuota e int. aprox. Pagos programados (TC) y Gastos. Corrobora con el banco.',
              'Que cuadre con el extracto real; evita intereses o mora.',
              'Banco pone la cifra final; en la app aprox. con cuotas e intereses anotados.',
              'Fecha: no pases. Registra pago o abono como lo tengas en el banco.',
              'Monto guía: confírmalo con el banco y anótalo en Pagos o Gastos.',
            ])
          : rotarFrase(`${id}-d0-${ymd}`, [
              'Banco o Saldo, sin pasarte de la fecha.',
              'Anótalo donde lo veas: evita intereses o mora.',
              'Revisa extracto: el calendario no espera; el registro, sí.',
              'Hoy aún a tiempo: entra a Saldo y a la ficha de la tarjeta.',
              'Anota pago a tiempo: la app luego acompaña con lo del banco.',
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
              `Corte en ${t.diasCorte} d. · ${t.nombreEntidad} · sugerido ~${corteSug}.`,
              `En ${t.diasCorte} d. corte, ${t.nombreEntidad}. Guía de pago ~${corteSug}.`,
              `Casi corte: ${t.diasCorte} d. · ${t.nombreEntidad} · aprox. ${corteSug}.`,
              `Faltan ${t.diasCorte} d. al corte · ${t.nombreEntidad} · monto aprox. ${corteSug}.`,
              `Corte a la vuelta: ${t.diasCorte} d., ${t.nombreEntidad}. Revisa ~${corteSug}.`,
            ])
          : rotarFrase(`${id}-t0-${ymd}`, [
              `Corte ${t.nombreEntidad} en unos ${t.diasCorte} d. (${t.etiquetaProxCorte || 'Saldo'}).`,
              `${t.nombreEntidad} · corte cerca, ${t.diasCorte} d.`,
              `Corte en ${t.diasCorte} d. · abre el extracto de ${t.nombreEntidad} en la app.`,
              `Próximo corte: ${t.diasCorte} d. · ${t.nombreEntidad} · Saldo y banco.`,
              `En pocos días corte, ${t.nombreEntidad} · revisa movimientos.`,
            ]),
        detalle: rotarFrase(`${id}-c-${ymd}`, [
          'Revisa compras y tramos a cuota en Gastos. Recordatorio y monto anotados en Pagos programados. Saldo: extracto.',
          'Cierre con cifra clara: mira en Saldo y confirma en banco.',
          'Ajusta compras a cuota en Gastos; el cierre lo cierra con el banco.',
          'Pagos programados ya lleva el corte: si cambia, edita en Más → Pagos programados y Saldo.',
          'Saldo: extracto. Antes de corte, que cuadre cupo, compras y pago aproximado.',
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
            `Uff, en ${t.nombreEntidad} el cupo se acaba: ~${pLibre}% libre (~${lTxt} de ${totTxt})`,
            `Ojo, ${t.nombreEntidad}: te queda como un ~${pLibre}% del cupo libre. Pronto 0 en la app.`,
            `${t.nombreEntidad} · poco aire: ~${pLibre}% del tope aún suelto (${lTxt} / ${totTxt}).`,
            `Casi sin cupo en ${t.nombreEntidad}: ~${pLibre}% libre. Revisa antes de seguir comprando.`,
            `Te queda poco: ~${pLibre}% de cupo con ${t.nombreEntidad} (${lTxt} libres de ${totTxt}).`,
          ]),
          detalle: rotarFrase(`${id20}-d-${ymd}`, [
            'Poco cupo libre: toca parar, abonar o subir el tope en la app. Saldo y banco, alineados.',
            'Baja un pago a cuota, abona a la deuda o no gastes: el límite ceñido en Más → Saldo.',
            'Si en banco ves otro “disponible”, corregimos el tope o el usado en la ficha de la tarjeta.',
            'Cupo al filo: evita sorpresas al corte; pasa por Saldo y confirma con el extracto real.',
            'Ese resto se va en un giro. Planea: abono, menos compra o ajuste de cupo; Saldo: números.',
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
            `${t.nombreEntidad} · ~${pU}% del cupo usado.`,
            `${pU}% de cupo en ${t.nombreEntidad} · ojo a Saldo.`,
            `Casi ${pU}% del cupo en ${t.nombreEntidad} (registro de la app).`,
            `Uso: ~${pU}% con ${t.nombreEntidad}. Mira límite en Saldo.`,
            `Cupo al límite: ${t.nombreEntidad} ~${pU}% usado, mira en Saldo.`,
          ]),
          detalle: rotarFrase(`${idU}-d-${ymd}`, [
            'Ajusta ritmo o límite en Saldo. Tú eliges.',
            'Si ajusta, baja carga o el tope anotado.',
            'Menos compras a cuota o abono: el % baja o sube tope; Saldo alinea con el banco.',
            'Ojo: no te quede sin colchón para sorpresas del cierre. Saldo: tarjeta concreta.',
            'Cupo fino: decide si subes pago, bajas tope o paras compras. Saldo: números.',
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
          `${t.nombreEntidad} · poco libre: ${lTxt} / ${totTxt}.`,
          `Margen bajo en ${t.nombreEntidad} · ~${lTxt} libre.`,
          `Casi llena la TC: ${t.nombreEntidad} · libre ${lTxt} de ${totTxt}.`,
          `Queda poco: ${lTxt} libre en ${t.nombreEntidad} (cupo total ${totTxt}).`,
          `${t.nombreEntidad} · poco aire: ${lTxt} de ${totTxt} aún sueltos.`,
        ]),
        detalle: rotarFrase(`${idL}-d-${ymd}`, [
          'Alinea con el banco en Saldo si no cuadra.',
          'Saldo: que el dato acompañe a la real del extracto.',
          'Antes de comprar, mira banco: el libre a veces tarda en reflejarse en la app.',
          'Poco colchón: pago, abono o no gastar. Saldo: cupo y usado anotado.',
          'Si en banco sobra o falta, corrige tope o saldo en Más → Saldo, tarjeta.',
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
