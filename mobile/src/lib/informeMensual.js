import {
  montoGastoAfectaSaldoEnMes,
  montoGastoCuentaParaPresupuestoEnMes,
  obtenerMesAño,
  normalizarOrigenCuenta,
  normalizarCategoria,
} from './finance';

const NOMBRES_MES = [
  'Enero',
  'Febrero',
  'Marzo',
  'Abril',
  'Mayo',
  'Junio',
  'Julio',
  'Agosto',
  'Septiembre',
  'Octubre',
  'Noviembre',
  'Diciembre',
];

/** Etiquetas cortas para ejes o leyendas de cuentas. */
const ETIQUETA_CUENTA = {
  efectivo: 'Efectivo',
  banco: 'Banco',
  tarjetaCredito: 'Tarjeta',
  nequi: 'Nequi',
  daviplata: 'Daviplata',
  billeteras: 'Billeteras',
};

export function opcionesMesesInforme(mesesAtras = 36) {
  const out = [];
  const now = new Date();
  for (let i = 0; i < mesesAtras; i += 1) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    const y = d.getFullYear();
    const m = d.getMonth() + 1;
    const value = `${y}-${String(m).padStart(2, '0')}`;
    const label = d.toLocaleDateString('es', { month: 'long', year: 'numeric' });
    out.push({ value, label: label.charAt(0).toUpperCase() + label.slice(1) });
  }
  return out;
}

export function etiquetaMesDesdeYM(ym) {
  const [a, b] = String(ym).split('-');
  const y = parseInt(a, 10);
  const m0 = parseInt(b, 10) - 1;
  if (!Number.isFinite(y) || m0 < 0 || m0 > 11) return String(ym);
  return `${NOMBRES_MES[m0]} ${y}`;
}

/**
 * @param {object} state
 * @param {string} yearMonth YYYY-MM
 * @returns {object|null}
 */
export function construirDatosInformeMensual(state, yearMonth) {
  const [ys, ms] = String(yearMonth).split('-');
  const año = parseInt(ys, 10);
  const mes = parseInt(ms, 10) - 1;
  if (!Number.isFinite(año) || mes < 0 || mes > 11) return null;

  const ingresos = state.ingresos || [];
  const gastos = state.gastos || [];
  const contribuciones = state.contribucionesMetas || [];
  const metas = state.metas || [];
  const categoriasCfg = (state.categorias || []).map(normalizarCategoria);

  let totalIngresos = 0;
  let numIngresos = 0;
  ingresos.forEach((i) => {
    if (i.esRetiroBolsillo) return;
    const { mes: m, año: a } = obtenerMesAño(i.fecha);
    if (m !== mes || a !== año) return;
    totalIngresos += Math.abs(parseFloat(i.cantidad) || 0);
    numIngresos += 1;
  });

  let totalGastosPres = 0;
  let totalGastosSaldo = 0;
  let numGastosMes = 0;
  const porCategoria = {};
  const porCuenta = {};

  gastos.forEach((g) => {
    const p = montoGastoCuentaParaPresupuestoEnMes(g, state, mes, año);
    const s = montoGastoAfectaSaldoEnMes(g, state, mes, año);
    if (p > 0) numGastosMes += 1;
    totalGastosPres += p;
    totalGastosSaldo += s;
    if (p > 0) {
      const cat = g.categoria || 'Otros';
      porCategoria[cat] = (porCategoria[cat] || 0) + p;
    }
    if (s > 0) {
      const orig = normalizarOrigenCuenta(g.origen);
      porCuenta[orig] = (porCuenta[orig] || 0) + s;
    }
  });

  const tope = parseFloat(state.presupuestoMensual) || 0;
  const disponiblePres = tope > 0 ? tope - totalGastosPres : null;

  const topGastos = gastos
    .map((g) => ({
      g,
      m: montoGastoAfectaSaldoEnMes(g, state, mes, año),
    }))
    .filter((x) => x.m > 0)
    .sort((a, b) => b.m - a.m)
    .slice(0, 8);

  let totalAportesMetas = 0;
  const aportesLines = [];
  contribuciones.forEach((c) => {
    const { mes: m, año: a } = obtenerMesAño(c.fecha);
    if (m !== mes || a !== año) return;
    const cant = Math.abs(parseFloat(c.cantidad) || 0);
    totalAportesMetas += cant;
    const nom = metas.find((x) => x.id === c.metaId)?.nombre || 'Meta';
    aportesLines.push({ nombre: nom, cant });
  });

  const cuentaRows = Object.entries(porCuenta)
    .map(([id, monto]) => ({
      id,
      label: ETIQUETA_CUENTA[id] || id,
      monto,
    }))
    .filter((r) => r.monto > 0.001)
    .sort((a, b) => b.monto - a.monto);

  const categoriaRows = Object.entries(porCategoria)
    .map(([nombre, monto]) => {
      const c = categoriasCfg.find((x) => x.nombre === nombre);
      return {
        nombre,
        monto,
        color: c?.color,
        icono: c?.icono,
      };
    })
    .filter((r) => r.monto > 0.001)
    .sort((a, b) => b.monto - a.monto);

  const resultado = totalIngresos - totalGastosPres;

  return {
    yearMonth,
    año,
    mes,
    labelMes: etiquetaMesDesdeYM(yearMonth),
    totalIngresos,
    totalGastosPres,
    totalGastosSaldo,
    resultado,
    tope,
    disponiblePres,
    porCategoria,
    categoriaRows,
    cuentaRows,
    topGastos,
    totalAportesMetas,
    aportesLines,
    numIngresos,
    numGastosMes,
    categoriasCfg,
  };
}
