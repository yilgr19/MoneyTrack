/** Misma lógica que js/utils.js, usando un objeto `data` en memoria (sustituye localStorage). */

export function formatearNumero(num, decimales = 2) {
  if (num === null || num === undefined || Number.isNaN(Number(num))) return '0,00';
  const n = parseFloat(num);
  return n.toLocaleString('es', {
    minimumFractionDigits: decimales,
    maximumFractionDigits: decimales,
  });
}

export const CUENTAS = [
  { id: 'efectivo', nombre: 'Efectivo' },
  { id: 'banco', nombre: 'Banco' },
  { id: 'tarjetaCredito', nombre: 'Tarjeta de crédito' },
  { id: 'nequi', nombre: 'Nequi' },
  { id: 'daviplata', nombre: 'Daviplata' },
];

export function obtenerSaldosIniciales(data) {
  const raw = data.saldosCuentas;
  if (raw && typeof raw === 'object' && !Array.isArray(raw)) {
    return CUENTAS.reduce((acc, c) => {
      acc[c.id] = parseFloat(raw[c.id]) || 0;
      return acc;
    }, {});
  }
  const legacy = {
    efectivo: parseFloat(data.saldoEfectivo) || 0,
    banco: parseFloat(data.saldoBanco) || 0,
  };
  return CUENTAS.reduce((acc, c) => {
    acc[c.id] = legacy[c.id] !== undefined ? legacy[c.id] : 0;
    return acc;
  }, {});
}

export function normalizarOrigenCuenta(origen) {
  if (!origen || typeof origen !== 'string') return '';
  const o = origen.trim();
  const map = {
    efectivo: 'efectivo',
    banco: 'banco',
    tarjetacredito: 'tarjetaCredito',
    nequi: 'nequi',
    daviplata: 'daviplata',
    tarjetadecredito: 'tarjetaCredito',
    tarjetadecrédito: 'tarjetaCredito',
    tarjeta: 'tarjetaCredito',
  };
  const key = o
    .toLowerCase()
    .replace(/\s/g, '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
  if (map[key]) return map[key];
  const c = CUENTAS.find((x) => x.nombre.toLowerCase() === o.toLowerCase() || x.id === o);
  return c ? c.id : o;
}

export function calcularSaldosPorCuenta(data) {
  const saldosIni = obtenerSaldosIniciales(data);
  const ingresos = data.ingresos || [];
  const gastos = data.gastos || [];
  const contribuciones = data.contribucionesMetas || [];
  const limiteTc = parseFloat(data.limiteTarjetaCredito) || 0;

  const saldos = {};
  CUENTAS.forEach((c) => {
    const ing = ingresos
      .filter((i) => normalizarOrigenCuenta(i.origen) === c.id)
      .reduce((s, i) => s + i.cantidad, 0);
    const gast = gastos
      .filter((g) => {
        const orig = normalizarOrigenCuenta(g.origen);
        return orig === c.id || (c.id === 'tarjetaCredito' && (orig === 'tarjetaCredito' || g.origen === 'Tarjeta de crédito'));
      })
      .reduce((s, g) => {
        const monto =
          c.id === 'tarjetaCredito' && g.cuotas > 1
            ? g.cuotaMensual || (g.cantidad || 0) / g.cuotas
            : g.cantidad || 0;
        return s + monto;
      }, 0);
    const contrib = contribuciones
      .filter((x) => normalizarOrigenCuenta(x.origen) === c.id)
      .reduce((s, x) => s + x.cantidad, 0);
    if (c.id === 'tarjetaCredito' && limiteTc > 0) {
      saldos[c.id] = Math.max(0, limiteTc - gast - contrib);
    } else {
      saldos[c.id] = saldosIni[c.id] + ing - gast - contrib;
    }
  });
  saldos.total = Object.values(saldos).reduce((a, b) => a + b, 0);
  saldos.totalReservado = contribuciones.reduce((s, c) => s + c.cantidad, 0);
  return saldos;
}

export function montoGastoPorCuenta(g, cuentaId) {
  if (cuentaId === 'tarjetaCredito' && g.cuotas > 1) {
    return g.cuotaMensual || g.cantidad / g.cuotas || 0;
  }
  return g.cantidad || 0;
}

export function montoGastoAfectaSaldo(g) {
  if (!g) return 0;
  const orig = normalizarOrigenCuenta(g.origen);
  if (orig !== 'tarjetaCredito') return g.cantidad || 0;
  return g.cuotas > 1 ? g.cuotaMensual || (g.cantidad || 0) / g.cuotas : g.cantidad || 0;
}

export function obtenerGastadoTarjetaCredito(data) {
  const gastos = data.gastos || [];
  return gastos
    .filter((g) => normalizarOrigenCuenta(g.origen) === 'tarjetaCredito')
    .reduce((s, g) => s + montoGastoPorCuenta(g, 'tarjetaCredito'), 0);
}

export function verificarAlertaTarjetaCredito(data) {
  const limite = parseFloat(data.limiteTarjetaCredito) || 0;
  if (limite <= 0) return { mostrar: false, gastado: 0, limite: 0, porcentaje: 0 };
  const gastado = obtenerGastadoTarjetaCredito(data);
  const porcentaje = limite > 0 ? (gastado / limite) * 100 : 0;
  return { mostrar: porcentaje >= 50, gastado, limite, porcentaje };
}

export function obtenerMesAño(fechaStr) {
  const d = new Date(fechaStr && fechaStr.includes('T') ? fechaStr : `${fechaStr || ''}T12:00:00`);
  return { mes: d.getMonth(), año: d.getFullYear() };
}

export function normalizarCategoria(cat) {
  if (typeof cat === 'string') return { nombre: cat, color: '#6b7280', limite: null, icono: '📋' };
  return {
    nombre: cat.nombre || cat,
    color: cat.color || '#6b7280',
    limite: cat.limite || null,
    icono: cat.icono || '📋',
  };
}

export function generarIdMeta() {
  return `meta_${Date.now()}_${Math.random().toString(36).slice(2, 9)}`;
}

export function generarIdPagoProgramado() {
  return `pago_${Date.now()}_${Math.random().toString(36).slice(2, 9)}`;
}

export function pagoVenceHoy(pago, hoy) {
  if (!pago.activo) return false;
  if (!pago.fechaInicio) return false;
  const [yIni, mIni, dIni] = (pago.fechaInicio + '').slice(0, 10).split('-').map(Number);
  const añoIni = yIni || 0;
  const mesIni = (mIni || 1) - 1;
  const diaIni = dIni || 1;
  if (hoy.getFullYear() < añoIni) return false;
  if (hoy.getFullYear() === añoIni && hoy.getMonth() < mesIni) return false;
  if (hoy.getFullYear() === añoIni && hoy.getMonth() === mesIni && hoy.getDate() < diaIni) return false;

  const ultima = pago.ultimaEjecucion ? new Date(pago.ultimaEjecucion + 'T12:00:00') : null;
  const diaHoy = hoy.getDate();
  const mesHoy = hoy.getMonth();
  const añoHoy = hoy.getFullYear();

  if (pago.frecuencia === 'mensual') {
    const diaPago = Math.min(28, parseInt(pago.diaPago, 10) || 1);
    if (diaHoy !== diaPago) return false;
    if (ultima && ultima.getFullYear() === añoHoy && ultima.getMonth() === mesHoy) return false;
    return true;
  }
  if (pago.frecuencia === 'quincenal') {
    const diaPago = parseInt(pago.diaPago, 10);
    const diasValidos = [1, 15];
    if (!diasValidos.includes(diaPago)) return false;
    if (diaHoy !== diaPago) return false;
    if (ultima) {
      const diff = (hoy - ultima) / (1000 * 60 * 60 * 24);
      if (diff < 14) return false;
    }
    return true;
  }
  if (pago.frecuencia === 'semanal') {
    const fechaInicio = new Date(pago.fechaInicio + 'T12:00:00');
    const diaSemanaInicio = fechaInicio.getDay();
    if (hoy.getDay() !== diaSemanaInicio) return false;
    if (ultima) {
      const diff = (hoy - ultima) / (1000 * 60 * 60 * 24);
      if (diff < 6) return false;
    }
    return true;
  }
  if (pago.frecuencia === 'unico') {
    const fechaPago = new Date(pago.fechaInicio + 'T12:00:00');
    return (
      hoy.getFullYear() === fechaPago.getFullYear() &&
      hoy.getMonth() === fechaPago.getMonth() &&
      hoy.getDate() === fechaPago.getDate()
    );
  }
  return false;
}

export function pagoDebeMostrarseParaPagar(pago, hoy) {
  if (!pago || pago.activo === false) return false;
  if (pago.frecuencia === 'unico') {
    if (!pago.fechaInicio) return false;
    const fechaPago = new Date(pago.fechaInicio + 'T12:00:00');
    fechaPago.setHours(0, 0, 0, 0);
    const hoyNorm = new Date(hoy);
    hoyNorm.setHours(0, 0, 0, 0);
    return fechaPago <= hoyNorm;
  }
  return true;
}
