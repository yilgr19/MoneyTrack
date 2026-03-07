/**
 * Formatea un número con separadores de miles y millones (ej: 940000.50 → "940.000,50")
 * Usa locale español para máxima legibilidad.
 */
function formatearNumero(num, decimales = 2) {
    if (num === null || num === undefined || isNaN(num)) return '0,00';
    const n = parseFloat(num);
    return n.toLocaleString('es', {
        minimumFractionDigits: decimales,
        maximumFractionDigits: decimales
    });
}

/** Cuentas disponibles en el sistema */
const CUENTAS = [
    { id: 'efectivo', nombre: 'Efectivo' },
    { id: 'banco', nombre: 'Banco' },
    { id: 'tarjetaCredito', nombre: 'Tarjeta de crédito' },
    { id: 'nequi', nombre: 'Nequi' },
    { id: 'daviplata', nombre: 'Daviplata' }
];

/** Obtiene los saldos iniciales (compatibilidad con datos antiguos) */
function obtenerSaldosIniciales() {
    const saldosCuentas = localStorage.getItem('saldosCuentas');
    if (saldosCuentas) {
        const parsed = JSON.parse(saldosCuentas);
        return CUENTAS.reduce((acc, c) => {
            acc[c.id] = parseFloat(parsed[c.id]) || 0;
            return acc;
        }, {});
    }
    const legacy = {
        efectivo: parseFloat(localStorage.getItem('saldoEfectivo')) || 0,
        banco: parseFloat(localStorage.getItem('saldoBanco')) || 0
    };
    return CUENTAS.reduce((acc, c) => {
        acc[c.id] = legacy[c.id] !== undefined ? legacy[c.id] : 0;
        return acc;
    }, {});
}

/** Calcula saldos actuales por cuenta (ingresos - gastos - contribuciones) */
function calcularSaldosPorCuenta() {
    const saldosIni = obtenerSaldosIniciales();
    const ingresos = JSON.parse(localStorage.getItem('ingresos') || '[]');
    const gastos = JSON.parse(localStorage.getItem('gastos') || '[]');
    const contribuciones = JSON.parse(localStorage.getItem('contribucionesMetas') || '[]');

    const saldos = {};
    CUENTAS.forEach(c => {
        const ing = ingresos.filter(i => i.origen === c.id).reduce((s, i) => s + i.cantidad, 0);
        const gast = gastos.filter(g => g.origen === c.id || (c.id === 'tarjetaCredito' && g.origen === 'Tarjeta de crédito')).reduce((s, g) => {
            const monto = (c.id === 'tarjetaCredito' && g.cuotas > 1)
                ? (g.cuotaMensual || (g.cantidad || 0) / g.cuotas)
                : (g.cantidad || 0);
            return s + monto;
        }, 0);
        const contrib = contribuciones.filter(x => x.origen === c.id).reduce((s, x) => s + x.cantidad, 0);
        saldos[c.id] = saldosIni[c.id] + ing - gast - contrib;
    });
    saldos.total = Object.values(saldos).reduce((a, b) => a + b, 0);
    saldos.totalReservado = contribuciones.reduce((s, c) => s + c.cantidad, 0);
    return saldos;
}

/** Devuelve el monto a descontar de una cuenta para un gasto. Tarjeta en cuotas: cuota mensual. Resto: total. */
function montoGastoPorCuenta(g, cuentaId) {
    if (cuentaId === 'tarjetaCredito' && g.cuotas > 1) {
        return g.cuotaMensual || (g.cantidad / g.cuotas) || 0;
    }
    return g.cantidad || 0;
}

/** Monto que afecta el saldo/crédito: para tarjeta en cuotas usa cuotaMensual, sino cantidad. */
function montoGastoAfectaSaldo(g) {
    if (!g || g.origen !== 'tarjetaCredito') return g ? (g.cantidad || 0) : 0;
    return (g.cuotas > 1) ? (g.cuotaMensual || (g.cantidad || 0) / g.cuotas) : (g.cantidad || 0);
}

/** Crédito usado en tarjeta: suma de las CUOTAS MENSUALES (no el total). En cuotas, solo se descuenta 1/cuotas cada mes. */
function obtenerGastadoTarjetaCredito() {
    const gastos = JSON.parse(localStorage.getItem('gastos') || '[]');
    return gastos.filter(g => (g.origen === 'tarjetaCredito' || g.origen === 'Tarjeta de crédito')).reduce((s, g) => {
        return s + montoGastoPorCuenta(g, 'tarjetaCredito');
    }, 0);
}

/** Verifica si debe mostrarse alerta de tarjeta al 50% o más. gastado = suma de cuotas mensuales. */
function verificarAlertaTarjetaCredito() {
    const limite = parseFloat(localStorage.getItem('limiteTarjetaCredito')) || 0;
    if (limite <= 0) return { mostrar: false, gastado: 0, limite: 0, porcentaje: 0 };
    const gastado = obtenerGastadoTarjetaCredito();
    const porcentaje = limite > 0 ? (gastado / limite) * 100 : 0;
    return { mostrar: porcentaje >= 50, gastado, limite, porcentaje };
}
