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

/** Normaliza origen a id de cuenta (Banco->banco, Tarjeta de crédito->tarjetaCredito) */
function normalizarOrigenCuenta(origen) {
    if (!origen || typeof origen !== 'string') return '';
    const o = origen.trim();
    const map = { 'efectivo':'efectivo','banco':'banco','tarjetacredito':'tarjetaCredito','nequi':'nequi','daviplata':'daviplata',
        'tarjetadecredito':'tarjetaCredito','tarjetadecrédito':'tarjetaCredito','tarjeta':'tarjetaCredito' };
    const key = o.toLowerCase().replace(/\s/g, '').normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    if (map[key]) return map[key];
    const c = CUENTAS.find(x => x.nombre.toLowerCase() === o.toLowerCase() || x.id === o);
    return c ? c.id : o;
}

/** Calcula saldos actuales por cuenta (ingresos - gastos - contribuciones) */
function calcularSaldosPorCuenta() {
    const saldosIni = obtenerSaldosIniciales();
    const ingresos = JSON.parse(localStorage.getItem('ingresos') || '[]');
    const gastos = JSON.parse(localStorage.getItem('gastos') || '[]');
    const contribuciones = JSON.parse(localStorage.getItem('contribucionesMetas') || '[]');
    const limiteTc = parseFloat(localStorage.getItem('limiteTarjetaCredito')) || 0;

    const saldos = {};
    CUENTAS.forEach(c => {
        const ing = ingresos.filter(i => normalizarOrigenCuenta(i.origen) === c.id).reduce((s, i) => s + i.cantidad, 0);
        const gast = gastos.filter(g => {
            const orig = normalizarOrigenCuenta(g.origen);
            return orig === c.id || (c.id === 'tarjetaCredito' && (orig === 'tarjetaCredito' || g.origen === 'Tarjeta de crédito'));
        }).reduce((s, g) => {
            const monto = (c.id === 'tarjetaCredito' && g.cuotas > 1)
                ? (g.cuotaMensual || (g.cantidad || 0) / g.cuotas)
                : (g.cantidad || 0);
            return s + monto;
        }, 0);
        const contrib = contribuciones.filter(x => normalizarOrigenCuenta(x.origen) === c.id).reduce((s, x) => s + x.cantidad, 0);
        // Tarjeta de crédito: crédito disponible = límite - gastado (si hay límite), sino saldo inicial - gastos
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

/** Devuelve el monto a descontar de una cuenta para un gasto. Tarjeta en cuotas: cuota mensual. Resto: total. */
function montoGastoPorCuenta(g, cuentaId) {
    if (cuentaId === 'tarjetaCredito' && g.cuotas > 1) {
        return g.cuotaMensual || (g.cantidad / g.cuotas) || 0;
    }
    return g.cantidad || 0;
}

/** Monto que afecta el saldo/crédito: para tarjeta en cuotas usa cuotaMensual, sino cantidad. */
function montoGastoAfectaSaldo(g) {
    if (!g) return 0;
    const orig = normalizarOrigenCuenta(g.origen);
    if (orig !== 'tarjetaCredito') return g.cantidad || 0;
    return (g.cuotas > 1) ? (g.cuotaMensual || (g.cantidad || 0) / g.cuotas) : (g.cantidad || 0);
}

/** Crédito usado en tarjeta: suma de las CUOTAS MENSUALES (no el total). En cuotas, solo se descuenta 1/cuotas cada mes. */
function obtenerGastadoTarjetaCredito() {
    const gastos = JSON.parse(localStorage.getItem('gastos') || '[]');
    return gastos.filter(g => normalizarOrigenCuenta(g.origen) === 'tarjetaCredito').reduce((s, g) => {
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

/** Obtiene pagos programados desde localStorage */
function obtenerPagosProgramados() {
    return JSON.parse(localStorage.getItem('pagosProgramados') || '[]');
}

/** Guarda pagos programados en localStorage */
function guardarPagosProgramados(pagos) {
    localStorage.setItem('pagosProgramados', JSON.stringify(pagos));
}

/** Indica si un pago programado vence hoy según su frecuencia y última ejecución */
function pagoVenceHoy(pago, hoy) {
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
        return hoy.getFullYear() === fechaPago.getFullYear() &&
               hoy.getMonth() === fechaPago.getMonth() &&
               hoy.getDate() === fechaPago.getDate();
    }
    return false;
}

/** Indica si un pago programado debe mostrarse para pagar (cuotas únicas: solo cuando ya venció o vence hoy) */
function pagoDebeMostrarseParaPagar(pago, hoy) {
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

/** Ejecuta pagos programados del día: crea gastos automáticamente. Se ejecuta al cargar cualquier página. */
function ejecutarPagosDelDia() {
    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0);
    const pagos = obtenerPagosProgramados();
    const categorias = JSON.parse(localStorage.getItem('categorias') || '[]');
    const categoriasNombres = categorias.map(c => typeof c === 'string' ? c : c.nombre);
    let gastos = JSON.parse(localStorage.getItem('gastos') || '[]');
    let modificado = false;
    const pad = n => String(n).padStart(2, '0');

    pagos.forEach(p => {
        if (!pagoVenceHoy(p, hoy)) return;
        const catValida = categoriasNombres.includes(p.categoria) ? p.categoria : (categoriasNombres[0] || 'Otros');
        const fechaStr = `${hoy.getFullYear()}-${pad(hoy.getMonth() + 1)}-${pad(hoy.getDate())}T12:00:00`;
        gastos.push({
            nombre: p.concepto,
            cantidad: p.monto,
            fecha: fechaStr,
            categoria: catValida,
            origen: p.cuenta,
            nota: 'Pago programado automático',
            cuotas: 1,
            cuotaMensual: p.monto,
            pagoProgramadoId: p.id
        });
        p.ultimaEjecucion = hoy.toISOString().slice(0, 10);
        modificado = true;
    });

    if (modificado) {
        guardarPagosProgramados(pagos);
        localStorage.setItem('gastos', JSON.stringify(gastos));
    }
}

// Ejecución automática desactivada: los pagos se registran desde Gastos con el botón "Pagar"
// ejecutarPagosDelDia();
