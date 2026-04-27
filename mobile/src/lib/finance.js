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
  /** Suma de billeteras y apps que no son Nequi ni Daviplata (Movii, Línea, etc.) */
  { id: 'billeteras', nombre: 'Otras plataformas' },
];

/** Plataformas digitales en Colombia (saldo inicial «Mis plataformas»). */
export const PLATAFORMA_OTRO_VALUE = '__otro__';

export const PLATAFORMAS_CO = [
  { value: 'nequi', label: 'Nequi' },
  { value: 'daviplata', label: 'Daviplata' },
  { value: 'movii', label: 'Movii' },
  { value: 'linea', label: 'Línea (Línea Pay)' },
  { value: 'dale', label: 'Dale!' },
  { value: 'rappipay', label: 'RappiPay' },
  { value: 'bold', label: 'Bold' },
  { value: 'tpaga', label: 'Tpaga' },
  { value: 'powwi', label: 'Powwi' },
  { value: 'claro_pay', label: 'Claro Pay' },
  { value: 'iris', label: 'Iris' },
  { value: 'tyba', label: 'Tyba' },
  { value: 'ziper', label: 'Zíper' },
  { value: 'a2censo', label: 'A2censo' },
  { value: 'movistar_wallet', label: 'Movistar Wallet' },
  { value: 'wompi', label: 'Wompi' },
  { value: 'bre_b', label: 'Bre-B' },
  { value: 'livibank', label: 'Livi / Uala' },
];

export function getPlataformasOptions() {
  return [
    ...PLATAFORMAS_CO,
    { value: PLATAFORMA_OTRO_VALUE, label: 'Otro (especificar)' },
  ];
}

/** Clave en saldosCuentas donde acumula cada opción del selector. */
export function cuentaSaldoPlataforma(platformValue) {
  if (platformValue === 'nequi') return 'nequi';
  if (platformValue === 'daviplata') return 'daviplata';
  return 'billeteras';
}

const PLATAFORMA_VALUE_TO_LABEL = (() => {
  const m = new Map();
  PLATAFORMAS_CO.forEach((p) => m.set(p.value, p.label));
  return m;
})();

export function getPlataformaLabelByValue(value) {
  if (!value || value === PLATAFORMA_OTRO_VALUE) return '';
  return PLATAFORMA_VALUE_TO_LABEL.get(value) || '';
}

export function totalPlataformasTresSaldos(saldosObj) {
  const n = parseFloat(saldosObj?.nequi) || 0;
  const d = parseFloat(saldosObj?.daviplata) || 0;
  const b = parseFloat(saldosObj?.billeteras) || 0;
  return n + d + b;
}

/** Valor interno del ítem «Otro banco» en el selector (no incluir en listas por país). */
export const BANCO_OTRO_VALUE = '__otro__';

/**
 * Bancos por moneda ISO (COP → Colombia, MXN → México, etc.).
 * Si la moneda no está mapeada o está vacía, se usa la lista DEFAULT (Latinoamérica general).
 */
const BANCOS_POR_MONEDA = {
  COP: [
    { value: 'cop_bancolombia', label: 'Bancolombia' },
    { value: 'cop_davivienda', label: 'Davivienda' },
    { value: 'cop_bbva', label: 'BBVA Colombia' },
    { value: 'cop_bogota', label: 'Banco de Bogotá' },
    { value: 'cop_scotiabank', label: 'Scotiabank Colpatria' },
    { value: 'cop_popular', label: 'Banco Popular' },
    { value: 'cop_occidente', label: 'Banco de Occidente' },
    { value: 'cop_caja_social', label: 'Banco Caja Social' },
    { value: 'cop_agrario', label: 'Banco Agrario' },
    { value: 'cop_av_villas', label: 'Banco AV Villas' },
    { value: 'cop_finandina', label: 'Banco Finandina' },
    { value: 'cop_itau', label: 'Itaú Colombia' },
  ],
  MXN: [
    { value: 'mxn_bbva', label: 'BBVA México' },
    { value: 'mxn_santander', label: 'Santander México' },
    { value: 'mxn_banamex', label: 'Citibanamex' },
    { value: 'mxn_banorte', label: 'Banorte' },
    { value: 'mxn_hsbc', label: 'HSBC México' },
    { value: 'mxn_scotiabank', label: 'Scotiabank México' },
    { value: 'mxn_inbursa', label: 'Inbursa' },
    { value: 'mxn_banregio', label: 'BanRegio' },
    { value: 'mxn_nu', label: 'Nu México' },
    { value: 'mxn_hey', label: 'Hey Banco' },
    { value: 'mxn_afirme', label: 'Afirme' },
    { value: 'mxn_bajio', label: 'Banco del Bajío' },
  ],
  ARS: [
    { value: 'ars_galicia', label: 'Banco Galicia' },
    { value: 'ars_santander', label: 'Santander Río' },
    { value: 'ars_bbva', label: 'BBVA Argentina' },
    { value: 'ars_macro', label: 'Banco Macro' },
    { value: 'ars_icbc', label: 'ICBC Argentina' },
    { value: 'ars_credicoop', label: 'Banco Credicoop' },
    { value: 'ars_brubank', label: 'Brubank' },
    { value: 'ars_naranja', label: 'Naranja X' },
    { value: 'ars_mercadopago', label: 'Mercado Pago' },
    { value: 'ars_uala', label: 'Ualá' },
    { value: 'ars_citi', label: 'Citibank Argentina' },
  ],
  CLP: [
    { value: 'clp_chile', label: 'Banco de Chile' },
    { value: 'clp_bci', label: 'BCI' },
    { value: 'clp_santander', label: 'Santander Chile' },
    { value: 'clp_scotiabank', label: 'Scotiabank Chile' },
    { value: 'clp_itau', label: 'Itaú Chile' },
    { value: 'clp_estado', label: 'Banco Estado' },
    { value: 'clp_security', label: 'Banco Security' },
    { value: 'clp_consorcio', label: 'Banco Consorcio' },
    { value: 'clp_bice', label: 'Banco BICE' },
  ],
  PEN: [
    { value: 'pen_bcp', label: 'BCP' },
    { value: 'pen_bbva', label: 'BBVA Perú' },
    { value: 'pen_scotiabank', label: 'Scotiabank Perú' },
    { value: 'pen_interbank', label: 'Interbank' },
    { value: 'pen_banbif', label: 'BanBif' },
    { value: 'pen_pichincha', label: 'Banco Pichincha' },
    { value: 'pen_gnb', label: 'GNB Perú' },
    { value: 'pen_credito', label: 'Banco de Crédito del Perú' },
  ],
  USD: [
    { value: 'usd_chase', label: 'Chase' },
    { value: 'usd_bofa', label: 'Bank of America' },
    { value: 'usd_wells', label: 'Wells Fargo' },
    { value: 'usd_citi', label: 'Citibank' },
    { value: 'usd_capital_one', label: 'Capital One' },
    { value: 'usd_td', label: 'TD Bank' },
    { value: 'usd_usbank', label: 'U.S. Bank' },
    { value: 'usd_pnc', label: 'PNC Bank' },
    { value: 'usd_truist', label: 'Truist' },
    { value: 'usd_hsbc_us', label: 'HSBC USA' },
  ],
  EUR: [
    { value: 'eur_santander', label: 'Santander' },
    { value: 'eur_bbva', label: 'BBVA' },
    { value: 'eur_caixa', label: 'CaixaBank' },
    { value: 'eur_ing', label: 'ING' },
    { value: 'eur_deutsche', label: 'Deutsche Bank' },
    { value: 'eur_bnpp', label: 'BNP Paribas' },
    { value: 'eur_unicredit', label: 'UniCredit' },
    { value: 'eur_intesa', label: 'Intesa Sanpaolo' },
  ],
  BRL: [
    { value: 'brl_nubank', label: 'Nubank' },
    { value: 'brl_itau', label: 'Itaú Unibanco' },
    { value: 'brl_bradesco', label: 'Bradesco' },
    { value: 'brl_bb', label: 'Banco do Brasil' },
    { value: 'brl_santander', label: 'Santander Brasil' },
    { value: 'brl_c6', label: 'C6 Bank' },
    { value: 'brl_inter', label: 'Inter' },
    { value: 'brl_btg', label: 'BTG Pactual' },
    { value: 'brl_caixa', label: 'Caixa Econômica Federal' },
    { value: 'brl_safra', label: 'Banco Safra' },
  ],
  GTQ: [
    { value: 'gtq_gyt', label: 'Banco G&T Continental' },
    { value: 'gtq_industrial', label: 'Banco Industrial' },
    { value: 'gtq_banrural', label: 'Banrural' },
    { value: 'gtq_bac', label: 'BAC Credomatic' },
    { value: 'gtq_agromercantil', label: 'Banco Agromercantil' },
    { value: 'gtq_promerica', label: 'Promerica' },
  ],
  CAD: [
    { value: 'cad_rbc', label: 'RBC' },
    { value: 'cad_td', label: 'TD Canada Trust' },
    { value: 'cad_scotia', label: 'Scotiabank Canadá' },
    { value: 'cad_bmo', label: 'BMO' },
    { value: 'cad_cibc', label: 'CIBC' },
    { value: 'cad_national', label: 'National Bank' },
  ],
  GBP: [
    { value: 'gbp_barclays', label: 'Barclays' },
    { value: 'gbp_hsbc', label: 'HSBC UK' },
    { value: 'gbp_lloyds', label: 'Lloyds Bank' },
    { value: 'gbp_natwest', label: 'NatWest' },
    { value: 'gbp_santander', label: 'Santander UK' },
  ],
  JPY: [
    { value: 'jpy_mufg', label: 'MUFG' },
    { value: 'jpy_smbc', label: 'SMBC (Mitsui)' },
    { value: 'jpy_mizuho', label: 'Mizuho' },
    { value: 'jpy_resona', label: 'Resona Bank' },
    { value: 'jpy_japanpost', label: 'Japan Post Bank' },
  ],
  /** Moneda no elegida o sin lista dedicada */
  DEFAULT: [
    { value: 'def_bancolombia', label: 'Bancolombia' },
    { value: 'def_bbva', label: 'BBVA' },
    { value: 'def_santander', label: 'Santander' },
    { value: 'def_scotiabank', label: 'Scotiabank' },
    { value: 'def_nu', label: 'Nu / Nubank' },
    { value: 'def_bcp', label: 'BCP (Perú)' },
    { value: 'def_banorte', label: 'Banorte' },
    { value: 'def_bac', label: 'BAC Credomatic' },
  ],
};

/**
 * Opciones del Picker para la moneda actual (siempre termina en «Otro»).
 */
export function getBancosOptionsForMoneda(moneda) {
  const code = moneda && String(moneda).toUpperCase().trim();
  const list =
    code && BANCOS_POR_MONEDA[code] && BANCOS_POR_MONEDA[code].length
      ? BANCOS_POR_MONEDA[code]
      : BANCOS_POR_MONEDA.DEFAULT;
  return [...list, { value: BANCO_OTRO_VALUE, label: 'Otro (especificar)' }];
}

/** Etiqueta legible para un value de banco (cualquier moneda), útil al cambiar de moneda con filas antiguas. */
const VALUE_TO_LABEL_BANCO = (() => {
  const m = new Map();
  Object.entries(BANCOS_POR_MONEDA).forEach(([k, arr]) => {
    if (!Array.isArray(arr)) return;
    arr.forEach((b) => m.set(b.value, b.label));
  });
  return m;
})();

export function getBankLabelByValue(value) {
  if (!value || value === BANCO_OTRO_VALUE) return '';
  return VALUE_TO_LABEL_BANCO.get(value) || '';
}

export function totalSaldoBancosDetalle(lineas) {
  if (!Array.isArray(lineas)) return 0;
  return lineas.reduce((s, r) => s + (parseFloat(r.saldo) || 0), 0);
}

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
    billeteras: 'billeteras',
    otrasplataformas: 'billeteras',
    otrasbilleteras: 'billeteras',
    movii: 'billeteras',
    linea: 'billeteras',
    lineapay: 'billeteras',
    dale: 'billeteras',
    rappipay: 'billeteras',
    bold: 'billeteras',
    tpaga: 'billeteras',
    powwi: 'billeteras',
    claropay: 'billeteras',
    iris: 'billeteras',
    tyba: 'billeteras',
    ziper: 'billeteras',
    a2censo: 'billeteras',
    wompi: 'billeteras',
    breb: 'billeteras',
    livibank: 'billeteras',
    uala: 'billeteras',
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

/** Valor guardado en ingresos/aportes cuando el destino es una fila de «Cuentas en banco». */
export const PREFIJO_ORIGEN_BANCO = 'banco:';
/** Valor guardado cuando el destino es una fila de «Mis plataformas». */
export const PREFIJO_ORIGEN_PLATAFORMA = 'plataforma:';

/**
 * Bucket de cuenta agregada (efectivo, banco, nequi…) para movimientos.
 * Reconoce orígenes detallados `banco:id` y `plataforma:id`.
 */
export function cuentaBucketMovimiento(origen, data) {
  if (!origen || typeof origen !== 'string') return '';
  const o = origen.trim();
  if (o.startsWith(PREFIJO_ORIGEN_BANCO)) return 'banco';
  if (o.startsWith(PREFIJO_ORIGEN_PLATAFORMA)) {
    const id = o.slice(PREFIJO_ORIGEN_PLATAFORMA.length);
    const row = (data?.plataformasDetalle || []).find((r) => r && r.id === id);
    return row ? cuentaSaldoPlataforma(row.platformValue || PLATAFORMA_OTRO_VALUE) : 'billeteras';
  }
  return normalizarOrigenCuenta(origen);
}

/**
 * Suma de cupos totales. Con filas en Saldo, si la suma de cupo por entidad queda 0, se usa aún
 * `limiteTarjetaCredito` (total guardado al confirmar) para no perder el tope y que Inicio muestre bien el saldo.
 */
export function limiteTotalTarjetasCredito(data) {
  const arr = data.tarjetasCredito;
  const legacy = parseFloat(data.limiteTarjetaCredito) || 0;
  if (Array.isArray(arr) && arr.length > 0) {
    const sumCupo = arr.reduce((s, t) => s + (parseFloat(t.cupoTotal) || 0), 0);
    if (sumCupo > 0) return sumCupo;
    if (legacy > 0) return legacy;
  }
  return legacy;
}

/** Suma de «Cupo utilizado (deuda)» por tarjeta en Saldo; afecta el cupo libre para pagar. */
export function totalCupoUtilizadoTarjetasCredito(data) {
  const arr = data.tarjetasCredito;
  if (!Array.isArray(arr) || arr.length === 0) return 0;
  return arr.reduce((s, t) => s + (parseFloat(t.cupoUtilizado) || 0), 0);
}

/**
 * Fila «Por cuenta» en Inicio: mostrar si hay saldo, ingresos a esa caja, o se dejó
 * fijado saldo inicial o detalle (bancos, plataformas, tarjeta).
 */
export function cuentaVisibleEnResumenInicio(cuentaId, data, saldosPorCuenta) {
  const nN = (v) => {
    const x = Number(v);
    return Number.isFinite(x) ? x : 0;
  };
  const saldo = nN(saldosPorCuenta?.[cuentaId]);
  if (Math.abs(saldo) > 1e-9) return true;

  const ingresos = data.ingresos || [];
  if (
    ingresos.some(
      (i) => cuentaBucketMovimiento(i?.origen, data) === cuentaId && nN(i?.cantidad) !== 0
    )
  ) {
    return true;
  }

  const s = data.saldosCuentas || {};
  if (cuentaId === 'efectivo') return nN(s.efectivo) !== 0;
  if (cuentaId === 'banco') {
    if ((data.bancosDetalle || []).length > 0) return true;
    return nN(s.banco) !== 0;
  }
  if (cuentaId === 'tarjetaCredito') {
    if (limiteTotalTarjetasCredito(data) > 0) return true;
    if ((data.tarjetasCredito || []).length > 0) return true;
    return nN(s.tarjetaCredito) !== 0;
  }
  if (cuentaId === 'nequi' || cuentaId === 'daviplata' || cuentaId === 'billeteras') {
    if (
      (data.plataformasDetalle || []).some(
        (r) => cuentaSaldoPlataforma(r?.platformValue || PLATAFORMA_OTRO_VALUE) === cuentaId
      )
    ) {
      return true;
    }
    return nN(s[cuentaId]) !== 0;
  }
  return true;
}

export function generarIdTarjetaCredito() {
  return `tc-${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
}

function padTiempo(n) {
  return String(n).padStart(2, '0');
}

/** ISO local sin zona: YYYY-MM-DDTHH:mm:ss (interpretación local al parsear). */
export function fechaHoraALocalISO(d) {
  if (!d || Number.isNaN(d.getTime())) return '';
  return `${d.getFullYear()}-${padTiempo(d.getMonth() + 1)}-${padTiempo(d.getDate())}T${padTiempo(d.getHours())}:${padTiempo(d.getMinutes())}:00`;
}

/** Solo calendario local: YYYY-MM-DD */
export function fechaALocalISO(d) {
  if (!d || Number.isNaN(d.getTime())) return '';
  return `${d.getFullYear()}-${padTiempo(d.getMonth() + 1)}-${padTiempo(d.getDate())}`;
}

export function parseFechaHoraLocal(str) {
  if (str == null || String(str).trim() === '') return null;
  const s = String(str).trim();
  const d = new Date(s.includes('T') ? s : `${s.slice(0, 10)}T12:00:00`);
  return Number.isNaN(d.getTime()) ? null : d;
}

/**
 * Patrón mensual: día del mes (la hora se fija al mediodía solo para el cálculo; en UI solo se pide fecha).
 */
function patronMensualDesdeTarjeta(t, kind) {
  const key = kind === 'corte' ? 'fechaHoraCorte' : 'fechaHoraLimitePago';
  const parsed = parseFechaHoraLocal(t[key]);
  if (parsed) {
    return new Date(2020, 0, parsed.getDate(), 12, 0, 0);
  }
  const diaLegacy = Math.min(
    28,
    Math.max(1, parseInt(kind === 'corte' ? t.diaCorte : t.diaLimitePago, 10) || (kind === 'corte' ? 15 : 5))
  );
  return new Date(2020, 0, diaLegacy, 12, 0, 0);
}

/** Próxima ocurrencia estrictamente después de `ref`. */
export function proximaOcurrenciaMensual(anchorPatron, ref = new Date()) {
  if (!anchorPatron || Number.isNaN(anchorPatron.getTime())) return null;
  const dom = anchorPatron.getDate();
  const H = anchorPatron.getHours();
  const M = anchorPatron.getMinutes();
  const S = anchorPatron.getSeconds();
  const MS = anchorPatron.getMilliseconds();
  const base = ref.getFullYear() * 12 + ref.getMonth();
  for (let k = 0; k < 48; k++) {
    const total = base + k;
    const y = Math.floor(total / 12);
    const m = total - y * 12;
    const lastDay = new Date(y, m + 1, 0).getDate();
    const day = Math.min(dom, lastDay);
    const cand = new Date(y, m, day, H, M, S, MS);
    if (cand.getTime() > ref.getTime()) return cand;
  }
  return null;
}

/** Fecha/hora de vencimiento del ciclo del mes de `ref` (puede ser pasada respecto a `ref`). */
export function instanteVencimientoCicloActual(anchorPatron, ref = new Date()) {
  if (!anchorPatron || Number.isNaN(anchorPatron.getTime())) return null;
  const dom = anchorPatron.getDate();
  const H = anchorPatron.getHours();
  const M = anchorPatron.getMinutes();
  const S = anchorPatron.getSeconds();
  const MS = anchorPatron.getMilliseconds();
  const y = ref.getFullYear();
  const m = ref.getMonth();
  const lastDay = new Date(y, m + 1, 0).getDate();
  const day = Math.min(dom, lastDay);
  return new Date(y, m, day, H, M, S, MS);
}

export function diasCalendarioHasta(futuro, ahora = new Date()) {
  if (!futuro || Number.isNaN(futuro.getTime())) return null;
  const ua = Date.UTC(ahora.getFullYear(), ahora.getMonth(), ahora.getDate());
  const ub = Date.UTC(futuro.getFullYear(), futuro.getMonth(), futuro.getDate());
  return Math.round((ub - ua) / 86400000);
}

/** Días hasta la próxima fecha de calendario con ese día del mes (1–28, legacy). */
export function diasHastaProximoDiaCalendario(diaDelMes, ref = new Date()) {
  const dObj = Math.min(28, Math.max(1, parseInt(diaDelMes, 10) || 1));
  const pat = new Date(2020, 0, dObj, 12, 0, 0);
  const prox = proximaOcurrenciaMensual(pat, ref);
  if (!prox) return 0;
  const d = diasCalendarioHasta(prox, ref);
  return d == null ? 0 : d;
}

/**
 * Sustituye recordatorios generados desde tarjetas; conserva el resto de pagos programados.
 * `data` = estado de la app (gastos, saldos, tarjetas, etc.); el monto usa extracto, cuotas e interés E.A.
 * Si hay pagos esCuotaDiferida (Gastos, compra TC a cuotas), no se generan Corte/Límite agregados (evita duplicar el cierre).
 * @param {object} data - Estado de la app (gastos, saldos, tarjetas, categorías, …).
 * @param {Date} [ref] - Hoy: alinear extracto con corte y movimientos.
 */
export function reemplazarPagosRecordatorioTarjetas(pagosExistentes, data, ref = new Date()) {
  const conCuotasDesdeGastos = (pagosExistentes || []).some(
    (p) => p && p.activo !== false && p.esCuotaDiferida === true
  );
  const filtrados = (pagosExistentes || []).filter((p) => !p.esRecordatorioTarjeta);
  if (!data || typeof data !== 'object' || !Array.isArray(data.tarjetasCredito) || data.tarjetasCredito.length === 0) {
    return filtrados;
  }
  if (conCuotasDesdeGastos) {
    return filtrados;
  }
  const categorias = data.categorias;
  const firstCat =
    Array.isArray(categorias) && categorias.length > 0
      ? typeof categorias[0] === 'string'
        ? categorias[0]
        : categorias[0].nombre || 'Otros'
      : 'Otros';
  const extras = [];
  for (const t of data.tarjetasCredito) {
    const nombre = String(t.nombreEntidad || '').trim();
    if (!nombre) continue;
    const fc = String(t.fechaHoraCorte || '').trim();
    const fl = String(t.fechaHoraLimitePago || '').trim();
    const ex = construirExtractoBancarioTarjeta(t, data, ref);
    const montoLimitePago = montoPagoSugeridoDesdeExtracto(ex);
    const deudaU = nNum(ex.cupoUtilizado);
    const capCierre = nNum(ex.capitalCierreLineas);
    const intEx = nNum(ex.intereses);
    const montoCorte =
      capCierre > 0
        ? redondear2(Math.min(deudaU, capCierre + intEx))
        : Math.max(0, montoLimitePago);
    const notaMovs = notaMovimientosCorteYTC(ex);
    const pushPago = (idSuf, tipo, conceptoPref, fechaISO, montoVal) => {
      const d0 = parseFechaHoraLocal(fechaISO);
      if (!d0) return;
      const diaPago = Math.min(28, d0.getDate());
      const fi = fechaISO.slice(0, 10);
      const cuentaFila = tipo === 'limite' ? 'banco' : 'tarjetaCredito';
      const notaF =
        `Estim. capital cierre + int. aprox. (E.A. ${(t.tasaEA && String(t.tasaEA).trim()) || 0}%). ${notaMovs} Pago: ${formatearNumero(montoVal, 0)}. Confirma en Gastos.`;
      extras.push({
        id: `tc-${t.id}-${idSuf}`,
        esRecordatorioTarjeta: true,
        tipoRecordatorioTarjeta: tipo,
        tarjetaId: t.id,
        concepto: `${conceptoPref} · ${nombre}`,
        monto: Math.max(0, montoVal),
        frecuencia: 'mensual',
        fechaInicio: fi,
        diaPago,
        cuenta: cuentaFila,
        categoria: firstCat,
        activo: true,
        nota: notaF,
      });
    };
    if (fc) pushPago('corte', 'corte', 'Corte TC', fc, montoCorte);
    if (fl) pushPago('limite', 'limite', 'Límite pago TC', fl, montoLimitePago);
  }
  return [...filtrados, ...extras];
}

/** Redondeo monetario 2 cifras; alinea suma con extracto y calculadora. */
function redondear2(v) {
  const x = parseFloat(v);
  if (Number.isNaN(x)) return 0;
  return Math.round(x * 100) / 100;
}

/**
 * tasa E.A. (%) a interés de un mes (aprox., mismo criterio que el extracto de tarjeta).
 */
export function tasaEfectivaMensualDesdeEAPorcentaje(tasaEAPorc) {
  const ea = parseFloat(tasaEAPorc);
  if (Number.isNaN(ea) || ea <= 0) return 0;
  return Math.pow(1 + ea / 100, 1 / 12) - 1;
}

function capitalAcumuladoCorteTC(tramos) {
  return (tramos || []).reduce((s, x) => {
    if (!x) return s;
    const c = nNum(x.montoCapitalCorteTC);
    if (c > 0) return s + c;
    return s + nNum(x.monto);
  }, 0);
}

/**
 * Un pago programado por fecha de corte: suma capital de cuotas; interés aprox. por cuota
 * (r_mes * capital) como en el extracto. `monto` = capital + interés; `montoCapitalCorteTC` = solo capital.
 */
export function agregarOFusionarPagoProgramadoCuotaCorte(pagos, p) {
  const {
    fechaCorteDate,
    monto,
    nombre,
    iCuota,
    nCuotas,
    categoria,
    cuenta,
    notaUsuario,
    tasaEA,
  } = p;
  if (!fechaCorteDate || Number.isNaN(fechaCorteDate.getTime())) return pagos || [];
  const y = fechaCorteDate.getFullYear();
  const mo = fechaCorteDate.getMonth() + 1;
  const da = fechaCorteDate.getDate();
  const ymd = `${y}-${String(mo).padStart(2, '0')}-${String(da).padStart(2, '0')}`;
  const id = `corte-tc-dif-${ymd}`;
  const linea = `${nombre} · cuota ${iCuota}/${nCuotas}`;
  const mAdd = nNum(monto);
  const rMes = tasaEfectivaMensualDesdeEAPorcentaje(tasaEA);
  const notaB = (notaUsuario && String(notaUsuario).trim()) || '';
  const list = pagos || [];
  const enMismaFecha = (x) =>
    x && x.esCuotaDiferida && String(x.fechaInicio || '').slice(0, 10) === ymd;
  const tramosCorte = list.filter((x) => enMismaFecha(x));
  const resto = list.filter((x) => !enMismaFecha(x));
  const capitalPrev = capitalAcumuladoCorteTC(tramosCorte);
  const capitalTotal = redondear2(capitalPrev + mAdd);
  const interesTotal = rMes > 0 ? redondear2(capitalTotal * rMes) : 0;
  const montoPagar = redondear2(capitalTotal + interesTotal);
  const tasaTxt = tasaEA != null && String(tasaEA).trim() !== '' ? nNum(tasaEA) : 0;
  const categoriaBase = tramosCorte.find((x) => x.categoria)?.categoria || categoria;
  const cuentaBase = tramosCorte.find((x) => x.cuenta)?.cuenta || cuenta;
  const lbl = fechaCorteDate.toLocaleDateString('es', { day: 'numeric', month: 'short' });
  const nnotaBloque = [tramosCorte.map((x) => x.nota).filter(Boolean).join(' · '), notaB, `${linea} (cuota) +${formatearNumero(mAdd, 0)} cap.`]
    .filter(Boolean)
    .join(' · ');
  const intLinea =
    rMes > 0
      ? `Cap. periodo ${formatearNumero(capitalTotal, 0)} + int. aprox. ${formatearNumero(interesTotal, 0)} (E.A. ${formatearNumero(tasaTxt, 2)}%)`
      : `Cap. periodo ${formatearNumero(capitalTotal, 0)}`;
  const notaFinal = nnotaBloque ? `${nnotaBloque} · ${intLinea}` : intLinea;
  const out = {
    id,
    concepto: tramosCorte.length
      ? `Pago corte TC — ${lbl} — ${formatearNumero(montoPagar, 0)}`
      : `${nombre} - Cuota ${iCuota} de ${nCuotas}`,
    monto: montoPagar,
    montoCapitalCorteTC: capitalTotal,
    frecuencia: 'unico',
    fechaInicio: ymd,
    diaPago: Math.min(28, da),
    cuenta: cuentaBase,
    categoria: categoriaBase,
    activo: true,
    nota: notaFinal,
    esCuotaDiferida: true,
  };
  return [...resto, out];
}

/**
 * Alertas y “reloj” por tarjeta + totales (gasto registrado en la app vs cupo).
 */
export function resumenAlertasTarjetasCredito(data, ref = new Date()) {
  const gastadoTotal = obtenerGastadoTarjetaCredito(data);
  const limiteTotal = limiteTotalTarjetasCredito(data);
  const pctGlobal = limiteTotal > 0 ? (gastadoTotal / limiteTotal) * 100 : 0;
  const arr = Array.isArray(data.tarjetasCredito) ? data.tarjetasCredito : [];

  const tarjetas = arr.map((t) => {
    const cupoT = parseFloat(t.cupoTotal) || 0;
    const cupoU = parseFloat(t.cupoUtilizado) || 0;
    const utilPct = cupoT > 0 ? (cupoU / cupoT) * 100 : 0;
    const patCorte = patronMensualDesdeTarjeta(t, 'corte');
    const patPago = patronMensualDesdeTarjeta(t, 'pago');
    const corteCicloHoy = instanteVencimientoCicloActual(patCorte, ref);
    const corteHoy = !!(
      corteCicloHoy &&
      corteCicloHoy.getFullYear() === ref.getFullYear() &&
      corteCicloHoy.getMonth() === ref.getMonth() &&
      corteCicloHoy.getDate() === ref.getDate()
    );
    const proxCorte = proximaOcurrenciaMensual(patCorte, ref);
    const proxPago = proximaOcurrenciaMensual(patPago, ref);
    const diasCorte = proxCorte ? diasCalendarioHasta(proxCorte, ref) ?? 0 : 0;
    const diasPago = proxPago ? diasCalendarioHasta(proxPago, ref) ?? 0 : 0;
    const etiquetaProxCorte = proxCorte ? proxCorte.toLocaleDateString('es', { dateStyle: 'short' }) : '';
    const etiquetaProxPago = proxPago ? proxPago.toLocaleDateString('es', { dateStyle: 'short' }) : '';
    return {
      id: t.id,
      nombreEntidad: (t.nombreEntidad && String(t.nombreEntidad).trim()) || 'Tarjeta',
      tasaEA: parseFloat(t.tasaEA) || 0,
      cupoTotal: cupoT,
      cupoUtilizado: cupoU,
      utilPct,
      corteHoy,
      diasCorte,
      diasPago,
      etiquetaProxCorte,
      etiquetaProxPago,
      alertaUtil: cupoT > 0 && utilPct >= 50,
      alertaPagoUrgente: diasPago <= 3 && diasPago >= 0,
      alertaCorte: corteHoy || (diasCorte <= 2 && diasCorte >= 0),
    };
  });

  const mostrarPorTarjeta = tarjetas.some(
    (x) => x.alertaUtil || x.alertaPagoUrgente || x.alertaCorte
  );
  /** Con tarjetas detalladas siempre mostramos el «reloj»; sin ellas, solo alertas legacy. */
  const mostrar = tarjetas.length > 0 || mostrarPorTarjeta || pctGlobal >= 50;

  return {
    mostrar,
    global: { gastado: gastadoTotal, limite: limiteTotal, porcentaje: pctGlobal },
    tarjetas,
  };
}

function nNum(v) {
  const x = Number(v);
  return Number.isFinite(x) ? x : 0;
}

export function calcularSaldosPorCuenta(data) {
  const saldosIni = obtenerSaldosIniciales(data);
  const ingresos = data.ingresos || [];
  const gastos = data.gastos || [];
  const contribuciones = data.contribucionesMetas || [];
  const limiteTc = limiteTotalTarjetasCredito(data);

  const saldos = {};
  CUENTAS.forEach((c) => {
    const ing = ingresos
      .filter((i) => cuentaBucketMovimiento(i.origen, data) === c.id)
      .reduce((s, i) => s + nNum(i.cantidad), 0);
    const gast = gastos
      .filter((g) => {
        const b = cuentaBucketMovimiento(g.origen, data);
        if (c.id === 'tarjetaCredito' && limiteTc > 0) {
          return b === 'tarjetaCredito' || g.origen === 'Tarjeta de crédito';
        }
        return b === c.id;
      })
      .reduce((s, g) => {
        const q = parseInt(g.cuotas, 10) || 1;
        const monto =
          c.id === 'tarjetaCredito' && q > 1
            ? nNum(g.cuotaMensual) || nNum(g.cantidad) / q
            : nNum(g.cantidad);
        return s + monto;
      }, 0);
    const contrib = contribuciones
      .filter((x) => cuentaBucketMovimiento(x.origen, data) === c.id)
      .reduce((s, x) => s + nNum(x.cantidad), 0);
    if (c.id === 'tarjetaCredito' && limiteTc > 0) {
      const deudaSaldo = totalCupoUtilizadoTarjetasCredito(data);
      const deudaMov = deudaGastosTarjetaAcumuladaHastaCorte(data);
      /** Tope − (cupo usado en Saldo + tramos con corte ≤ hoy desde Gastos). */
      const deuda = Math.min(nNum(limiteTc), nNum(deudaSaldo) + nNum(deudaMov));
      saldos[c.id] = Math.max(0, nNum(limiteTc) - deuda);
    } else {
      saldos[c.id] = nNum(saldosIni[c.id]) + ing - gast - contrib;
    }
  });
  /** Solo cuentas de CUENTAS, siempre numérico (Object.values mezclaba strings si hubo 0 + "100" en reservas). */
  saldos.total = CUENTAS.reduce((s, c) => s + nNum(saldos[c.id]), 0);
  saldos.totalReservado = contribuciones.reduce((s, c) => s + nNum(c.cantidad), 0);
  return saldos;
}

/**
 * Suma efectivo, banco, Nequi, Daviplata y billeteras. No incluye el cupo disponible de tarjeta (eso es otra caja).
 */
export function totalSaldoLiquido(data) {
  const s = calcularSaldosPorCuenta(data);
  return (
    (s.efectivo || 0) +
    (s.banco || 0) +
    (s.nequi || 0) +
    (s.daviplata || 0) +
    (s.billeteras || 0)
  );
}

export function montoGastoPorCuenta(g, cuentaId) {
  if (cuentaId === 'tarjetaCredito' && g.cuotas > 1) {
    return g.cuotaMensual || g.cantidad / g.cuotas || 0;
  }
  return g.cantidad || 0;
}

function patCortePrimeraTarjeta(data) {
  const arr = data?.tarjetasCredito;
  if (!Array.isArray(arr) || arr.length === 0) return null;
  return patronMensualDesdeTarjeta(arr[0], 'corte');
}

/** n-ésima fecha de corte (1 = primer corte estrictamente posterior a la compra). */
function fechaCorteNDesdeGasto(fechaGasto, pat, n) {
  if (!pat || n < 1) return null;
  let ref = parseFechaHoraLocal(fechaGasto);
  if (!ref) ref = new Date();
  let last = null;
  for (let k = 0; k < n; k++) {
    const prox = proximaOcurrenciaMensual(pat, ref);
    if (!prox) return null;
    last = prox;
    ref = prox;
  }
  return last;
}

/**
 * Fechas (una por cuota) en que “cae” el cargo según ciclos de corte. Sin tarjeta con corte en Saldo, mes a mes
 * desde la compra.
 */
export function fechasCortesParaGastoTarjeta(fechaGasto, nCuotas, data) {
  const n = Math.max(0, parseInt(nCuotas, 10) || 0);
  if (n < 1) return [];
  const r0 = parseFechaHoraLocal(fechaGasto) || new Date();
  const pat = patCortePrimeraTarjeta(data);
  if (!pat) {
    return Array.from(
      { length: n },
      (_, i) => new Date(r0.getFullYear(), r0.getMonth() + i, Math.min(28, r0.getDate()), 12, 0, 0)
    );
  }
  const out = [];
  for (let i = 1; i <= n; i++) {
    const d = fechaCorteNDesdeGasto(fechaGasto, pat, i);
    if (d && !Number.isNaN(d.getTime())) out.push(d);
  }
  return out;
}

function corteCivilAplicaHastaFecha(fechaCorte, ref) {
  if (!fechaCorte || !ref) return false;
  const a = new Date(
    fechaCorte.getFullYear(),
    fechaCorte.getMonth(),
    fechaCorte.getDate()
  );
  const b = new Date(ref.getFullYear(), ref.getMonth(), ref.getDate());
  return a.getTime() <= b.getTime();
}

/**
 * Suma acumulada (por movimientos en Gastos) de lo que ya “corre” a deuda: cada cuota cuyo corte
 * es hoy o ya ocurrió, más el contado en su primer corte. Se suma al cupo usado de Saldo para el
 * cupo libre; evita doble conteo: si anotas lo mismo en Saldo y en Gastos, ajusta el cupo usado.
 */
export function deudaGastosTarjetaAcumuladaHastaCorte(data, ref = new Date()) {
  const gastos = data.gastos || [];
  let sum = 0;
  for (const g of gastos) {
    if (normalizarOrigenCuenta(g.origen) !== 'tarjetaCredito') continue;
    const q = Math.max(1, parseInt(g.cuotas, 10) || 1);
    const fechas = fechasCortesParaGastoTarjeta(g.fecha, q, data);
    if (fechas.length === 0) continue;
    const cuo = q > 1 ? nNum(g.cuotaMensual) || nNum(g.cantidad) / q : nNum(g.cantidad);
    let nCortesYa = 0;
    for (const d of fechas) {
      if (corteCivilAplicaHastaFecha(d, ref)) nCortesYa++;
    }
    if (q > 1) {
      sum += nCortesYa * cuo;
    } else if (nCortesYa >= 1) {
      sum += nNum(g.cantidad);
    }
  }
  return sum;
}

/**
 * Importe de un gasto que se imputa a un mes calendario: tarjeta en N cuotas reparte una cuota en cada
 * mes de corte; contado, el mes del primer corte posterior a la compra.
 * Sin patrón de corte, reparte en meses consecutivos desde la fecha de compra.
 */
export function montoGastoAfectaSaldoEnMes(g, data, mes, año) {
  if (!g) return 0;
  const orig = normalizarOrigenCuenta(g.origen);
  if (orig !== 'tarjetaCredito') {
    const m = obtenerMesAño(g.fecha);
    if (m.mes === mes && m.año === año) {
      return nNum(g.cantidad);
    }
    return 0;
  }
  const q = parseInt(g.cuotas, 10) || 1;
  const cuo = q > 1 ? nNum(g.cuotaMensual) || nNum(g.cantidad) / q : nNum(g.cantidad);
  const pat = patCortePrimeraTarjeta(data);
  const r0 = parseFechaHoraLocal(g.fecha) || new Date();
  if (!pat) {
    for (let i = 0; i < q; i++) {
      const t = new Date(r0.getFullYear(), r0.getMonth() + i, Math.min(28, r0.getDate()), 12, 0, 0);
      if (t.getMonth() === mes && t.getFullYear() === año) {
        return q > 1 ? cuo : nNum(g.cantidad);
      }
    }
    return 0;
  }
  for (let i = 1; i <= q; i++) {
    const fc = fechaCorteNDesdeGasto(g.fecha, pat, i);
    if (fc && fc.getMonth() === mes && fc.getFullYear() === año) {
      return q > 1 ? cuo : nNum(g.cantidad);
    }
  }
  return 0;
}

/** Años y meses donde un gasto tiene al menos un tramo (para reportes e histórico). */
export function aniosMesesDondeAfectaGasto(g, data) {
  if (!g) return [];
  if (normalizarOrigenCuenta(g.origen) !== 'tarjetaCredito') {
    const m = obtenerMesAño(g.fecha);
    return m.mes < 0 ? [] : [{ mes: m.mes, año: m.año }];
  }
  const q = Math.max(1, parseInt(g.cuotas, 10) || 1);
  const f = fechasCortesParaGastoTarjeta(g.fecha, q, data);
  if (f.length === 0) {
    const m = obtenerMesAño(g.fecha);
    return m.mes < 0 ? [] : [{ mes: m.mes, año: m.año }];
  }
  return f.map((d) => ({ mes: d.getMonth(), año: d.getFullYear() }));
}

export function montoGastoAfectaSaldo(g) {
  if (!g) return 0;
  const orig = normalizarOrigenCuenta(g.origen);
  if (orig !== 'tarjetaCredito') return g.cantidad || 0;
  return g.cuotas > 1 ? g.cuotaMensual || (g.cantidad || 0) / g.cuotas : g.cantidad || 0;
}

/**
 * Saldo estimado por cada fila de banco (detalle en Saldo), repartiendo ingresos/gastos
 * genéricos al bucket «banco» entre líneas según su peso.
 */
export function liquidacionLineasBanco(data) {
  const lines = data.bancosDetalle || [];
  if (!lines.length) return [];

  const ingresos = data.ingresos || [];
  const gastos = data.gastos || [];
  const contrib = data.contribucionesMetas || [];

  const ingEsp = {};
  let ingGen = 0;
  for (const i of ingresos) {
    const o = String(i.origen || '');
    if (o.startsWith(PREFIJO_ORIGEN_BANCO)) {
      const id = o.slice(PREFIJO_ORIGEN_BANCO.length);
      ingEsp[id] = (ingEsp[id] || 0) + (parseFloat(i.cantidad) || 0);
    } else if (cuentaBucketMovimiento(o, data) === 'banco') {
      ingGen += parseFloat(i.cantidad) || 0;
    }
  }

  const contribEsp = {};
  let contribGen = 0;
  for (const x of contrib) {
    const o = String(x.origen || '');
    if (o.startsWith(PREFIJO_ORIGEN_BANCO)) {
      const id = o.slice(PREFIJO_ORIGEN_BANCO.length);
      contribEsp[id] = (contribEsp[id] || 0) + (parseFloat(x.cantidad) || 0);
    } else if (cuentaBucketMovimiento(o, data) === 'banco') {
      contribGen += parseFloat(x.cantidad) || 0;
    }
  }

  let gastGen = 0;
  for (const g of gastos) {
    if (cuentaBucketMovimiento(g.origen, data) === 'banco') {
      gastGen += montoGastoAfectaSaldo(g);
    }
  }

  const base = lines.map((r) => ({
    id: r.id,
    nombre: String(r.nombre || '').trim() || 'Banco',
    ini: parseFloat(r.saldo) || 0,
    ingEsp: ingEsp[r.id] || 0,
    contribEsp: contribEsp[r.id] || 0,
  }));

  const peso = base.map((b) => Math.max(0, b.ini + b.ingEsp));
  const sumPeso = peso.reduce((a, b) => a + b, 0);
  function distribuir(monto) {
    if (!monto) return base.map(() => 0);
    if (sumPeso > 0) return peso.map((p) => monto * (p / sumPeso));
    const n = base.length;
    return base.map(() => monto / n);
  }
  const addIngGen = distribuir(ingGen);

  const pre = base.map((b, i) => b.ini + b.ingEsp + addIngGen[i]);
  const peso2 = pre.map((p) => Math.max(0, p));
  const sum2 = peso2.reduce((a, b) => a + b, 0);
  function distribuir2(monto) {
    if (!monto) return base.map(() => 0);
    if (sum2 > 0) return peso2.map((p) => monto * (p / sum2));
    const n = base.length;
    return base.map(() => monto / n);
  }
  const subGast = distribuir2(gastGen);
  const subContribG = distribuir2(contribGen);

  return base.map((b, i) => ({
    id: b.id,
    nombre: b.nombre,
    saldo: pre[i] - subGast[i] - b.contribEsp - subContribG[i],
  }));
}

function liquidacionSubPlataformaBucket(data, bucket) {
  const lines = (data.plataformasDetalle || []).filter(
    (r) => r && cuentaSaldoPlataforma(r.platformValue || PLATAFORMA_OTRO_VALUE) === bucket
  );
  if (!lines.length) return [];

  const ingresos = data.ingresos || [];
  const gastos = data.gastos || [];
  const contrib = data.contribucionesMetas || [];

  const ingEsp = {};
  let ingGen = 0;
  for (const i of ingresos) {
    const o = String(i.origen || '');
    if (o.startsWith(PREFIJO_ORIGEN_PLATAFORMA)) {
      const id = o.slice(PREFIJO_ORIGEN_PLATAFORMA.length);
      ingEsp[id] = (ingEsp[id] || 0) + (parseFloat(i.cantidad) || 0);
    } else if (cuentaBucketMovimiento(o, data) === bucket) {
      ingGen += parseFloat(i.cantidad) || 0;
    }
  }

  const contribEsp = {};
  let contribGen = 0;
  for (const x of contrib) {
    const o = String(x.origen || '');
    if (o.startsWith(PREFIJO_ORIGEN_PLATAFORMA)) {
      const id = o.slice(PREFIJO_ORIGEN_PLATAFORMA.length);
      contribEsp[id] = (contribEsp[id] || 0) + (parseFloat(x.cantidad) || 0);
    } else if (cuentaBucketMovimiento(o, data) === bucket) {
      contribGen += parseFloat(x.cantidad) || 0;
    }
  }

  let gastGen = 0;
  for (const g of gastos) {
    if (cuentaBucketMovimiento(g.origen, data) === bucket) {
      gastGen += montoGastoAfectaSaldo(g);
    }
  }

  const base = lines.map((r) => ({
    id: r.id,
    nombre: String(r.nombre || '').trim() || getPlataformaLabelByValue(r.platformValue) || 'Plataforma',
    ini: parseFloat(r.saldo) || 0,
    ingEsp: ingEsp[r.id] || 0,
    contribEsp: contribEsp[r.id] || 0,
  }));

  const peso = base.map((b) => Math.max(0, b.ini + b.ingEsp));
  const sumPeso = peso.reduce((a, b) => a + b, 0);
  function distribuir(monto) {
    if (!monto) return base.map(() => 0);
    if (sumPeso > 0) return peso.map((p) => monto * (p / sumPeso));
    const n = base.length;
    return base.map(() => monto / n);
  }
  const addIngGen = distribuir(ingGen);

  const pre = base.map((b, i) => b.ini + b.ingEsp + addIngGen[i]);
  const peso2 = pre.map((p) => Math.max(0, p));
  const sum2 = peso2.reduce((a, b) => a + b, 0);
  function distribuir2(monto) {
    if (!monto) return base.map(() => 0);
    if (sum2 > 0) return peso2.map((p) => monto * (p / sum2));
    const n = base.length;
    return base.map(() => monto / n);
  }
  const subGast = distribuir2(gastGen);
  const subContribG = distribuir2(contribGen);

  return base.map((b, i) => ({
    id: b.id,
    nombre: b.nombre,
    saldo: pre[i] - subGast[i] - b.contribEsp - subContribG[i],
  }));
}

export function liquidacionLineasPlataforma(data) {
  const out = [];
  for (const bucket of ['nequi', 'daviplata', 'billeteras']) {
    out.push(...liquidacionSubPlataformaBucket(data, bucket));
  }
  return out;
}

/**
 * Cuentas para destino de un ingreso (y vista de referencia de saldos en Gastos).
 * Incluye saldo 0,00 o negativo, para que sigas viendo cajas y elijas a dónde entra plata aunque estés en cero.
 */
export function obtenerCuentasDestinoIngreso(data) {
  const saldos = calcularSaldosPorCuenta(data);
  const limiteTc = limiteTotalTarjetasCredito(data);
  const moneda = (data.moneda && String(data.moneda).trim()) || '';
  const suf = moneda ? ` ${moneda}` : '';
  const out = [];

  const push = (value, nombreDisplay, saldo) => {
    const n = parseFloat(saldo);
    const v = Number.isNaN(n) ? 0 : n;
    out.push({
      value,
      label: `${nombreDisplay} (${formatearNumero(v)}${suf})`,
      saldo: v,
    });
  };

  push('efectivo', 'Efectivo', saldos.efectivo);

  const bancos = data.bancosDetalle || [];
  if (bancos.length > 0) {
    liquidacionLineasBanco(data).forEach((row) => {
      push(`${PREFIJO_ORIGEN_BANCO}${row.id}`, row.nombre, row.saldo);
    });
  } else {
    const nombre = CUENTAS.find((c) => c.id === 'banco')?.nombre || 'Banco';
    push('banco', nombre, saldos.banco);
  }

  if (limiteTc > 0) {
    const nombre = CUENTAS.find((c) => c.id === 'tarjetaCredito')?.nombre || 'Tarjeta de crédito';
    push('tarjetaCredito', nombre, saldos.tarjetaCredito);
  }

  const plt = data.plataformasDetalle || [];
  if (plt.length > 0) {
    liquidacionLineasPlataforma(data).forEach((row) => {
      push(`${PREFIJO_ORIGEN_PLATAFORMA}${row.id}`, row.nombre, row.saldo);
    });
  } else {
    push('nequi', 'Nequi', saldos.nequi);
    push('daviplata', 'Daviplata', saldos.daviplata);
    const nb = CUENTAS.find((c) => c.id === 'billeteras')?.nombre || 'Otras plataformas';
    push('billeteras', nb, saldos.billeteras);
  }

  return out;
}

/**
 * Saldo disponible hoy en la cuenta/ línea con la que se registra un gasto o ingreso (misma idea que en Saldo por fila).
 */
export function obtenerSaldoDisponibleParaOrigenMovimiento(data, origen) {
  const o = String(origen || '').trim();
  if (o === 'Tarjeta de crédito') {
    return Math.max(0, calcularSaldosPorCuenta(data).tarjetaCredito || 0);
  }
  if (o === 'efectivo' || o === 'tarjetaCredito' || o === 'banco' || o === 'nequi' || o === 'daviplata' || o === 'billeteras') {
    return Math.max(0, calcularSaldosPorCuenta(data)[o] || 0);
  }
  if (o.startsWith(PREFIJO_ORIGEN_BANCO)) {
    const id = o.slice(PREFIJO_ORIGEN_BANCO.length);
    const row = (liquidacionLineasBanco(data) || []).find((r) => r.id === id);
    return row ? Math.max(0, row.saldo) : 0;
  }
  if (o.startsWith(PREFIJO_ORIGEN_PLATAFORMA)) {
    const id = o.slice(PREFIJO_ORIGEN_PLATAFORMA.length);
    const row = (liquidacionLineasPlataforma(data) || []).find((r) => r.id === id);
    return row ? Math.max(0, row.saldo) : 0;
  }
  const b = normalizarOrigenCuenta(o) || o;
  if (b === 'tarjetaCredito' || b === 'efectivo' || b === 'banco' || b === 'nequi' || b === 'daviplata' || b === 'billeteras') {
    return Math.max(0, calcularSaldosPorCuenta(data)[b] || 0);
  }
  return 0;
}

/**
 * Cuentas desde las que se puede pagar un gasto de `monto` (una sola carga) o, en TC, la cuota mensual indicada.
 * Incluye filas de banco y plataforma con saldo por línea, como en Ingresos.
 */
export function obtenerCuentasOrigenGastoElegible(data, monto, cuotaMensualTarjeta) {
  const m = Math.max(0, parseFloat(monto) || 0);
  const s = calcularSaldosPorCuenta(data);
  const moneda = (data.moneda && String(data.moneda).trim()) || '';
  const suf = moneda ? ` ${moneda}` : '';
  const requiereTc = Math.max(0, parseFloat(cuotaMensualTarjeta) || 0) || m;
  const out = [];
  const pushGasto = (value, nombreDisplay, saldo, esTarjeta) => {
    const requiere = esTarjeta ? requiereTc : m;
    if (requiere === 0) {
      if (saldo > 0) {
        out.push({ value, label: `${nombreDisplay} (${formatearNumero(saldo)}${suf})`, saldo, esTarjeta: !!esTarjeta });
      }
      return;
    }
    if (saldo > 0 && saldo >= requiere) {
      out.push({ value, label: `${nombreDisplay} (${formatearNumero(saldo)}${suf})`, saldo, esTarjeta: !!esTarjeta });
    }
  };

  if (s.efectivo > 0) pushGasto('efectivo', 'Efectivo', s.efectivo, false);

  const bancos = data.bancosDetalle || [];
  if (bancos.length > 0) {
    liquidacionLineasBanco(data).forEach((row) => {
      pushGasto(`${PREFIJO_ORIGEN_BANCO}${row.id}`, row.nombre, row.saldo, false);
    });
  } else if (s.banco > 0) {
    const nb = CUENTAS.find((c) => c.id === 'banco')?.nombre || 'Banco';
    pushGasto('banco', nb, s.banco, false);
  }

  if (s.tarjetaCredito > 0) {
    const ntc = CUENTAS.find((c) => c.id === 'tarjetaCredito')?.nombre || 'Tarjeta de crédito';
    pushGasto('tarjetaCredito', ntc, s.tarjetaCredito, true);
  }

  const plt = data.plataformasDetalle || [];
  if (plt.length > 0) {
    liquidacionLineasPlataforma(data).forEach((row) => {
      pushGasto(`${PREFIJO_ORIGEN_PLATAFORMA}${row.id}`, row.nombre, row.saldo, false);
    });
  } else {
    if (s.nequi > 0) pushGasto('nequi', 'Nequi', s.nequi, false);
    if (s.daviplata > 0) pushGasto('daviplata', 'Daviplata', s.daviplata, false);
    if (s.billeteras > 0) {
      const nb = CUENTAS.find((c) => c.id === 'billeteras')?.nombre || 'Otras plataformas';
      pushGasto('billeteras', nb, s.billeteras, false);
    }
  }

  return out;
}

/** Suma de tramos con corte ≤ hoy (coherente con el cupo libre y fechas de corte). */
export function obtenerGastadoTarjetaCredito(data, ref) {
  return deudaGastosTarjetaAcumuladaHastaCorte(data, ref || new Date());
}

export function verificarAlertaTarjetaCredito(data) {
  const r = resumenAlertasTarjetasCredito(data);
  return {
    mostrar: r.mostrar,
    gastado: r.global.gastado,
    limite: r.global.limite,
    porcentaje: r.global.porcentaje,
    tarjetas: r.tarjetas,
  };
}

/**
 * Mismas fechas de corte que `fechasCortesParaGastoTarjeta`, pero con el patrón (día) de `pat` dado
 * (p. ej. la tarjeta concreta), no la primera de la lista.
 */
export function fechasCortesParaGastoYPat(fechaGasto, nCuotas, pat) {
  const n = Math.max(0, parseInt(nCuotas, 10) || 0);
  if (n < 1) return [];
  const r0 = parseFechaHoraLocal(fechaGasto) || new Date();
  if (!pat) {
    return Array.from(
      { length: n },
      (_, i) => new Date(r0.getFullYear(), r0.getMonth() + i, Math.min(28, r0.getDate()), 12, 0, 0)
    );
  }
  const out = [];
  for (let i = 1; i <= n; i++) {
    const d = fechaCorteNDesdeGasto(fechaGasto, pat, i);
    if (d && !Number.isNaN(d.getTime())) out.push(d);
  }
  return out;
}

function fraccionCupoDeTarjeta(t, data) {
  const limT = nNum(t?.cupoTotal);
  const limS = limiteTotalTarjetasCredito(data);
  if (limS <= 0) return 1;
  if (limT <= 0) return 0;
  return Math.min(1, limT / limS);
}

function lineasMovimientosCorteDia(t, data, ref) {
  const pat = patronMensualDesdeTarjeta(t, 'corte');
  const y = ref.getFullYear();
  const m0 = ref.getMonth();
  const d0 = ref.getDate();
  const out = [];
  for (const g of data.gastos || []) {
    if (normalizarOrigenCuenta(g.origen) !== 'tarjetaCredito') continue;
    const q = Math.max(1, parseInt(g.cuotas, 10) || 1);
    const fechas = fechasCortesParaGastoYPat(g.fecha, q, pat);
    const cuo = q > 1 ? nNum(g.cuotaMensual) || nNum(g.cantidad) / q : nNum(g.cantidad);
    for (let i = 0; i < fechas.length; i++) {
      const d = fechas[i];
      if (d.getFullYear() === y && d.getMonth() === m0 && d.getDate() === d0) {
        out.push({
          descripcion: String(g.nombre || 'Compra').trim() || 'Compra',
          categoria: g.categoria || '—',
          capitalMes: cuo,
          cuotaLabel: q > 1 ? `Cuota ${i + 1} de ${q}` : 'Un pago (corte del cargo)',
        });
      }
    }
  }
  return out;
}

function capitalDeLineasODeuda(lineas, deuda) {
  const s = lineas.reduce((a, b) => a + nNum(b.capitalMes), 0);
  if (s > 0) return s;
  return nNum(deuda) > 0 ? nNum(deuda) : 0;
}

/** Pago a cuotas a tasa fija (cuota fija) — coste total aprox. con tasa E.A. a periodo mensual. */
export function proyeccionCostoTotalCuotasFijas(principal, tasaEAPorc, nCuotas) {
  const P = Math.max(0, nNum(principal));
  const p = Math.max(1, parseInt(nCuotas, 10) || 1);
  if (P <= 0) return 0;
  const ea = Math.max(0, nNum(tasaEAPorc) / 100);
  if (ea <= 0) return P;
  const im = Math.pow(1 + ea, 1 / 12) - 1;
  if (im <= 0) return P;
  const cuota = (P * im * Math.pow(1 + im, p)) / (Math.pow(1 + im, p) - 1);
  return p * cuota;
}

/**
 * Estructura de «extracto» para UI: fechas, cupo, detalle, totales, proyección básica a 3/6 plazos.
 * Interés: si hay movimientos que cierran hoy, r_mes × capital del periodo (misma fórmula que pago programado
 * por corte). Si no hay tramos hoy, estimación sobre deuda total (revolving sin “nuevo” cierre).
 */
export function construirExtractoBancarioTarjeta(t, data, ref = new Date()) {
  const patP = patronMensualDesdeTarjeta(t, 'pago');
  const limT = nNum(t.cupoTotal);
  const limS = limiteTotalTarjetasCredito(data);
  const cupoU = nNum(t.cupoUtilizado);
  const frac = fraccionCupoDeTarjeta(t, data);
  const deudaMovG = nNum(deudaGastosTarjetaAcumuladaHastaCorte(data, ref)) * frac;
  const deudaEfect = Math.min(limT > 0 ? limT : limS, cupoU + deudaMovG);
  const disp = limT > 0 ? Math.max(0, limT - deudaEfect) : 0;
  const lineas = lineasMovimientosCorteDia(t, data, ref);
  const capitalCierreLineas = redondear2(lineas.reduce((s, l) => s + nNum(l.capitalMes), 0));
  const capitalPeriodo = capitalCierreLineas > 0 ? capitalCierreLineas : capitalDeLineasODeuda(lineas, deudaEfect);
  const ea = nNum(t.tasaEA) || 0;
  const rMes = ea > 0 ? Math.pow(1 + ea / 100, 1 / 12) - 1 : 0;
  const interesSobreDeuda = deudaEfect > 0 && rMes > 0 ? redondear2(nNum(deudaEfect) * rMes) : 0;
  const interesSobreCapCierre =
    capitalCierreLineas > 0 && rMes > 0 ? redondear2(nNum(capitalCierreLineas) * rMes) : 0;
  const interesesEst = capitalCierreLineas > 0 ? interesSobreCapCierre : interesSobreDeuda;
  const fijos = 0;
  const pagoMin = deudaEfect > 0 ? Math.min(deudaEfect, Math.max(0, deudaEfect * 0.03)) : 0;
  const pagoTotal = deudaEfect;
  const proxPago = patP ? proximaOcurrenciaMensual(patP, ref) : null;
  const t3 = proyeccionCostoTotalCuotasFijas(deudaEfect, ea, 3);
  const t6 = proyeccionCostoTotalCuotasFijas(deudaEfect, ea, 6);
  return {
    nombre: String(t.nombreEntidad || 'Tarjeta'),
    etiquetaCorte: ref.toLocaleDateString('es', { dateStyle: 'long' }),
    etiquetaProxPago:
      proxPago && !Number.isNaN(proxPago.getTime())
        ? proxPago.toLocaleDateString('es', { dateStyle: 'long' })
        : '—',
    cupoTotal: limT,
    cupoUtilizado: deudaEfect,
    cupoDisponible: disp,
    lineas,
    /** Suma de cuotas de Gastos cuyo corte es exactamente el día de `ref` (sin rellenar con deuda). */
    capitalCierreLineas,
    capitalPeriodo: capitalCierreLineas > 0 ? capitalCierreLineas : capitalDeLineasODeuda(lineas, deudaEfect),
    intereses: interesesEst,
    costosFijos: fijos,
    pagoMinimo: pagoMin,
    pagoTotalObl: pagoTotal,
    proy3: t3,
    proy6: t6,
    ahorro6vs3: Math.max(0, t6 - t3),
    tasaEA: ea,
  };
}

/**
 * Pago a programar: capital del cierre (cuotas en el periodo) + interés prorrateado, o pago mín. / deuda.
 */
export function montoPagoSugeridoDesdeExtracto(ex) {
  if (!ex) return 0;
  const deuda = nNum(ex.cupoUtilizado);
  const capCierre = nNum(ex.capitalCierreLineas);
  if (deuda <= 0) return 0;
  if (capCierre > 0) {
    return redondear2(Math.min(deuda, capCierre + nNum(ex.intereses)));
  }
  if (nNum(ex.pagoMinimo) > 0) return nNum(ex.pagoMinimo);
  return Math.min(deuda, nNum(ex.pagoTotalObl) || deuda);
}

function notaMovimientosCorteYTC(ex) {
  const lineas = ex.lineas || [];
  if (lineas.length === 0) {
    return 'Sin tramos a este corte hoy. Deuda = cupo y mov. en Gastos; revisa en Saldo.';
  }
  const partes = lineas.slice(0, 4).map((l) => {
    const c = l.cuotaLabel ? `${l.cuotaLabel} · ` : '';
    return `${c}${l.descripcion} ${formatearNumero(nNum(l.capitalMes), 0)}`.replace(/\s+/g, ' ').trim();
  });
  const mas = lineas.length > 4 ? ` +${lineas.length - 4} movimientos` : '';
  return partes.join(' · ') + mas;
}

/**
 * Bloque de Inicio: proy. 3 vs 6 cuotas sobre deuda + tasa (primera tarjeta con tasa, si existe).
 */
export function proyeccionEficienciaInicio(data, ref = new Date()) {
  const arr = data?.tarjetasCredito;
  if (!Array.isArray(arr) || arr.length === 0) return null;
  const t = arr.find((x) => nNum(x.tasaEA) > 0) || arr[0];
  if (!t) return null;
  const limS = limiteTotalTarjetasCredito(data);
  const dSal = totalCupoUtilizadoTarjetasCredito(data);
  const dG = deudaGastosTarjetaAcumuladaHastaCorte(data, ref);
  const deudaE = limS > 0 ? Math.min(limS, dSal + dG) : 0;
  const ea = nNum(t.tasaEA) || 0;
  if (deudaE <= 0) return null;
  const t3 = proyeccionCostoTotalCuotasFijas(deudaE, ea, 3);
  const t6 = proyeccionCostoTotalCuotasFijas(deudaE, ea, 6);
  return {
    nombre: String(t.nombreEntidad || 'Tarjeta'),
    tasa: ea,
    deuda: deudaE,
    total3: t3,
    total6: t6,
    ahorro: Math.max(0, t6 - t3),
  };
}

/**
 * Mes/año calendario del movimiento. Para `YYYY-MM-DD` (como en Ingresos/Gastos) se lee el mes desde el
 * texto para no depender de `Date` y zonas horarias (p. ej. 27 → 26 al parsear en UTC).
 */
export function obtenerMesAño(fechaStr) {
  if (fechaStr == null) return { mes: -1, año: -1 };
  const s = String(fechaStr).trim();
  if (!s) return { mes: -1, año: -1 };
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) {
    const y = parseInt(m[1], 10);
    const mon = parseInt(m[2], 10) - 1;
    if (Number.isFinite(y) && mon >= 0 && mon <= 11) {
      return { mes: mon, año: y };
    }
  }
  const d = new Date(s.includes('T') ? s : `${s}T12:00:00`);
  if (Number.isNaN(d.getTime())) return { mes: -1, año: -1 };
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

/** Nombre de icono Ionicons (contorno) — distinto al emoji de categorías. */
export const ICONO_META_POR_DEFECTO = 'trophy-outline';

export function normalizarMeta(meta) {
  if (!meta || typeof meta !== 'object') {
    return { id: '', nombre: '', objetivo: 0, plazo: null, icono: ICONO_META_POR_DEFECTO };
  }
  return {
    ...meta,
    icono: typeof meta.icono === 'string' && meta.icono.trim() ? meta.icono.trim() : ICONO_META_POR_DEFECTO,
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
    const anchor = parseFechaHoraLocal(pago.fechaInicio);
    const diaPagoEff = anchor
      ? Math.min(28, anchor.getDate())
      : Math.min(28, parseInt(pago.diaPago, 10) || 1);
    if (diaHoy !== diaPagoEff) return false;
    const rawFi = String(pago.fechaInicio || '');
    if (anchor && rawFi.length > 10) {
      const due = new Date(añoHoy, mesHoy, diaHoy, anchor.getHours(), anchor.getMinutes(), 0);
      if (hoy.getTime() < due.getTime()) return false;
    }
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
    const fechaPago = parseFechaHoraLocal(pago.fechaInicio);
    if (!fechaPago) return false;
    return (
      hoy.getFullYear() === fechaPago.getFullYear() &&
      hoy.getMonth() === fechaPago.getMonth() &&
      hoy.getDate() === fechaPago.getDate()
    );
  }
  return false;
}

export function pagoDebeMostrarseParaPagar(pago, ahora = new Date()) {
  if (!pago || pago.activo === false) return false;
  if (pago.frecuencia === 'unico') {
    if (!pago.fechaInicio) return false;
    const raw = String(pago.fechaInicio);
    const fechaPago = parseFechaHoraLocal(raw);
    if (!fechaPago) return false;
    if (raw.length <= 10) {
      const h0 = new Date(ahora.getFullYear(), ahora.getMonth(), ahora.getDate()).getTime();
      const f0 = new Date(fechaPago.getFullYear(), fechaPago.getMonth(), fechaPago.getDate()).getTime();
      return f0 <= h0;
    }
    return ahora.getTime() >= fechaPago.getTime();
  }
  if (pago.esRecordatorioTarjeta) {
    const anchor = parseFechaHoraLocal(pago.fechaInicio);
    if (!anchor) return true;
    const patron = new Date(2020, 0, anchor.getDate(), 12, 0, 0);
    const venc = instanteVencimientoCicloActual(patron, ahora);
    if (!venc) return true;
    const dias = diasCalendarioHasta(venc, ahora);
    if (dias != null && dias > 14) return false;
    if (dias != null && dias < -3) return false;
    return true;
  }
  return true;
}
