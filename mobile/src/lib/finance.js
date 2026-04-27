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

/** Suma de cupos totales configurados; si no hay tarjetas detalladas, usa limiteTarjetaCredito legacy. */
export function limiteTotalTarjetasCredito(data) {
  const arr = data.tarjetasCredito;
  if (Array.isArray(arr) && arr.length > 0) {
    return arr.reduce((s, t) => s + (parseFloat(t.cupoTotal) || 0), 0);
  }
  return parseFloat(data.limiteTarjetaCredito) || 0;
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
 */
export function reemplazarPagosRecordatorioTarjetas(pagosExistentes, tarjetasCredito, categorias) {
  const firstCat =
    Array.isArray(categorias) && categorias.length > 0
      ? typeof categorias[0] === 'string'
        ? categorias[0]
        : categorias[0].nombre || 'Otros'
      : 'Otros';
  const filtrados = (pagosExistentes || []).filter((p) => !p.esRecordatorioTarjeta);
  const extras = [];
  for (const t of tarjetasCredito || []) {
    const nombre = String(t.nombreEntidad || '').trim();
    if (!nombre) continue;
    const fc = String(t.fechaHoraCorte || '').trim();
    const fl = String(t.fechaHoraLimitePago || '').trim();
    const monto = Math.max(parseFloat(t.cupoUtilizado) || 0, 0);
    const pushPago = (idSuf, tipo, conceptoPref, fechaISO) => {
      const d0 = parseFechaHoraLocal(fechaISO);
      if (!d0) return;
      const diaPago = Math.min(28, d0.getDate());
      const fi = fechaISO.slice(0, 10);
      extras.push({
        id: `tc-${t.id}-${idSuf}`,
        esRecordatorioTarjeta: true,
        tipoRecordatorioTarjeta: tipo,
        tarjetaId: t.id,
        concepto: `${conceptoPref} · ${nombre}`,
        monto,
        frecuencia: 'mensual',
        fechaInicio: fi,
        diaPago,
        cuenta: 'tarjetaCredito',
        categoria: firstCat,
        activo: true,
        nota: 'Recordatorio desde Tarjetas (Saldo). Confirma en Gastos.',
      });
    };
    if (fc) pushPago('corte', 'corte', 'Corte TC', fc);
    if (fl) pushPago('limite', 'limite', 'Límite pago TC', fl);
  }
  return [...filtrados, ...extras];
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
      diasCorte,
      diasPago,
      etiquetaProxCorte,
      etiquetaProxPago,
      alertaUtil: cupoT > 0 && utilPct >= 50,
      alertaPagoUrgente: diasPago <= 3 && diasPago >= 0,
      alertaCorte: diasCorte <= 2 && diasCorte >= 0,
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
      .filter((x) => cuentaBucketMovimiento(x.origen, data) === c.id)
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
 * Cuentas con saldo > 0 para elegir destino de un ingreso (etiqueta con monto y moneda).
 */
export function obtenerCuentasDestinoIngreso(data) {
  const saldos = calcularSaldosPorCuenta(data);
  const moneda = (data.moneda && String(data.moneda).trim()) || '';
  const suf = moneda ? ` ${moneda}` : '';
  const out = [];

  const push = (value, nombreDisplay, saldo) => {
    if (saldo > 0) {
      out.push({
        value,
        label: `${nombreDisplay} (${formatearNumero(saldo)}${suf})`,
        saldo,
      });
    }
  };

  if (saldos.efectivo > 0) push('efectivo', 'Efectivo', saldos.efectivo);

  const bancos = data.bancosDetalle || [];
  if (bancos.length > 0) {
    liquidacionLineasBanco(data).forEach((row) => {
      push(`${PREFIJO_ORIGEN_BANCO}${row.id}`, row.nombre, row.saldo);
    });
  } else if (saldos.banco > 0) {
    const nombre = CUENTAS.find((c) => c.id === 'banco')?.nombre || 'Banco';
    push('banco', nombre, saldos.banco);
  }

  if (saldos.tarjetaCredito > 0) {
    const nombre = CUENTAS.find((c) => c.id === 'tarjetaCredito')?.nombre || 'Tarjeta de crédito';
    push('tarjetaCredito', nombre, saldos.tarjetaCredito);
  }

  const plt = data.plataformasDetalle || [];
  if (plt.length > 0) {
    liquidacionLineasPlataforma(data).forEach((row) => {
      push(`${PREFIJO_ORIGEN_PLATAFORMA}${row.id}`, row.nombre, row.saldo);
    });
  } else {
    if (saldos.nequi > 0) push('nequi', 'Nequi', saldos.nequi);
    if (saldos.daviplata > 0) push('daviplata', 'Daviplata', saldos.daviplata);
    if (saldos.billeteras > 0) {
      const nombre = CUENTAS.find((c) => c.id === 'billeteras')?.nombre || 'Otras plataformas';
      push('billeteras', nombre, saldos.billeteras);
    }
  }

  return out;
}

export function obtenerGastadoTarjetaCredito(data) {
  const gastos = data.gastos || [];
  return gastos
    .filter((g) => normalizarOrigenCuenta(g.origen) === 'tarjetaCredito')
    .reduce((s, g) => s + montoGastoPorCuenta(g, 'tarjetaCredito'), 0);
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
