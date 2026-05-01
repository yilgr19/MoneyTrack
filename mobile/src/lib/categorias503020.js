/**
 * Regla 50/30/20: grupos de categorías y utilidades para dashboard.
 */

export const GRUPO_NECESIDADES = 'necesidades';
export const GRUPO_DESEOS = 'deseos';
export const GRUPO_AHORRO_DEUDA = 'ahorro_deuda';

export const GRUPO503020_TODOS = [GRUPO_NECESIDADES, GRUPO_DESEOS, GRUPO_AHORRO_DEUDA];

export const ETIQUETA_GRUPO_503020 = {
  [GRUPO_NECESIDADES]: 'Necesidades (50%)',
  [GRUPO_DESEOS]: 'Deseos (30%)',
  [GRUPO_AHORRO_DEUDA]: 'Ahorro / deuda (20%)',
};

/** Meta de gasto como fracción del ingreso del mes (50/30/20). */
export const META_FRACCION_GRUPO = {
  [GRUPO_NECESIDADES]: 0.5,
  [GRUPO_DESEOS]: 0.3,
  [GRUPO_AHORRO_DEUDA]: 0.2,
};

/**
 * Categorías sugeridas alineadas al prompt (Ionicons + grupo + color_hex).
 * icono emoji opcional para compatibilidad con listas que aún usan emoji.
 */
export const CATEGORIAS_POR_DEFECTO_503020 = [
  {
    nombre: 'Vivienda',
    color_hex: '#6366f1',
    color: '#6366f1',
    iconoIon: 'home-outline',
    icono: '🏠',
    grupo503020: GRUPO_NECESIDADES,
    limite: null,
  },
  {
    nombre: 'Transporte',
    color_hex: '#3b82f6',
    color: '#3b82f6',
    iconoIon: 'car-outline',
    icono: '🚗',
    grupo503020: GRUPO_NECESIDADES,
    limite: null,
  },
  {
    nombre: 'Alimentación / Súper',
    color_hex: '#22c55e',
    color: '#22c55e',
    iconoIon: 'basket-outline',
    icono: '🛒',
    grupo503020: GRUPO_NECESIDADES,
    limite: null,
  },
  {
    nombre: 'Salud',
    color_hex: '#ef4444',
    color: '#ef4444',
    iconoIon: 'pulse-outline',
    icono: '💊',
    grupo503020: GRUPO_NECESIDADES,
    limite: null,
  },
  {
    nombre: 'Educación',
    color_hex: '#8b5cf6',
    color: '#8b5cf6',
    iconoIon: 'school-outline',
    icono: '📚',
    grupo503020: GRUPO_NECESIDADES,
    limite: null,
  },
  {
    nombre: 'Restaurantes / Café',
    color_hex: '#ec4899',
    color: '#ec4899',
    iconoIon: 'restaurant-outline',
    icono: '☕',
    grupo503020: GRUPO_DESEOS,
    limite: null,
  },
  {
    nombre: 'Entretenimiento / Streaming',
    color_hex: '#a855f7',
    color: '#a855f7',
    iconoIon: 'play-circle-outline',
    icono: '🎬',
    grupo503020: GRUPO_DESEOS,
    limite: null,
  },
  {
    nombre: 'Viajes',
    color_hex: '#0ea5e9',
    color: '#0ea5e9',
    iconoIon: 'airplane-outline',
    icono: '✈️',
    grupo503020: GRUPO_DESEOS,
    limite: null,
  },
  {
    nombre: 'Hobbies',
    color_hex: '#f97316',
    color: '#f97316',
    iconoIon: 'game-controller-outline',
    icono: '🎮',
    grupo503020: GRUPO_DESEOS,
    limite: null,
  },
  {
    nombre: 'Inversiones',
    color_hex: '#14b8a6',
    color: '#14b8a6',
    iconoIon: 'stats-chart-outline',
    icono: '📈',
    grupo503020: GRUPO_AHORRO_DEUDA,
    limite: null,
  },
  {
    nombre: 'Fondo de emergencia',
    color_hex: '#eab308',
    color: '#eab308',
    iconoIon: 'shield-checkmark-outline',
    icono: '🛡️',
    grupo503020: GRUPO_AHORRO_DEUDA,
    limite: null,
  },
  {
    nombre: 'Pago de tarjetas',
    color_hex: '#64748b',
    color: '#64748b',
    iconoIon: 'card-outline',
    icono: '💳',
    grupo503020: GRUPO_AHORRO_DEUDA,
    limite: null,
  },
];

/** Ionicons adicionales para personalizar categorías (nombre válido en @expo/vector-icons Ionicons). */
export const CATALOGO_ICONOS_ION_CATEGORIA = [
  'paw-outline',
  'barbell-outline',
  'gift-outline',
  'construct-outline',
  'baby-outline',
  'laptop-outline',
  'musical-notes-outline',
  'film-outline',
  'football-outline',
  'cafe-outline',
  'beer-outline',
  'fast-food-outline',
  'shirt-outline',
  'watch-outline',
  'phone-portrait-outline',
  'wifi-outline',
  'flash-outline',
  'leaf-outline',
  'flower-outline',
  'book-outline',
  'briefcase-outline',
  'medkit-outline',
  'cut-outline',
  'color-palette-outline',
  'train-outline',
  'bus-outline',
  'bicycle-outline',
];

const MAP_NOMBRE_A_GRUPO = (() => {
  const m = new Map();
  const add = (keys, grupo) => {
    for (const k of keys) {
      m.set(k, grupo);
    }
  };
  CATEGORIAS_POR_DEFECTO_503020.forEach((c) => {
    m.set(normalizarNombreClave(c.nombre), c.grupo503020);
  });
  add(
    [
      'supermercado',
      'hogar',
      'servicios',
      'ropa',
      'tecnologia',
      'salud',
      'educacion',
      'transporte',
      'vivienda',
      'alimentacion',
      'super',
      'mercado',
    ],
    GRUPO_NECESIDADES
  );
  add(
    [
      'entretenimiento',
      'restaurantes',
      'regalos',
      'viajes',
      'hobbies',
      'café',
      'cafe',
      'streaming',
    ],
    GRUPO_DESEOS
  );
  add(['ahorro', 'deuda', 'inversiones', 'inversión', 'emergencia', 'fondo'], GRUPO_AHORRO_DEUDA);
  return m;
})();

function normalizarNombreClave(s) {
  return String(s || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

/**
 * Infiere grupo 50/30/20 por nombre de categoría (usuarios previos sin campo grupo503020).
 */
export function inferirGrupo503020DesdeNombre(nombre) {
  const k = normalizarNombreClave(nombre);
  if (!k) return GRUPO_DESEOS;
  if (MAP_NOMBRE_A_GRUPO.has(k)) return MAP_NOMBRE_A_GRUPO.get(k);
  if (k.includes('tarjeta') && (k.includes('pago') || k.includes('abono'))) return GRUPO_AHORRO_DEUDA;
  return GRUPO_DESEOS;
}

export function esGrupo503020Valido(g) {
  return GRUPO503020_TODOS.includes(g);
}

/**
 * Suma gastos del mes por grupo usando categorías normalizadas y mapa nombre→monto.
 */
export function gastosPorGrupo503020EnMes(gastosMesPorCategoria, categoriasNormalizadas) {
  const out = {
    [GRUPO_NECESIDADES]: 0,
    [GRUPO_DESEOS]: 0,
    [GRUPO_AHORRO_DEUDA]: 0,
  };
  const catMap = new Map((categoriasNormalizadas || []).map((c) => [c.nombre, c]));
  for (const [nombreCat, raw] of Object.entries(gastosMesPorCategoria || {})) {
    const monto = parseFloat(raw) || 0;
    if (monto <= 0) continue;
    const cfg = catMap.get(nombreCat);
    const grupo =
      cfg && esGrupo503020Valido(cfg.grupo503020)
        ? cfg.grupo503020
        : inferirGrupo503020DesdeNombre(nombreCat);
    out[grupo] += monto;
  }
  return out;
}

export function fraccionGastoSobreIngreso(gasto, ingreso) {
  if (!ingreso || ingreso <= 0) return 0;
  return Math.max(0, gasto / ingreso);
}
