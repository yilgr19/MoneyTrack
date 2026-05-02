/**
 * Paletas intercambiables — mismas claves que `colors` en theme.js (layout idéntico, solo tonos).
 */

export const THEME_ID_ORIGINAL = 'original';
export const THEME_ID_ROSA = 'rosa';
export const THEME_ID_VINO = 'vino';
export const THEME_ID_MILITAR = 'militar';

export const OPCIONES_TEMA_APP = [
  { id: THEME_ID_ORIGINAL, label: 'Original', subtitle: 'Púrpura MoneyTrack', emoji: '💜' },
  { id: THEME_ID_ROSA, label: 'Rosa', subtitle: 'Suave y luminoso', emoji: '🌸' },
  { id: THEME_ID_VINO, label: 'Vino', subtitle: 'Elegante y cálido', emoji: '🍷' },
  { id: THEME_ID_MILITAR, label: 'Militar', subtitle: 'Verde moderno', emoji: '🌿' },
];

const IDS = new Set(OPCIONES_TEMA_APP.map((o) => o.id));

export function normalizeTemaId(raw) {
  const s = String(raw || '').trim().toLowerCase();
  if (IDS.has(s)) return s;
  return THEME_ID_ORIGINAL;
}

/** Púrpura profundo — referencia al logo. */
export const paletteOriginal = {
  bg: '#0c0812',
  bgElevated: '#140e1c',
  surface: 'rgba(32, 26, 44, 0.92)',
  surfaceSolid: '#1c1526',
  surfaceHighlight: 'rgba(75, 36, 108, 0.2)',
  stroke: 'rgba(199, 195, 227, 0.2)',
  strokeStrong: 'rgba(199, 195, 227, 0.38)',
  text: '#f7f5fb',
  textSecondary: 'rgba(232, 228, 245, 0.9)',
  textMuted: '#9b94b8',
  textFaint: '#6d6685',
  accent: '#C7C3E3',
  accentBright: '#e4e0f5',
  accentDeep: '#4B246C',
  accentGold: '#D9B44A',
  mint: '#7DC191',
  success: '#6BAF82',
  warning: '#D9B44A',
  danger: '#c77b88',
  orange: '#D9B44A',
  chartBlue: '#A7D8DE',
  gradTop: '#2a1f3d',
  gradMid: '#150f1f',
  gradBottom: '#0c0812',
  btnFrom: '#5a2f7d',
  btnTo: '#3E1F5A',
  tabBar: 'rgba(12, 8, 18, 0.94)',
  tabBorder: 'rgba(199, 195, 227, 0.08)',
  headerBg: 'rgba(18, 12, 26, 0.97)',
  alertBg: 'rgba(217, 180, 74, 0.14)',
  alertBorder: 'rgba(217, 180, 74, 0.38)',
  barTrack: 'rgba(167, 216, 222, 0.12)',
};

/** Rosas y cremas — estilo app “chic” sin perder contraste. */
export const paletteRosa = {
  bg: '#100a10',
  bgElevated: '#180f16',
  surface: 'rgba(60, 32, 52, 0.88)',
  surfaceSolid: '#22141c',
  surfaceHighlight: 'rgba(244, 114, 182, 0.18)',
  stroke: 'rgba(252, 211, 232, 0.22)',
  strokeStrong: 'rgba(252, 211, 232, 0.4)',
  text: '#fff5fb',
  textSecondary: 'rgba(255, 232, 242, 0.92)',
  textMuted: '#c4a3b8',
  textFaint: '#8f7488',
  accent: '#fbcfe8',
  accentBright: '#fdf2f8',
  accentDeep: '#be185d',
  accentGold: '#f9a8d4',
  mint: '#f472b6',
  success: '#86efac',
  warning: '#fbbf24',
  danger: '#fb7185',
  orange: '#fdba74',
  chartBlue: '#e879f9',
  gradTop: '#301828',
  gradMid: '#1a0f16',
  gradBottom: '#100a10',
  btnFrom: '#db2777',
  btnTo: '#831843',
  tabBar: 'rgba(16, 10, 16, 0.94)',
  tabBorder: 'rgba(251, 207, 232, 0.1)',
  headerBg: 'rgba(28, 14, 24, 0.97)',
  alertBg: 'rgba(244, 114, 182, 0.14)',
  alertBorder: 'rgba(244, 114, 182, 0.38)',
  barTrack: 'rgba(232, 121, 249, 0.12)',
};

/** Burdeos, oro viejo y sombras cálidas. */
export const paletteVino = {
  bg: '#0a0708',
  bgElevated: '#120c0e',
  surface: 'rgba(48, 22, 28, 0.9)',
  surfaceSolid: '#1a1014',
  surfaceHighlight: 'rgba(114, 47, 55, 0.28)',
  stroke: 'rgba(201, 168, 130, 0.22)',
  strokeStrong: 'rgba(201, 168, 130, 0.42)',
  text: '#faf6f3',
  textSecondary: 'rgba(245, 235, 228, 0.9)',
  textMuted: '#b8a090',
  textFaint: '#7d6c62',
  accent: '#e8d5c4',
  accentBright: '#f5ebe3',
  accentDeep: '#722f37',
  accentGold: '#c9a882',
  mint: '#8da396',
  success: '#6d9b7a',
  warning: '#c9a882',
  danger: '#c97b7b',
  orange: '#c9a882',
  chartBlue: '#a8c4ce',
  gradTop: '#2a1418',
  gradMid: '#150c10',
  gradBottom: '#0a0708',
  btnFrom: '#8b2942',
  btnTo: '#3d141c',
  tabBar: 'rgba(10, 7, 8, 0.94)',
  tabBorder: 'rgba(201, 168, 130, 0.1)',
  headerBg: 'rgba(22, 12, 16, 0.97)',
  alertBg: 'rgba(201, 168, 130, 0.16)',
  alertBorder: 'rgba(201, 168, 130, 0.38)',
  barTrack: 'rgba(168, 196, 206, 0.12)',
};

/** Verde oliva / lima — moderno y sobrio. */
export const paletteMilitar = {
  bg: '#070a08',
  bgElevated: '#0d120e',
  surface: 'rgba(28, 38, 30, 0.92)',
  surfaceSolid: '#121a14',
  surfaceHighlight: 'rgba(132, 204, 22, 0.14)',
  stroke: 'rgba(180, 210, 170, 0.2)',
  strokeStrong: 'rgba(180, 210, 170, 0.38)',
  text: '#f2faf0',
  textSecondary: 'rgba(228, 240, 224, 0.9)',
  textMuted: '#8faa8c',
  textFaint: '#5c6e58',
  accent: '#d4e8cc',
  accentBright: '#ecf8e8',
  accentDeep: '#3f6212',
  accentGold: '#ca8a04',
  mint: '#84cc16',
  success: '#65a30d',
  warning: '#ca8a04',
  danger: '#f87171',
  orange: '#eab308',
  chartBlue: '#67e8f9',
  gradTop: '#1a2218',
  gradMid: '#0e140f',
  gradBottom: '#070a08',
  btnFrom: '#4d7c0f',
  btnTo: '#1a2e05',
  tabBar: 'rgba(7, 10, 8, 0.94)',
  tabBorder: 'rgba(180, 210, 170, 0.08)',
  headerBg: 'rgba(12, 18, 14, 0.97)',
  alertBg: 'rgba(202, 138, 4, 0.14)',
  alertBorder: 'rgba(202, 138, 4, 0.36)',
  barTrack: 'rgba(103, 232, 249, 0.12)',
};

export function getColorsForTemaId(id) {
  switch (normalizeTemaId(id)) {
    case THEME_ID_ROSA:
      return paletteRosa;
    case THEME_ID_VINO:
      return paletteVino;
    case THEME_ID_MILITAR:
      return paletteMilitar;
    default:
      return paletteOriginal;
  }
}
