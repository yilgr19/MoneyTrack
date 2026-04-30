import { Platform, StyleSheet } from 'react-native';

/**
 * Paleta alineada al logo Money Track — púrpura profundo, lavanda, menta y oro.
 * Referencias: primary #4B246C / #3E1F5A, mint #7DC191, gold #D9B44A, lavanda #C7C3E3, azul gráfico #A7D8DE.
 */
export const colors = {
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
  /** Acento oro del logo — destacados y alertas informativas */
  accentGold: '#D9B44A',
  mint: '#7DC191',
  success: '#6BAF82',
  warning: '#D9B44A',
  danger: '#c77b88',
  orange: '#D9B44A',
  /** Datos / gráficos secundarios */
  chartBlue: '#A7D8DE',
  /** Gradient stops (LinearGradient) */
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

/**
 * Colores vivos y distinguibles para bolsillos (no apagados).
 * Elegidos para no solaparse con la paleta por defecto de categorías.
 */
export const COLORES_BOLSILLO = [
  '#2dd4bf',
  '#4ade80',
  '#a78bfa',
  '#fbbf24',
  '#fb7185',
  '#38bdf8',
  '#34d399',
  '#fb923c',
  '#c026d3',
  '#8b5cf6',
  '#22d3ee',
  '#e879f9',
];

/**
 * Acentos de icono por contexto: más significado, menos “todo lavanda”.
 * `fg` = color del glifo; `bg` = mancha suave bajo el icono (menú Más).
 */
export const iconSemantic = {
  moreMenu: {
    /** Tarjeta «Registrar gasto» en el panel del FAB (+) */
    GastosFab: { fg: '#fb7185', bg: 'rgba(251, 113, 133, 0.14)' },
    Ingresos: { fg: '#34d399', bg: 'rgba(52, 211, 153, 0.16)' },
    ExtractosTarjetas: { fg: colors.chartBlue, bg: 'rgba(167, 216, 222, 0.12)' },
    InformesMensuales: { fg: '#34d399', bg: 'rgba(52, 211, 153, 0.14)' },
    Categorias: { fg: '#a78bfa', bg: 'rgba(167, 139, 250, 0.14)' },
    MisBolsillos: { fg: '#2dd4bf', bg: 'rgba(45, 212, 191, 0.12)' },
    Metas: { fg: colors.accentGold, bg: 'rgba(217, 180, 74, 0.18)' },
    PagosProgramados: { fg: '#fb923c', bg: 'rgba(251, 146, 60, 0.14)' },
    Movimientos: { fg: '#22d3ee', bg: 'rgba(34, 211, 238, 0.12)' },
    Administrar: { fg: '#94a3b8', bg: 'rgba(148, 163, 184, 0.12)' },
    AsistenteCompras: { fg: '#e879f9', bg: 'rgba(232, 121, 249, 0.14)' },
  },
  /** Pestaña activa (la inactiva sigue en gris de React Navigation) */
  tabActive: {
    Inicio: '#c4b5fd',
    Gastos: '#fb7185',
    Saldo: '#4ade80',
    Mas: '#7dd3fc',
  },
  /** Metas — giro de colores vivos (Inicio, Metas, carrusel) */
  metasIconPalette: [
    '#fbbf24',
    '#fb7185',
    '#a78bfa',
    '#2dd4bf',
    '#4ade80',
    '#fb923c',
    '#c026d3',
    '#22d3ee',
    '#e879f9',
  ],
  /** Saldo inicial — glifo + fondo del círculo */
  saldoEdit: {
    'cash-outline': { fg: '#facc15', bg: 'rgba(250, 204, 21, 0.12)' },
    'wallet-outline': { fg: '#4ade80', bg: 'rgba(74, 222, 128, 0.12)' },
    'business-outline': { fg: '#60a5fa', bg: 'rgba(96, 165, 250, 0.12)' },
    'card-outline': { fg: '#c084fc', bg: 'rgba(192, 132, 252, 0.12)' },
    'apps-outline': { fg: '#f472b6', bg: 'rgba(244, 114, 182, 0.12)' },
    'pie-chart-outline': { fg: '#34d399', bg: 'rgba(52, 211, 153, 0.12)' },
    'document-text-outline': { fg: '#a8a29e', bg: 'rgba(168, 162, 158, 0.1)' },
  },
};

/** Color estable por nombre de icono de meta (Inicio, Metas, carrusel). */
export function colorIconoMetaDesdeNombre(ionName) {
  const s = String(ionName || 'meta');
  let h = 0;
  for (let i = 0; i < s.length; i++) {
    h = (Math.imul(31, h) + s.charCodeAt(i)) | 0;
  }
  const p = iconSemantic.metasIconPalette;
  return p[Math.abs(h) % p.length];
}

export const spacing = {
  xs: 6,
  sm: 10,
  md: 16,
  lg: 22,
  xl: 28,
  xxl: 36,
};

export const radii = {
  sm: 10,
  md: 14,
  lg: 18,
  xl: 24,
  pill: 999,
};

/** Sombras suaves — iOS + Android */
export const shadows = {
  card: Platform.select({
    ios: {
      shadowColor: '#000',
      shadowOffset: { width: 0, height: 12 },
      shadowOpacity: 0.4,
      shadowRadius: 24,
    },
    android: { elevation: 8 },
    default: {},
  }),
  soft: Platform.select({
    ios: {
      shadowColor: '#4B246C',
      shadowOffset: { width: 0, height: 4 },
      shadowOpacity: 0.2,
      shadowRadius: 12,
    },
    android: { elevation: 3 },
    default: {},
  }),
};

export const typography = {
  hero: {
    fontSize: 30,
    fontWeight: '700',
    color: colors.text,
    letterSpacing: -0.8,
  },
  title: {
    fontSize: 22,
    fontWeight: '700',
    color: colors.text,
    letterSpacing: -0.4,
  },
  subtitle: {
    fontSize: 15,
    fontWeight: '500',
    color: colors.textMuted,
    letterSpacing: 0.2,
  },
  label: {
    fontSize: 12,
    fontWeight: '600',
    color: colors.accent,
    letterSpacing: 1.2,
    textTransform: 'uppercase',
  },
  body: { fontSize: 15, color: colors.textSecondary, lineHeight: 22 },
  small: { fontSize: 13, color: colors.textMuted, lineHeight: 18 },
  monoAmount: {
    fontSize: 17,
    fontWeight: '600',
    color: colors.text,
    fontVariant: ['tabular-nums'],
  },
};

export const screenPadding = {
  paddingHorizontal: spacing.lg,
  paddingBottom: spacing.xxl + 8,
};

/** Altura efectiva de la barra de pestañas (debe coincidir con AppNavigator) para padding del scroll */
export const TAB_BAR_SCROLL_PADDING = Platform.select({
  ios: 88,
  web: 64,
  android: 68,
  default: 68,
});

/** Filas etiqueta/valor: evita que el texto largo empuje montos fuera de pantalla */
export const layoutStyles = StyleSheet.create({
  rowBetween: {
    flexDirection: 'row',
    justifyContent: 'space-between',
    alignItems: 'flex-start',
    marginTop: spacing.sm,
    gap: spacing.sm,
  },
  rowLabel: {
    flex: 1,
    flexShrink: 1,
    minWidth: 0,
    paddingRight: spacing.sm,
  },
  rowValue: {
    flexShrink: 0,
    textAlign: 'right',
    maxWidth: '58%',
  },
  statRow: {
    flexDirection: 'row',
    justifyContent: 'space-between',
    alignItems: 'flex-start',
    marginTop: spacing.sm,
    gap: spacing.sm,
  },
  statLabel: {
    flex: 1,
    flexShrink: 1,
    minWidth: 0,
    paddingRight: spacing.sm,
  },
  statValue: {
    flexShrink: 0,
    textAlign: 'right',
    maxWidth: '58%',
  },
});
