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
