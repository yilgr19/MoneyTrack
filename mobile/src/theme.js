import { Platform, StyleSheet } from 'react-native';

/** Paleta: fondo profundo, violetas y acentos fríos — aspecto “fintech” moderno */
export const colors = {
  bg: '#08060e',
  bgElevated: '#0e0c16',
  surface: 'rgba(22, 20, 35, 0.88)',
  surfaceSolid: '#161423',
  surfaceHighlight: 'rgba(139, 92, 246, 0.12)',
  stroke: 'rgba(167, 139, 250, 0.18)',
  strokeStrong: 'rgba(167, 139, 250, 0.35)',
  text: '#f8f7fc',
  textSecondary: 'rgba(216, 210, 245, 0.92)',
  textMuted: '#9088b0',
  textFaint: '#6b6580',
  accent: '#c4b5fd',
  accentBright: '#a78bfa',
  accentDeep: '#7c3aed',
  mint: '#5eead4',
  success: '#34d399',
  warning: '#fbbf24',
  danger: '#fb7185',
  orange: '#fdba74',
  /** Gradient stops (LinearGradient) */
  gradTop: '#1a1030',
  gradMid: '#0f0b18',
  gradBottom: '#08060e',
  btnFrom: '#8b5cf6',
  btnTo: '#6d28d9',
  tabBar: 'rgba(12, 10, 20, 0.94)',
  tabBorder: 'rgba(255,255,255,0.06)',
  headerBg: 'rgba(14, 12, 22, 0.97)',
  alertBg: 'rgba(180, 83, 9, 0.22)',
  alertBorder: 'rgba(251, 191, 36, 0.35)',
  barTrack: 'rgba(255,255,255,0.06)',
};

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
      shadowColor: '#7c3aed',
      shadowOffset: { width: 0, height: 4 },
      shadowOpacity: 0.15,
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
