/**
 * Tokens derivados de la paleta activa (tipografía, sombras, iconos semánticos).
 */
import { Platform, StyleSheet } from 'react-native';
import { paletteOriginal } from './colorPalettes';

/** Paleta por defecto para `colorIconoMetaDesdeNombre` fuera de ThemeProvider. */
export const METAS_ICON_PALETTE_DEFAULT = [
  '#fbbf24',
  '#fb7185',
  '#a78bfa',
  '#2dd4bf',
  '#4ade80',
  '#fb923c',
  '#c026d3',
  '#22d3ee',
  '#e879f9',
];

export function buildTypography(colors) {
  return {
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
}

export function buildShadows(colors) {
  return {
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
        shadowColor: colors.accentDeep,
        shadowOffset: { width: 0, height: 4 },
        shadowOpacity: 0.2,
        shadowRadius: 12,
      },
      android: { elevation: 3 },
      default: {},
    }),
  };
}

function rgbaFromHex(hex, alpha) {
  const h = String(hex || '').replace('#', '');
  if (h.length !== 6) return `rgba(255,255,255,${alpha})`;
  const r = parseInt(h.slice(0, 2), 16);
  const g = parseInt(h.slice(2, 4), 16);
  const b = parseInt(h.slice(4, 6), 16);
  return `rgba(${r},${g},${b},${alpha})`;
}

export function buildIconSemantic(colors) {
  const c = colors;
  return {
    moreMenu: {
      GastosFab: { fg: c.danger, bg: rgbaFromHex(c.danger, 0.14) },
      Ingresos: { fg: c.mint, bg: rgbaFromHex(c.mint, 0.16) },
      ExtractosTarjetas: { fg: c.chartBlue, bg: rgbaFromHex(c.chartBlue, 0.12) },
      InformesMensuales: { fg: c.mint, bg: rgbaFromHex(c.mint, 0.14) },
      Categorias: { fg: c.accentBright, bg: rgbaFromHex(c.accentDeep, 0.2) },
      MisBolsillos: { fg: c.chartBlue, bg: rgbaFromHex(c.chartBlue, 0.12) },
      Metas: { fg: c.accentGold, bg: rgbaFromHex(c.accentGold, 0.18) },
      PagosProgramados: { fg: c.orange, bg: rgbaFromHex(c.orange, 0.14) },
      Movimientos: { fg: c.chartBlue, bg: rgbaFromHex(c.chartBlue, 0.12) },
      Administrar: { fg: c.textMuted, bg: rgbaFromHex(c.textMuted, 0.12) },
      AsistenteCompras: { fg: c.chartBlue, bg: rgbaFromHex(c.chartBlue, 0.14) },
    },
    tabActive: {
      Inicio: c.accentBright,
      Gastos: c.danger,
      Saldo: c.mint,
      Mas: c.chartBlue,
    },
    metasIconPalette: [
      c.accentGold,
      c.danger,
      c.mint,
      c.chartBlue,
      c.success,
      c.orange,
      c.accentBright,
      c.warning,
      c.accent,
    ],
    saldoEdit: {
      'cash-outline': { fg: c.accentGold, bg: rgbaFromHex(c.accentGold, 0.12) },
      'wallet-outline': { fg: c.mint, bg: rgbaFromHex(c.mint, 0.12) },
      'business-outline': { fg: c.chartBlue, bg: rgbaFromHex(c.chartBlue, 0.12) },
      'card-outline': { fg: c.accent, bg: rgbaFromHex(c.accent, 0.12) },
      'apps-outline': { fg: c.danger, bg: rgbaFromHex(c.danger, 0.12) },
      'pie-chart-outline': { fg: c.mint, bg: rgbaFromHex(c.mint, 0.12) },
      'document-text-outline': { fg: c.textMuted, bg: rgbaFromHex(c.textMuted, 0.1) },
    },
  };
}

/** Color estable por nombre de icono de meta. */
export function colorIconoMetaDesdeNombre(ionName, palette) {
  const p = Array.isArray(palette) && palette.length ? palette : METAS_ICON_PALETTE_DEFAULT;
  const s = String(ionName || 'meta');
  let h = 0;
  for (let i = 0; i < s.length; i++) {
    h = (Math.imul(31, h) + s.charCodeAt(i)) | 0;
  }
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

export const screenPadding = {
  paddingHorizontal: spacing.lg,
  paddingBottom: spacing.xxl + 8,
};

export const TAB_BAR_SCROLL_PADDING = Platform.select({
  ios: 88,
  web: 64,
  android: 68,
  default: 68,
});

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

/**
 * Colores vivos para bolsillos (independientes del tema para distinguir filas).
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

/** Tipografía / sombras / iconSemantic estáticos (tema original) — compat. */
export const typography = buildTypography(paletteOriginal);
export const shadows = buildShadows(paletteOriginal);
export const iconSemantic = buildIconSemantic(paletteOriginal);
