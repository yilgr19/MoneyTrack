/**
 * Tema MoneyTrack — paletas en `theme/colorPalettes.js`, tokens en `theme/themeTokens.js`.
 * `colors` / `typography` / `shadows` / `iconSemantic` son proxies sincronizados con ThemeProvider
 * (acceso en render o dentro de useMemo que dependa del tema). Preferir `useTheme()` en código nuevo.
 */
export {
  colorsProxy as colors,
  typographyProxy as typography,
  shadowsProxy as shadows,
  iconSemanticProxy as iconSemantic,
} from './theme/themeActive';
export {
  THEME_ID_ORIGINAL,
  THEME_ID_ROSA,
  THEME_ID_VINO,
  THEME_ID_MILITAR,
  OPCIONES_TEMA_APP,
  normalizeTemaId,
  getColorsForTemaId,
  paletteOriginal,
  paletteRosa,
  paletteVino,
  paletteMilitar,
} from './theme/colorPalettes';
export {
  buildTypography,
  buildShadows,
  buildIconSemantic,
  colorIconoMetaDesdeNombre,
  METAS_ICON_PALETTE_DEFAULT,
  spacing,
  radii,
  screenPadding,
  TAB_BAR_SCROLL_PADDING,
  layoutStyles,
  COLORES_BOLSILLO,
} from './theme/themeTokens';
