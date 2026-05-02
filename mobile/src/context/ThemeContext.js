import React, { createContext, useContext, useMemo } from 'react';
import { normalizeTemaId, getColorsForTemaId } from '../theme/colorPalettes';
import { buildTypography, buildShadows, buildIconSemantic, colorIconoMetaDesdeNombre } from '../theme/themeTokens';
import { setActiveThemeSnapshot } from '../theme/themeActive';
import { spacing, radii, screenPadding, TAB_BAR_SCROLL_PADDING, layoutStyles, COLORES_BOLSILLO } from '../theme/themeTokens';

const ThemeContext = createContext(null);

export function ThemeProvider({ temaId, children }) {
  const value = useMemo(() => {
    const id = normalizeTemaId(temaId);
    const colors = getColorsForTemaId(id);
    setActiveThemeSnapshot(colors);
    const typography = buildTypography(colors);
    const shadows = buildShadows(colors);
    const iconSemantic = buildIconSemantic(colors);
    const colorIconoMeta = (ionName) => colorIconoMetaDesdeNombre(ionName, iconSemantic.metasIconPalette);
    return {
      temaId: id,
      colors,
      typography,
      shadows,
      iconSemantic,
      colorIconoMetaDesdeNombre: colorIconoMeta,
      spacing,
      radii,
      screenPadding,
      TAB_BAR_SCROLL_PADDING,
      layoutStyles,
      COLORES_BOLSILLO,
    };
  }, [temaId]);

  return <ThemeContext.Provider value={value}>{children}</ThemeContext.Provider>;
}

export function useTheme() {
  const ctx = useContext(ThemeContext);
  if (!ctx) {
    throw new Error('useTheme debe usarse dentro de ThemeProvider');
  }
  return ctx;
}
