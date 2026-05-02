/**
 * Paleta y tokens “activos” sincronizados desde ThemeProvider para que componentes
 * que aún importan `colors` / `typography` / `shadows` desde `../theme` reflejen el tema elegido.
 */
import { paletteOriginal } from './colorPalettes';
import { buildTypography, buildShadows, buildIconSemantic } from './themeTokens';

let activeColors = paletteOriginal;
let activeTypography = buildTypography(paletteOriginal);
let activeShadows = buildShadows(paletteOriginal);
let activeIconSemantic = buildIconSemantic(paletteOriginal);

export function setActiveThemeSnapshot(colors) {
  activeColors = colors;
  activeTypography = buildTypography(colors);
  activeShadows = buildShadows(colors);
  activeIconSemantic = buildIconSemantic(colors);
}

export function getActiveColors() {
  return activeColors;
}
export function getActiveTypography() {
  return activeTypography;
}
export function getActiveShadows() {
  return activeShadows;
}
export function getActiveIconSemantic() {
  return activeIconSemantic;
}

/** Lectura en cada acceso a la propiedad (compatible con StyleSheet y JSX). */
export const colorsProxy = new Proxy(
  {},
  {
    get(_, prop) {
      return activeColors[prop];
    },
  }
);

export const typographyProxy = new Proxy(
  {},
  {
    get(_, prop) {
      return activeTypography[prop];
    },
  }
);

export const shadowsProxy = new Proxy(
  {},
  {
    get(_, prop) {
      return activeShadows[prop];
    },
  }
);

export const iconSemanticProxy = new Proxy(
  {},
  {
    get(_, prop) {
      return activeIconSemantic[prop];
    },
  }
);
