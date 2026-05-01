import React from 'react';
import Svg, { Line, Rect } from 'react-native-svg';
import { colors } from '../theme';

/**
 * Marca minimalista MoneyTrack: tendencia en barras + línea base (oro).
 * Misma idea visual que los PNG generados para icono / splash — serio, limpio, sin ilustración recargada.
 */
export function MoneyTrackMark({ size = 120 }) {
  return (
    <Svg width={size} height={size} viewBox="0 0 100 100">
      <Line
        x1="20"
        y1="71"
        x2="80"
        y2="71"
        stroke={colors.accentGold}
        strokeWidth="1.25"
        strokeOpacity={0.92}
        strokeLinecap="round"
      />
      <Rect x="23" y="50" width="11" height="21" rx="2.2" fill={colors.mint} fillOpacity={0.96} />
      <Rect x="44.5" y="40" width="11" height="31" rx="2.2" fill={colors.chartBlue} fillOpacity={0.96} />
      <Rect x="66" y="30" width="11" height="41" rx="2.2" fill={colors.mint} fillOpacity={0.96} />
    </Svg>
  );
}
