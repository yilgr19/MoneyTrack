/**
 * Trazo de anillo (donut) en SVG, ángulos en grados desde arriba en sentido horario.
 * @param {number} cx
 * @param {number} cy
 * @param {number} rIn
 * @param {number} rOut
 * @param {number} a0
 * @param {number} a1
 * @returns {string}
 */
export function ringSlicePath(cx, cy, rIn, rOut, a0, a1) {
  if (a1 - a0 <= 0) return '';
  /** Casi vuelta completa: evita singularidad del arco 360° */
  if (a1 - a0 >= 359.5) {
    a1 = a0 + 359.98;
  }
  const p = (r, a) => {
    const t = ((a - 90) * Math.PI) / 180;
    return { x: cx + r * Math.cos(t), y: cy + r * Math.sin(t) };
  };
  const o0 = p(rOut, a0);
  const o1 = p(rOut, a1);
  const i1 = p(rIn, a1);
  const i0 = p(rIn, a0);
  const large = a1 - a0 > 180 ? 1 : 0;
  return `M ${o0.x} ${o0.y} A ${rOut} ${rOut} 0 ${large} 1 ${o1.x} ${o1.y} L ${i1.x} ${i1.y} A ${rIn} ${rIn} 0 ${large} 0 ${i0.x} ${i0.y} Z`;
}

/** Anillo 360° en dos trazos: un solo path 0→360 degenera (inicio = fin en el arco). */
export function fullDonutPaths(cx, cy, rIn, rOut) {
  const a = ringSlicePath(cx, cy, rIn, rOut, 0, 180);
  const b = ringSlicePath(cx, cy, rIn, rOut, 180, 360);
  return [a, b].filter(Boolean);
}
