/**
 * Prefijos para títulos de notificaciones (barra del sistema y campana in-app).
 */
export function tituloNotifConNombre(state, titulo) {
  const raw = String(titulo || '').trim();
  if (!raw) return '';
  const n = String(state?.nombreUsuario || '').trim();
  if (!n) return raw;
  return `${n} — ${raw}`;
}
