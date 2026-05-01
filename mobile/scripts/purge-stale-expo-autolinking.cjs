/**
 * Elimina carpetas huérfanas de autolinking de Expo bajo node_modules (p. ej. `.expo-sharing-*`,
 * `.expo-camera-*`). Si Metro las vigila y desaparecen (npm, OneDrive), Windows puede lanzar ENOENT en `watch`.
 */
const fs = require('fs');
const path = require('path');

const STALE_AUTOLINK_PREFIXES = ['.expo-sharing-', '.expo-camera-'];

/**
 * @param {{ quiet?: boolean }} [options] — si `quiet`, solo imprime si hubo borrados.
 * @returns {number} carpetas eliminadas
 */
function purgeStaleExpoAutolinkingDirs(options = {}) {
  const { quiet } = options;
  const nodeModules = path.join(__dirname, '..', 'node_modules');
  if (!fs.existsSync(nodeModules)) {
    return 0;
  }
  let removed = 0;
  for (const name of fs.readdirSync(nodeModules, { withFileTypes: true })) {
    if (!name.isDirectory()) continue;
    const hit = STALE_AUTOLINK_PREFIXES.some((p) => name.name.startsWith(p));
    if (!hit) continue;
    const full = path.join(nodeModules, name.name);
    try {
      fs.rmSync(full, { recursive: true, force: true });
      removed += 1;
      if (!quiet) {
        console.log('[MoneyTrack] Eliminado resto de autolinking:', name.name);
      }
    } catch (e) {
      console.warn('[MoneyTrack] No se pudo borrar', name.name, e.message);
    }
  }
  if (!quiet && removed === 0) {
    console.log('[MoneyTrack] No había carpetas .expo-* de autolinking que limpiar.');
  }
  if (quiet && removed > 0) {
    console.log(`[MoneyTrack] Limpieza: ${removed} carpeta(s) de autolinking Expo eliminada(s) antes de Metro.`);
  }
  return removed;
}

/** @deprecated usar purgeStaleExpoAutolinkingDirs */
function purgeStaleExpoSharingDirs(options) {
  return purgeStaleExpoAutolinkingDirs(options);
}

module.exports = { purgeStaleExpoAutolinkingDirs, purgeStaleExpoSharingDirs };

if (require.main === module) {
  purgeStaleExpoAutolinkingDirs({ quiet: false });
}
