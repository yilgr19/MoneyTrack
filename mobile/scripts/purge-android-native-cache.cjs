/**
 * Tras fallos Ninja MAX_PATH, borra cachés nativas bajo el repo para forzar rebuild limpio.
 * No toca GRADLE_USER_HOME global (p. ej. C:\GradleMoneyTrack).
 */
const fs = require('fs');
const path = require('path');

const mobileRoot = path.resolve(__dirname, '..');
const dirs = [
  path.join(mobileRoot, '.gradle-user-home'),
  path.join(mobileRoot, 'android', '.cxx'),
  path.join(mobileRoot, 'android', 'app', 'build'),
  path.join(mobileRoot, 'android', 'build'),
  path.join(mobileRoot, 'node_modules', 'react-native-screens', 'android', '.cxx'),
];

function rmrf(d) {
  try {
    fs.rmSync(d, { recursive: true, force: true });
    console.log('Eliminado:', d);
  } catch (e) {
    console.warn('Omitido (no existe o en uso):', d, e.message || e);
  }
}

console.log('[MoneyTrack] Limpiando cachés nativas Android bajo el proyecto…');
dirs.forEach(rmrf);
console.log('Listo. Vuelve a compilar con: npm run run:android');
