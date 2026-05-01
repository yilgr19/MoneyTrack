/**
 * Tras fallos Ninja MAX_PATH o «Unable to delete directory» en Windows (OneDrive / daemon),
 * borra cachés nativas bajo el repo para forzar rebuild limpio.
 * No toca GRADLE_USER_HOME global (p. ej. C:\GradleMoneyTrack).
 */
const { spawnSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const mobileRoot = path.resolve(__dirname, '..');
const androidDir = path.join(mobileRoot, 'android');
const gradlew = process.platform === 'win32' ? 'gradlew.bat' : './gradlew';

/** Mismo criterio que `run-expo.cjs`: en Windows evita rutas largas bajo OneDrive. */
const gradleUserHomeDefault =
  process.platform === 'win32'
    ? 'C:\\GradleMoneyTrack'
    : path.join(mobileRoot, '.gradle-user-home');
const gradleUserHome = process.env.GRADLE_USER_HOME || gradleUserHomeDefault;

/** Carpetas `build` de plugins de Gradle dentro de node_modules (suelen bloquearse con OneDrive). */
const expoGradlePluginBuilds = [
  path.join(
    mobileRoot,
    'node_modules',
    'expo-modules-autolinking',
    'android',
    'expo-gradle-plugin',
    'expo-autolinking-plugin',
    'build',
  ),
  path.join(
    mobileRoot,
    'node_modules',
    'expo-modules-autolinking',
    'android',
    'expo-gradle-plugin',
    'expo-autolinking-plugin-shared',
    'build',
  ),
  path.join(
    mobileRoot,
    'node_modules',
    'expo-modules-autolinking',
    'android',
    'expo-gradle-plugin',
    'expo-autolinking-settings-plugin',
    'build',
  ),
  path.join(mobileRoot, 'node_modules', 'expo-modules-core', 'expo-module-gradle-plugin', 'build'),
];

const dirs = [
  path.join(mobileRoot, '.gradle-user-home'),
  path.join(mobileRoot, 'android', '.cxx'),
  path.join(mobileRoot, 'android', 'app', 'build'),
  path.join(mobileRoot, 'android', 'build'),
  path.join(mobileRoot, 'node_modules', 'react-native-screens', 'android', '.cxx'),
  path.join(mobileRoot, 'node_modules', 'react-native-screens', 'android', 'build'),
  ...expoGradlePluginBuilds,
];

function rmrf(d) {
  try {
    fs.rmSync(d, { recursive: true, force: true });
    console.log('Eliminado:', d);
  } catch (e) {
    console.warn('Omitido (no existe o en uso):', d, e.message || e);
  }
}

/**
 * Gradle «packageDebugResources»: Unable to delete …/packageDebugResources — en Windows + OneDrive
 * suelen quedar merged.dir/stripped.dir bloqueados. Solo tocamos paquetes nativos típicos (expo / RN),
 * no todo node_modules (evita miles de rutas falsas tipo @babel/.../android).
 */
function isReactNativeScope(scopeName) {
  return scopeName === '@expo' || scopeName.startsWith('@react-native');
}

function isReactNativeTopLevel(pkgName) {
  return (
    pkgName.startsWith('expo') ||
    pkgName.startsWith('react-native') ||
    pkgName === 'expo' ||
    pkgName === 'expo-router'
  );
}

function rmAllAndroidBuildUnderNodeModules() {
  const nm = path.join(mobileRoot, 'node_modules');
  if (!fs.existsSync(nm)) return;
  let entries;
  try {
    entries = fs.readdirSync(nm, { withFileTypes: true });
  } catch (e) {
    console.warn('[MoneyTrack] No se pudo leer node_modules:', e.message || e);
    return;
  }
  for (const ent of entries) {
    if (!ent.isDirectory() || ent.name.startsWith('.')) continue;
    const scope = path.join(nm, ent.name);
    if (ent.name.startsWith('@')) {
      if (!isReactNativeScope(ent.name)) continue;
      let subs;
      try {
        subs = fs.readdirSync(scope, { withFileTypes: true });
      } catch {
        continue;
      }
      for (const sub of subs) {
        if (!sub.isDirectory()) continue;
        const base = path.join(scope, sub.name, 'android');
        rmrf(path.join(base, 'build'));
        rmrf(path.join(base, '.cxx'));
      }
    } else {
      if (!isReactNativeTopLevel(ent.name)) continue;
      const base = path.join(scope, 'android');
      rmrf(path.join(base, 'build'));
      rmrf(path.join(base, '.cxx'));
    }
  }
}

console.log('[MoneyTrack] Deteniendo daemons de Gradle (libera archivos en node_modules)…');
if (fs.existsSync(path.join(androidDir, process.platform === 'win32' ? 'gradlew.bat' : 'gradlew'))) {
  const r = spawnSync(gradlew, ['--stop'], {
    cwd: androidDir,
    shell: true,
    stdio: 'inherit',
  });
  if (r.status !== 0) {
    console.warn('[MoneyTrack] gradlew --stop devolvió código', r.status, '(puede ignorarse si no había daemon).');
  }
} else {
  console.warn('[MoneyTrack] No hay android/gradlew; omite --stop.');
}

console.log('[MoneyTrack] Limpiando cachés nativas Android bajo el proyecto…');
dirs.forEach(rmrf);

console.log('[MoneyTrack] Borrando android/build (y .cxx) en todos los paquetes bajo node_modules…');
rmAllAndroidBuildUnderNodeModules();

/** Prefab/.so en caché de Gradle a veces quedan rotos con OneDrive («not a regular file»). */
const cachesRoot = path.join(gradleUserHome, 'caches');
try {
  if (fs.existsSync(cachesRoot)) {
    for (const ent of fs.readdirSync(cachesRoot, { withFileTypes: true })) {
      if (!ent.isDirectory() || !/^\d+\.\d+/.test(ent.name)) continue;
      const transforms = path.join(cachesRoot, ent.name, 'transforms');
      rmrf(transforms);
    }
  }
} catch (e) {
  console.warn('[MoneyTrack] No se pudo limpiar transforms de Gradle:', e.message || e);
}

console.log(
  'Listo. Compila siempre con: npm run run:android (así Metro y GRADLE_USER_HOME coinciden con run-expo).',
);
console.log(
  'Si Gradle usa otra carpeta (p. ej. C:\\GradleUserHome), define GRADLE_USER_HOME=C:\\GradleMoneyTrack o borra a mano …\\caches\\*\\transforms.',
);
console.log(
  'Si sigue «Unable to delete»: cierra Android Studio y Cursor un momento, pausa OneDrive en esta carpeta o mueve el repo fuera de OneDrive (p. ej. C:\\dev\\MoneyTrack).',
);
