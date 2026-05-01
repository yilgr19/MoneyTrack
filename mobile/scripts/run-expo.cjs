/**
 * Por defecto: depuración pensada para dispositivo físico (LAN + Metro puerto por defecto Expo/RN + dev client + --clear)
 * y `adb reverse` automático (8081/8082) si hay dispositivos — sin pasos manuales extra.
 *
 * Alternativas:
 *   `npm run start:localhost` — Metro 8081, sin --lan (emulador/USB clásico).
 *   `npm run start:tunnel` — sin depender de LAN ni reverse.
 *   MONEYTRACK_SKIP_ADB_REVERSE=1 — no ejecutar adb reverse al arrancar.
 *   MONEYTRACK_METRO_LOCALHOST=1 — mismo efecto que start:localhost para `npm start`.
 *
 * Fija el "Expo home" dentro del proyecto para evitar EPERM en C:\Users\<user>\.expo
 * (Expo CLI ignora EXPO_HOME; usa __UNSAFE_EXPO_HOME_DIRECTORY — ver UserSettings.js)
 */
const { spawn, spawnSync } = require('child_process');
const fs = require('fs');
const path = require('path');
const { purgeStaleExpoSharingDirs } = require('./purge-stale-expo-autolinking.cjs');

const expoHome = path.resolve(__dirname, '..', '.expo-home');
try {
  fs.mkdirSync(expoHome, { recursive: true });
} catch (_) {
  /* si falla, Expo intentará el home por defecto */
}
process.env.__UNSAFE_EXPO_HOME_DIRECTORY = expoHome;

/** Evita avisos/fallos en tareas Gradle que esperan bundle de Expo (`NODE_ENV`). */
if (process.env.NODE_ENV == null || process.env.NODE_ENV === '') {
  process.env.NODE_ENV = 'development';
}

/**
 * Cache de Gradle: en Windows forzamos ruta corta (C:\GradleMoneyTrack) aunque exista GRADLE_USER_HOME
 * global apuntando a …\mobile\.gradle-user-home (Ninja / RN nativo: MAX_PATH 260).
 * MONEYTRACK_SKIP_SHORT_GRADLE_HOME=1 restaura el comportamiento anterior.
 */
let gradleUserHome;
if (process.platform === 'win32') {
  gradleUserHome =
    process.env.MONEYTRACK_SKIP_SHORT_GRADLE_HOME === '1'
      ? process.env.GRADLE_USER_HOME || path.resolve(__dirname, '..', '.gradle-user-home')
      : 'C:\\GradleMoneyTrack';
} else {
  gradleUserHome =
    process.env.GRADLE_USER_HOME ||
    path.resolve(__dirname, '..', '.gradle-user-home');
}
try {
  fs.mkdirSync(gradleUserHome, { recursive: true });
} catch (_) {}
process.env.GRADLE_USER_HOME = gradleUserHome;
if (process.env.EXPO_NO_TELEMETRY == null) {
  process.env.EXPO_NO_TELEMETRY = '1';
}

/** Evita crash de Metro (lstat UNKNOWN) por carpetas build bajo expo-constants en Windows/OneDrive. */
if (process.platform === 'win32') {
  const poison = path.resolve(__dirname, '..', 'node_modules', 'expo-constants', 'android', 'bin');
  try {
    fs.rmSync(poison, { recursive: true, force: true });
  } catch (_) {
    /* ignorar */
  }
}

const passArgs = process.argv.slice(2);

/**
 * - Wi‑Fi / LAN: --dev-client + --lan; Metro en el mismo puerto que espera el binario debug (8081).
 *   Antes forzábamos 8082: la app + adb reverse apuntan a 8081 → pantalla en blanco / "could not connect".
 * - USB / emulador: --localhost (mismo 8081 por defecto de Expo).
 * Antes, `--localhost` salía antes de añadir --dev-client y el build de desarrollo no enlazaba bien con Metro.
 */
function injectWirelessDefaults(argv) {
  if (argv[0] !== 'start') return argv;
  if (argv.includes('--web')) return argv;
  if (argv.includes('--offline')) return argv;
  if (argv.includes('--tunnel')) return argv;

  const out = [...argv];
  const useLocalhost =
    out.includes('--localhost') || process.env.MONEYTRACK_METRO_LOCALHOST === '1';
  if (process.env.MONEYTRACK_METRO_LOCALHOST === '1' && !out.includes('--localhost')) {
    out.push('--localhost');
  }

  if (!out.includes('--dev-client')) out.push('--dev-client');
  if (!out.includes('--clear')) out.push('--clear');

  if (useLocalhost) {
    return out;
  }

  if (!out.includes('--lan')) out.push('--lan');
  /* No fijar --port: Expo usa 8081 por defecto, alineado con RN debug y adb reverse. */
  return out;
}

function shouldTryAdbReverse(argv) {
  if (process.env.MONEYTRACK_SKIP_ADB_REVERSE === '1') return false;
  const cmd = argv[0];
  if (cmd === 'run:android') return true;
  if (cmd !== 'start') return false;
  if (argv.includes('--web') || argv.includes('--offline') || argv.includes('--tunnel')) return false;
  return true;
}

const finalArgs = injectWirelessDefaults(passArgs);

/** Evita crash ENOENT de Metro por carpetas `.expo-sharing-*` colgantes bajo node_modules. */
purgeStaleExpoSharingDirs({ quiet: true });

if (shouldTryAdbReverse(passArgs)) {
  spawnSync(process.execPath, [path.join(__dirname, 'adb-reverse.cjs'), '--soft'], {
    cwd: path.resolve(__dirname, '..'),
    stdio: 'inherit',
    env: process.env,
  });
}

const child = spawn('npx', ['expo', ...finalArgs], {
  stdio: 'inherit',
  shell: true,
  env: process.env,
  cwd: path.resolve(__dirname, '..'),
});

child.on('exit', (code, signal) => {
  if (signal) process.exit(1);
  process.exit(code ?? 0);
});
