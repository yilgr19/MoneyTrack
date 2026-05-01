/**
 * Por defecto: depuración pensada para dispositivo físico (LAN + Metro 8082 + dev client + --clear)
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
 * Cache de Gradle: en Windows ignoramos GRADLE_USER_HOME antigua (ej. …\mobile\.gradle-user-home) para no superar MAX_PATH en Ninja/RN.
 */
let gradleUserHome;
if (process.platform === 'win32') {
  gradleUserHome = 'C:\\GradleMoneyTrack';
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

function hasPortArg(argv) {
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--port' || a.startsWith('--port=')) return true;
  }
  return false;
}

/** LAN + 8082 + dev-client: el teléfono en la misma WiFi alcanza la IP del PC sin túnel. */
function injectWirelessDefaults(argv) {
  if (argv[0] !== 'start') return argv;
  if (argv.includes('--web')) return argv;
  if (argv.includes('--offline')) return argv;
  if (argv.includes('--tunnel')) return argv;
  if (argv.includes('--localhost')) return argv;
  if (process.env.MONEYTRACK_METRO_LOCALHOST === '1') return argv;

  const out = [...argv];
  if (!out.includes('--dev-client')) out.push('--dev-client');
  if (!out.includes('--lan')) out.push('--lan');
  if (!out.includes('--clear')) out.push('--clear');
  if (!hasPortArg(out)) out.push('--port', '8082');
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
