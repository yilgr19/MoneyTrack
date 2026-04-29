/**
 * Puertos Metro: `npm start` → 8081. `npm run start:wifi` → 8082 + --lan (recomendado si el móvil
 * no usa localhost hacia el PC).
 *
 * Depuración inalámbrica (ADB por Wi‑Fi): tras `adb pair` / `adb connect IP:PUERTO`, el túnel
 * `adb reverse` sigue funcionando igual que con USB → `npm run android:adb-reverse` y luego
 * `npm start` (8081) suele bastar. Si no carga el bundle, usa `npm run start:wifi` y asegura
 * que el firewall de Windows permita TCP 8082 en la red privada; el dispositivo debe alcanzar
 * la IP LAN del PC (misma WiFi, sin aislamiento de cliente/AP).
 *
 * Fija el "Expo home" dentro del proyecto para evitar EPERM en C:\Users\<user>\.expo
 * (Expo CLI ignora EXPO_HOME; usa __UNSAFE_EXPO_HOME_DIRECTORY — ver UserSettings.js)
 */
const { spawn } = require('child_process');
const fs = require('fs');
const path = require('path');

const expoHome = path.resolve(__dirname, '..', '.expo-home');
try {
  fs.mkdirSync(expoHome, { recursive: true });
} catch (_) {
  /* si falla, Expo intentará el home por defecto */
}
process.env.__UNSAFE_EXPO_HOME_DIRECTORY = expoHome;

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
const child = spawn('npx', ['expo', ...passArgs], {
  stdio: 'inherit',
  shell: true,
  env: process.env,
  cwd: path.resolve(__dirname, '..'),
});

child.on('exit', (code, signal) => {
  if (signal) process.exit(1);
  process.exit(code ?? 0);
});
