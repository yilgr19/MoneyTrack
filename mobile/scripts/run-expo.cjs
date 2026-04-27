/**
 * Fija el "Expo home" dentro del proyecto para evitar EPERM en C:\Users\<user>\.expo
 * (Expo CLI ignora EXPO_HOME; usa __UNSAFE_EXPO_HOME_DIRECTORY — ver UserSettings.js)
 */
const { spawn } = require('child_process');
const path = require('path');

const expoHome = path.resolve(__dirname, '..', '.expo-home');
process.env.__UNSAFE_EXPO_HOME_DIRECTORY = expoHome;
if (process.env.EXPO_NO_TELEMETRY == null) {
  process.env.EXPO_NO_TELEMETRY = '1';
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
