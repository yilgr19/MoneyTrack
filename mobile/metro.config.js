// https://docs.expo.dev/guides/customizing-metro
const { getDefaultConfig } = require('expo/metro-config');

/** @type {import('expo/metro-config').MetroConfig} */
const config = getDefaultConfig(__dirname);

/**
 * En Windows + OneDrive, Gradle puede dejar bajo node_modules rutas raras en
 * expo-constants/android/bin; el FallbackWatcher de Metro hace lstat y revienta.
 */
const poison = [/node_modules[/\\]expo-constants[/\\]android[/\\]bin[/\\].*/];

/**
 * `lightningcss` declara optionalDependencies para todas las plataformas; en Windows solo
 * se instala win32-*. Metro a veces intenta vigilar carpetas inexistentes (symlink roto /
 * OneDrive) y falla con ENOENT en p. ej. lightningcss-linux-arm-gnueabihf.
 */
const lightningcssOtrosSo =
  process.platform === 'win32'
    ? [/node_modules[/\\]lightningcss-(?!win32)[^/\\]*[/\\]?.*/]
    : [];

/**
 * Restos de autolinking de Expo bajo node_modules (p. ej. tras quitar un paquete o con OneDrive).
 * Si Metro intenta vigilarlas y ya no existen → ENOENT en Windows.
 */
const expoAutolinkingStaleDirs = [
  /node_modules[/\\]\.expo-sharing-[^/\\]+[/\\]?.*/,
  /node_modules[/\\]\.expo-camera-[^/\\]+[/\\]?.*/,
];

config.resolver.blockList = [
  ...(config.resolver.blockList ?? []),
  ...poison,
  ...lightningcssOtrosSo,
  ...expoAutolinkingStaleDirs,
];

module.exports = config;
