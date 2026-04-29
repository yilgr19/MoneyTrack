// https://docs.expo.dev/guides/customizing-metro
const { getDefaultConfig } = require('expo/metro-config');

/** @type {import('expo/metro-config').MetroConfig} */
const config = getDefaultConfig(__dirname);

/**
 * En Windows + OneDrive, Gradle puede dejar bajo node_modules rutas raras en
 * expo-constants/android/bin; el FallbackWatcher de Metro hace lstat y revienta.
 */
const poison = [/node_modules[/\\]expo-constants[/\\]android[/\\]bin[/\\].*/];
config.resolver.blockList = [...(config.resolver.blockList ?? []), ...poison];

module.exports = config;
