import { Platform } from 'react-native';
import Constants, { ExecutionEnvironment } from 'expo-constants';

/**
 * `expo-text-extractor` hace `requireNativeModule('ExpoTextExtractor')` al cargarse → en Expo Go revienta.
 * En Expo Go no hay módulos nativos propios; hace falta build con `prebuild`/`run:android` (o EAS).
 *
 * No usar `NativeModules.ExpoTextExtractor`: con TurboModules / nueva arquitectura suele ser null
 * hasta cargar el módulo, y el OCR nativo nunca se intentaba aunque el APK sí lo incluye.
 */
export function getOptionalExpoTextExtractor() {
  if (Platform.OS === 'web') return null;

  const enExpoGo =
    Constants.executionEnvironment === ExecutionEnvironment.StoreClient ||
    Constants.appOwnership === 'expo';
  if (enExpoGo) return null;

  try {
    // eslint-disable-next-line global-require
    return require('expo-text-extractor');
  } catch {
    return null;
  }
}
