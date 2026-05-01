import { isRunningInExpoGo } from 'expo';
import { Platform } from 'react-native';

/** True si la app corre en Expo Go (no hay notificaciones programadas en la barra del sistema). */
export function entornoEsExpoGo() {
  if (Platform.OS === 'web') return false;
  return isRunningInExpoGo();
}

/**
 * `expo-notifications` en Expo Go (SDK 53+) dispara avisos/errores: no se debe importar ese módulo.
 * Sí carga con un development build o build de producción. En web no hay notificaciones nativas.
 */
export function notificacionesSistemaDisponibles() {
  if (Platform.OS === 'web') return false;
  return !isRunningInExpoGo();
}
