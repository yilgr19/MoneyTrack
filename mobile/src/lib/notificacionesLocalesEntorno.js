import { isRunningInExpoGo } from 'expo';
import { Platform } from 'react-native';

/**
 * `expo-notifications` en Expo Go (SDK 53+) dispara avisos/errores: no se debe importar ese módulo.
 * Sí carga con un development build o build de producción. En web no hay notificaciones nativas.
 */
export function notificacionesSistemaDisponibles() {
  if (Platform.OS === 'web') return false;
  return !isRunningInExpoGo();
}
