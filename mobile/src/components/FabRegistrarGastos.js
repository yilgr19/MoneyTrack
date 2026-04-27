import React from 'react';
import { View, Pressable, StyleSheet, Platform } from 'react-native';
import { Ionicons } from '@expo/vector-icons';
import { rootNavigationRef } from '../navigation/rootNavigationRef';
import { colors, spacing, shadows, TAB_BAR_SCROLL_PADDING } from '../theme';

/**
 * Botón flotante (+) hacia la pestaña Gastos (registro de movimientos).
 * `visible` controlado desde el navigator (pestaña activa) — no usar hooks de navegación
 * en un componente colocado junto a Tab.Navigator.
 */
export default function FabRegistrarGastos({ visible = true }) {
  if (!visible) return null;

  const bottom = TAB_BAR_SCROLL_PADDING + spacing.lg;

  return (
    <View
      style={[styles.wrap, { bottom }]}
      pointerEvents="box-none"
    >
      <Pressable
        onPress={() => {
          if (rootNavigationRef.isReady()) rootNavigationRef.navigate('Gastos');
        }}
        style={({ pressed }) => [styles.fab, pressed && { opacity: 0.9 }]}
        accessibilityRole="button"
        accessibilityLabel="Registrar gasto, ir a la pestaña Gastos"
      >
        <Ionicons name="add" size={32} color={colors.mint} />
      </Pressable>
    </View>
  );
}

const SIZE = 58;

const styles = StyleSheet.create({
  wrap: {
    position: 'absolute',
    right: spacing.lg,
    zIndex: 50,
  },
  fab: {
    width: SIZE,
    height: SIZE,
    borderRadius: SIZE / 2,
    backgroundColor: colors.accentDeep,
    alignItems: 'center',
    justifyContent: 'center',
    borderWidth: 2,
    borderColor: colors.accent,
    ...Platform.select({
      ios: { ...shadows.card },
      android: { elevation: 8 },
    }),
  },
});
