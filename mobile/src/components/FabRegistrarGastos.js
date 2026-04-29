import React, { useState } from 'react';
import { View, Pressable, StyleSheet, Platform, Modal, Text, TouchableOpacity } from 'react-native';
import { Ionicons } from '@expo/vector-icons';
import { rootNavigationRef } from '../navigation/rootNavigationRef';
import { colors, spacing, shadows, TAB_BAR_SCROLL_PADDING, radii, typography, iconSemantic } from '../theme';

/**
 * Botón flotante (+): abre un panel con tarjetas (registrar gasto, asistente de compras).
 * `visible` controlado desde el navigator (pestaña activa) — no usar hooks de navegación
 * en un componente colocado junto a Tab.Navigator.
 */
export default function FabRegistrarGastos({ visible = true }) {
  const [sheetOpen, setSheetOpen] = useState(false);

  if (!visible) return null;

  function cerrarYA(fn) {
    setSheetOpen(false);
    fn();
  }

  return (
    <View
      style={[styles.wrap, { bottom: TAB_BAR_SCROLL_PADDING + spacing.lg }]}
      pointerEvents="box-none"
      collapsable={false}
    >
      <Modal visible={sheetOpen} animationType="fade" transparent onRequestClose={() => setSheetOpen(false)}>
        <View style={styles.mOverlay}>
          <Pressable style={StyleSheet.absoluteFill} onPress={() => setSheetOpen(false)} />
          <View style={styles.mSheet}>
            <Text style={styles.sheetTitle}>Acciones rápidas</Text>
            <Text style={styles.sheetHint}>Elige cómo sigues desde el botón +</Text>

            <TouchableOpacity
              style={styles.row}
              onPress={() =>
                cerrarYA(() => {
                  if (rootNavigationRef.isReady()) rootNavigationRef.navigate('Gastos');
                })
              }
              activeOpacity={0.72}
              accessibilityRole="button"
              accessibilityLabel="Registrar gasto, ir a la pestaña Gastos"
            >
              <View
                style={[
                  styles.iconCircle,
                  { backgroundColor: iconSemantic.moreMenu.GastosFab.bg },
                ]}
              >
                <Ionicons
                  name="card-outline"
                  size={22}
                  color={iconSemantic.moreMenu.GastosFab.fg}
                />
              </View>
              <View style={styles.textBlock}>
                <Text style={styles.rowTitle}>Registrar gasto</Text>
                <Text style={styles.rowSub}>Nuevo gasto en la pestaña Gastos</Text>
              </View>
              <Ionicons name="chevron-forward" size={22} color={colors.textFaint} />
            </TouchableOpacity>

            <TouchableOpacity
              style={styles.row}
              onPress={() =>
                cerrarYA(() => {
                  if (!rootNavigationRef.isReady()) return;
                  rootNavigationRef.navigate('Mas', { screen: 'AsistenteCompras' });
                })
              }
              activeOpacity={0.72}
              accessibilityRole="button"
              accessibilityLabel="Abrir Asistente de compras"
            >
              <View
                style={[
                  styles.iconCircle,
                  { backgroundColor: iconSemantic.moreMenu.AsistenteCompras.bg },
                ]}
              >
                <Ionicons
                  name="basket-outline"
                  size={22}
                  color={iconSemantic.moreMenu.AsistenteCompras.fg}
                />
              </View>
              <View style={styles.textBlock}>
                <Text style={styles.rowTitle}>Asistente de compras</Text>
                <Text style={styles.rowSub}>Checklist y cupo mensual antes de comprar</Text>
              </View>
              <Ionicons name="chevron-forward" size={22} color={colors.textFaint} />
            </TouchableOpacity>

            <Pressable onPress={() => setSheetOpen(false)} style={styles.closeBtn}>
              <Text style={styles.closeBtnText}>Cerrar</Text>
            </Pressable>
          </View>
        </View>
      </Modal>

      <Pressable
        onPress={() => setSheetOpen(true)}
        style={({ pressed }) => [styles.fab, pressed && { opacity: 0.9 }]}
        accessibilityRole="button"
        accessibilityLabel="Abrir acciones rápidas: registrar gasto o asistente de compras"
      >
        <Ionicons name="add" size={32} color={colors.mint} />
      </Pressable>
    </View>
  );
}

const SIZE = 58;

const styles = StyleSheet.create({
  /**
   * Tamaño fijo: en Android un View absolute solo con right/bottom puede ocupar todo el ancho
   * y tapar listas/scroll (rectángulo oscuro + bloqueo de toques).
   */
  wrap: {
    position: 'absolute',
    right: spacing.lg,
    width: SIZE,
    height: SIZE,
    zIndex: 50,
    alignItems: 'center',
    justifyContent: 'center',
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
  mOverlay: {
    flex: 1,
    backgroundColor: 'rgba(0,0,0,0.52)',
    justifyContent: 'flex-end',
  },
  mSheet: {
    backgroundColor: colors.bgElevated,
    borderTopLeftRadius: radii.xl,
    borderTopRightRadius: radii.xl,
    padding: spacing.lg,
    paddingBottom: Platform.OS === 'ios' ? 34 : spacing.lg,
    borderWidth: 1,
    borderColor: colors.stroke,
    maxHeight: '88%',
  },
  sheetTitle: { ...typography.title, marginBottom: spacing.xs },
  sheetHint: { ...typography.small, color: colors.textMuted, marginBottom: spacing.md, lineHeight: 20 },
  row: {
    flexDirection: 'row',
    alignItems: 'center',
    backgroundColor: colors.surface,
    borderRadius: radii.lg,
    padding: spacing.md,
    marginBottom: spacing.sm,
    borderWidth: 1,
    borderColor: colors.stroke,
  },
  iconCircle: {
    width: 48,
    height: 48,
    borderRadius: radii.md,
    alignItems: 'center',
    justifyContent: 'center',
    marginRight: spacing.md,
    borderWidth: 1,
    borderColor: colors.stroke,
    flexShrink: 0,
  },
  textBlock: { flex: 1, minWidth: 0 },
  rowTitle: { color: colors.text, fontSize: 17, fontWeight: '700', letterSpacing: -0.2 },
  rowSub: { color: colors.textMuted, fontSize: 13, marginTop: 3, lineHeight: 18 },
  closeBtn: { alignSelf: 'center', marginTop: spacing.sm, paddingVertical: spacing.sm },
  closeBtnText: { ...typography.body, color: colors.accentBright, fontWeight: '600' },
});
