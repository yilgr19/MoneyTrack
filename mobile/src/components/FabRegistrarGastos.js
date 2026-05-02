import React, { useState, useMemo } from 'react';
import { View, Pressable, StyleSheet, Platform, Modal, Text, TouchableOpacity, useWindowDimensions } from 'react-native';
import { useSafeAreaInsets } from 'react-native-safe-area-context';
import { Ionicons } from '@expo/vector-icons';
import { rootNavigationRef } from '../navigation/rootNavigationRef';
import { useTheme } from '../context/ThemeContext';

/**
 * Botón flotante (+): abre un panel con tarjetas (registrar gasto, asistente de compras).
 * `visible` controlado desde el navigator (pestaña activa) — no usar hooks de navegación
 * en un componente colocado junto a Tab.Navigator.
 */
export default function FabRegistrarGastos({ visible = true }) {
  const [sheetOpen, setSheetOpen] = useState(false);
  const insets = useSafeAreaInsets();
  const { height: winH } = useWindowDimensions();
  const { colors, typography, shadows, iconSemantic, spacing, radii, TAB_BAR_SCROLL_PADDING, temaId } = useTheme();

  const styles = useMemo(
    () =>
      StyleSheet.create({
        layerFill: {
          ...StyleSheet.absoluteFillObject,
          zIndex: 2000,
          ...Platform.select({
            ios: {},
            android: { elevation: 24 },
            default: {},
          }),
        },
        wrap: {
          position: 'absolute',
          right: spacing.lg,
          width: SIZE,
          height: SIZE,
          zIndex: 1,
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
            android: { elevation: 26 },
          }),
        },
        mOverlay: {
          flex: 1,
          justifyContent: 'flex-end',
        },
        mBackdrop: {
          ...StyleSheet.absoluteFillObject,
          backgroundColor: 'rgba(0,0,0,0.52)',
          zIndex: 0,
        },
        mSheet: {
          zIndex: 1,
          width: '100%',
          alignSelf: 'stretch',
          backgroundColor: colors.bgElevated,
          borderTopLeftRadius: radii.xl,
          borderTopRightRadius: radii.xl,
          padding: spacing.lg,
          borderWidth: 1,
          borderColor: colors.stroke,
          ...Platform.select({
            ios: { ...shadows.card },
            android: { elevation: 16 },
            default: {},
          }),
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
      }),
    [temaId, colors, typography, shadows, spacing, radii]
  );

  if (!visible) return null;

  function cerrarYA(fn) {
    setSheetOpen(false);
    fn();
  }

  const bottomFab = TAB_BAR_SCROLL_PADDING + spacing.lg;
  const sheetPadBottom = Math.max(insets.bottom, spacing.lg) + (Platform.OS === 'ios' ? spacing.sm : 0);
  const sheetMaxH = Math.min(Math.round(winH * 0.88), winH - insets.top - spacing.md);

  return (
    <View style={styles.layerFill} pointerEvents="box-none" collapsable={false}>
      <Modal
        visible={sheetOpen}
        animationType="slide"
        transparent
        statusBarTranslucent={Platform.OS === 'android'}
        onRequestClose={() => setSheetOpen(false)}
      >
        <View style={styles.mOverlay}>
          <Pressable
            style={styles.mBackdrop}
            onPress={() => setSheetOpen(false)}
            accessibilityRole="button"
            accessibilityLabel="Cerrar acciones rápidas"
          />
          <View style={[styles.mSheet, { paddingBottom: sheetPadBottom, maxHeight: sheetMaxH }]}>
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

      <View
        style={[styles.wrap, { bottom: bottomFab }]}
        pointerEvents="box-none"
        collapsable={false}
      >
        <Pressable
          onPress={() => setSheetOpen(true)}
          style={({ pressed }) => [styles.fab, pressed && { opacity: 0.9 }]}
          accessibilityRole="button"
          accessibilityLabel="Abrir acciones rápidas: registrar gasto o asistente de compras"
          hitSlop={12}
          android_ripple={{ borderless: true, color: 'rgba(255,255,255,0.2)' }}
        >
          <Ionicons name="add" size={32} color={colors.mint} />
        </Pressable>
      </View>
    </View>
  );
}

const SIZE = 58;
