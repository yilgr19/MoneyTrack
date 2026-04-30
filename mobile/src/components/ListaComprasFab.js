import React, { useMemo, useState, useCallback } from 'react';
import {
  View,
  Text,
  Modal,
  TouchableOpacity,
  Pressable,
  ScrollView,
  StyleSheet,
  Alert,
  Platform,
} from 'react-native';
import { LinearGradient } from 'expo-linear-gradient';
import { Ionicons } from '@expo/vector-icons';
import { useNavigation } from '@react-navigation/native';
import { useSafeAreaInsets } from 'react-native-safe-area-context';
import { useApp } from '../context/AppContext';
import { ordenarLineasListaSuper } from '../lib/asistenteComprasLogic';
import { colors, spacing, radii, typography, shadows, TAB_BAR_SCROLL_PADDING } from '../theme';

function etiquetaUrgencia(u) {
  if (u === 'urgente') return 'Urgente';
  if (u === 'puede_esperar') return 'Puede esperar';
  return 'Normal';
}

export default function ListaComprasFab() {
  const insets = useSafeAreaInsets();
  const navigation = useNavigation();
  const { state, replaceState } = useApp();
  const [open, setOpen] = useState(false);

  const pendientes = useMemo(
    () => ordenarLineasListaSuper(state?.listaSuperCompraItems || []),
    [state?.listaSuperCompraItems]
  );

  const urgentes = useMemo(() => pendientes.filter((l) => l.urgencia === 'urgente'), [pendientes]);
  const resto = useMemo(() => pendientes.filter((l) => l.urgencia !== 'urgente'), [pendientes]);

  const copyGancho = useMemo(() => {
    const salt = (pendientes.length || 0) % 3;
    const m = [
      'Revisa antes de comprar.',
      'Ítems que pediste recordar.',
      'Vistazo rápido al súper.',
    ];
    return m[salt];
  }, [pendientes.length]);

  const recordatorioGastosBtns = useMemo(
    () => [
      { text: 'OK', style: 'default' },
      { text: 'Ir a Gastos', onPress: () => navigation.navigate('Gastos') },
    ],
    [navigation]
  );

  const marcarComprado = useCallback(
    (ln) => {
      const nombre = String(ln?.nombre || '').trim() || 'Ítem';
      replaceState((s) => ({
        ...s,
        listaSuperCompraItems: (s.listaSuperCompraItems || []).filter((l) => l.id !== ln.id),
      }));
      Alert.alert(
        'Comprado',
        `«${nombre}» quitado de la lista. Registra el gasto en Gastos.`,
        recordatorioGastosBtns
      );
    },
    [replaceState, recordatorioGastosBtns]
  );

  const irListaCompleta = useCallback(() => {
    setOpen(false);
    navigation.navigate('Mas', { screen: 'AsistenteCompras', params: { tab: 'super' } });
  }, [navigation]);

  const marcarTodosComprados = useCallback(() => {
    const n = pendientes.length;
    if (n === 0) return;
    Alert.alert(
      'Todos comprados',
      `Se quitarán ${n} ítem${n !== 1 ? 's' : ''} de la lista. ¿Compraste todo en esta salida?`,
      [
        { text: 'Cancelar', style: 'cancel' },
        {
          text: 'Sí, vaciar lista',
          onPress: () => {
            replaceState((s) => ({ ...s, listaSuperCompraItems: [] }));
            setOpen(false);
            Alert.alert('Lista vaciada', 'Registra en Gastos lo que pagaste.', recordatorioGastosBtns);
          },
        },
      ]
    );
  }, [pendientes.length, replaceState, recordatorioGastosBtns]);

  if (pendientes.length === 0) {
    return null;
  }

  return (
    <>
      <View
        style={[styles.fabWrap, { bottom: TAB_BAR_SCROLL_PADDING + spacing.lg }]}
        pointerEvents="box-none"
        collapsable={false}
      >
        <TouchableOpacity
          onPress={() => setOpen(true)}
          activeOpacity={0.82}
          accessibilityRole="button"
          accessibilityLabel="Lista de compras"
          style={styles.fabTouch}
        >
          <View style={styles.fabInner}>
            <Ionicons name="basket" size={22} color={colors.mint} style={{ marginRight: 8 }} />
            <View style={styles.fabTxtCol}>
              <Text style={styles.fabTit}>Lista de</Text>
              <Text style={styles.fabTit}>compras</Text>
            </View>
          </View>
          <View style={styles.fabBadge}>
            <Text style={styles.fabBadgeTxt}>{pendientes.length > 99 ? '99+' : pendientes.length}</Text>
          </View>
        </TouchableOpacity>
      </View>

      <Modal visible={open} animationType="slide" transparent onRequestClose={() => setOpen(false)}>
        <View style={styles.modalRoot}>
          <Pressable style={styles.backdrop} onPress={() => setOpen(false)} android_ripple={null} />
          <View
            style={[
              styles.sheet,
              {
                paddingTop: insets.top > 0 ? spacing.xs : spacing.md,
                paddingBottom: Math.max(insets.bottom, spacing.lg),
              },
            ]}
          >
            <View style={styles.handle} />
            <View style={styles.headRow}>
              <Text style={styles.sheetTit}>Lista de compras</Text>
              <TouchableOpacity
                onPress={() => setOpen(false)}
                hitSlop={12}
                style={styles.cerrar}
                accessibilityLabel="Cerrar"
              >
                <Ionicons name="close" size={22} color={colors.text} />
              </TouchableOpacity>
            </View>

            <LinearGradient
              colors={['rgba(45, 212, 191, 0.22)', 'rgba(18, 14, 28, 0.98)', 'rgba(12, 8, 18, 1)']}
              start={{ x: 0, y: 0 }}
              end={{ x: 1, y: 1 }}
              style={styles.heroGrad}
            >
              <View style={styles.heroTop}>
                <View style={styles.heroIcon}>
                  <Ionicons name="basket" size={26} color="#2dd4bf" />
                </View>
                <View style={{ flex: 1, minWidth: 0 }}>
                  <Text style={styles.heroTitle}>Lista súper</Text>
                  <Text style={styles.heroSub}>{copyGancho}</Text>
                </View>
                <View style={styles.countPill}>
                  <Text style={styles.countNum}>{pendientes.length}</Text>
                  <Text style={styles.countLbl}>ítems</Text>
                </View>
              </View>
            </LinearGradient>

            <TouchableOpacity onPress={irListaCompleta} style={styles.linkFull} activeOpacity={0.85}>
              <Ionicons name="open-outline" size={18} color={colors.mint} />
              <Text style={styles.linkFullTxt}>Lista completa (Asistente)</Text>
            </TouchableOpacity>

            <TouchableOpacity
              style={styles.btnTodos}
              onPress={marcarTodosComprados}
              activeOpacity={0.88}
              accessibilityRole="button"
              accessibilityLabel="Todos comprados, vaciar lista"
            >
              <Ionicons name="checkmark-done-circle-outline" size={22} color={colors.accentBright} />
              <Text style={styles.btnTodosTxt}>Todos comprados</Text>
            </TouchableOpacity>
            <Text style={styles.btnTodosHint}>Si compraste todo, vacía aquí. O marca ítem por ítem.</Text>

            <ScrollView
              style={styles.scroll}
              contentContainerStyle={{ paddingBottom: spacing.xl }}
              showsVerticalScrollIndicator={false}
              keyboardShouldPersistTaps="handled"
            >
              {urgentes.length > 0 ? (
                <>
                  <Text style={styles.seccionTit}>Urgentes</Text>
                  {urgentes.map((ln) => (
                    <View key={ln.id} style={styles.row}>
                      <View style={styles.rowTxt}>
                        <Text style={styles.rowNombre}>{ln.nombre}</Text>
                        <Text style={[styles.rowUrg, { color: '#fb7185' }]}>{etiquetaUrgencia(ln.urgencia)}</Text>
                      </View>
                      <TouchableOpacity
                        style={styles.btnHecho}
                        onPress={() => marcarComprado(ln)}
                        activeOpacity={0.85}
                        accessibilityLabel={`Comprado: ${ln.nombre}`}
                      >
                        <Ionicons name="bag-check-outline" size={20} color={colors.mint} />
                        <Text style={styles.btnHechoTxt}>Comprado</Text>
                      </TouchableOpacity>
                    </View>
                  ))}
                </>
              ) : null}

              {resto.length > 0 ? (
                <>
                  <Text style={[styles.seccionTit, urgentes.length > 0 && { marginTop: spacing.md }]}>
                    {urgentes.length > 0 ? 'Resto del checklist' : 'Tu lista'}
                  </Text>
                  {resto.map((ln) => (
                    <View key={ln.id} style={styles.row}>
                      <View style={styles.rowTxt}>
                        <Text style={styles.rowNombre}>{ln.nombre}</Text>
                        <Text
                          style={[
                            styles.rowUrg,
                            ln.urgencia === 'puede_esperar' && { color: colors.textFaint },
                            ln.urgencia === 'normal' && { color: colors.chartBlue },
                          ]}
                        >
                          {etiquetaUrgencia(ln.urgencia)}
                        </Text>
                      </View>
                      <TouchableOpacity
                        style={styles.btnHecho}
                        onPress={() => marcarComprado(ln)}
                        activeOpacity={0.85}
                        accessibilityLabel={`Comprado: ${ln.nombre}`}
                      >
                        <Ionicons name="bag-check-outline" size={20} color={colors.mint} />
                        <Text style={styles.btnHechoTxt}>Comprado</Text>
                      </TouchableOpacity>
                    </View>
                  ))}
                </>
              ) : null}
            </ScrollView>
          </View>
        </View>
      </Modal>
    </>
  );
}

const styles = StyleSheet.create({
  fabWrap: {
    position: 'absolute',
    left: spacing.lg,
    zIndex: 49,
    maxWidth: '46%',
  },
  fabTouch: {
    position: 'relative',
    ...Platform.select({
      ios: {
        shadowColor: '#000',
        shadowOffset: { width: 0, height: 2 },
        shadowOpacity: 0.25,
        shadowRadius: 6,
      },
      android: { elevation: 4 },
      default: {},
    }),
  },
  fabInner: {
    flexDirection: 'row',
    alignItems: 'center',
    paddingVertical: 10,
    paddingHorizontal: 14,
    borderRadius: radii.pill,
    backgroundColor: 'rgba(12, 8, 18, 0.38)',
    borderWidth: 1,
    borderColor: 'rgba(125, 193, 145, 0.42)',
    minHeight: 52,
  },
  fabTxtCol: { flexShrink: 1 },
  fabTit: {
    fontSize: 12,
    fontWeight: '800',
    color: colors.text,
    letterSpacing: -0.2,
    lineHeight: 15,
    textShadowColor: 'rgba(0,0,0,0.45)',
    textShadowOffset: { width: 0, height: 1 },
    textShadowRadius: 2,
  },
  fabBadge: {
    position: 'absolute',
    top: -4,
    right: -4,
    minWidth: 20,
    height: 20,
    borderRadius: 10,
    backgroundColor: colors.danger,
    alignItems: 'center',
    justifyContent: 'center',
    paddingHorizontal: 5,
    borderWidth: 2,
    borderColor: colors.bg,
  },
  fabBadgeTxt: { color: '#fff', fontSize: 10, fontWeight: '800' },
  modalRoot: { flex: 1, justifyContent: 'flex-end' },
  backdrop: { ...StyleSheet.absoluteFillObject, backgroundColor: 'rgba(0,0,0,0.6)' },
  sheet: {
    maxHeight: '88%',
    backgroundColor: colors.surfaceSolid,
    borderTopLeftRadius: radii.xl + 4,
    borderTopRightRadius: radii.xl + 4,
    borderWidth: 1,
    borderColor: 'rgba(199, 195, 227, 0.22)',
    paddingHorizontal: spacing.lg,
    overflow: 'hidden',
    ...Platform.select({ ios: shadows.card, android: { elevation: 14 } }),
  },
  handle: {
    alignSelf: 'center',
    width: 44,
    height: 5,
    borderRadius: 3,
    backgroundColor: 'rgba(199, 195, 227, 0.35)',
    marginBottom: spacing.md,
  },
  headRow: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    marginBottom: spacing.md,
  },
  sheetTit: { ...typography.title, fontSize: 20, flex: 1, marginRight: spacing.sm },
  cerrar: {
    width: 40,
    height: 40,
    borderRadius: 20,
    backgroundColor: 'rgba(255,255,255,0.08)',
    alignItems: 'center',
    justifyContent: 'center',
    flexShrink: 0,
  },
  heroGrad: {
    borderRadius: radii.lg,
    padding: spacing.md,
    marginBottom: spacing.md,
    borderWidth: 1,
    borderColor: 'rgba(45, 212, 191, 0.35)',
  },
  heroTop: { flexDirection: 'row', alignItems: 'flex-start' },
  heroIcon: {
    width: 48,
    height: 48,
    borderRadius: 24,
    backgroundColor: 'rgba(45, 212, 191, 0.14)',
    borderWidth: 1,
    borderColor: 'rgba(45, 212, 191, 0.4)',
    alignItems: 'center',
    justifyContent: 'center',
    marginRight: spacing.md,
    flexShrink: 0,
  },
  heroTitle: {
    fontSize: 18,
    fontWeight: '800',
    color: colors.text,
    marginBottom: 4,
  },
  heroSub: { fontSize: 13, lineHeight: 19, color: colors.textSecondary, fontWeight: '500' },
  countPill: {
    alignItems: 'center',
    justifyContent: 'center',
    paddingHorizontal: spacing.sm,
    paddingVertical: 6,
    minWidth: 52,
    borderRadius: radii.md,
    backgroundColor: 'rgba(45, 212, 191, 0.2)',
    borderWidth: 1,
    borderColor: 'rgba(45, 212, 191, 0.45)',
    flexShrink: 0,
  },
  countNum: { fontSize: 20, fontWeight: '900', color: '#2dd4bf' },
  countLbl: { fontSize: 10, fontWeight: '700', color: colors.textMuted, textTransform: 'uppercase' },
  linkFull: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 8,
    marginBottom: spacing.md,
    alignSelf: 'flex-start',
  },
  linkFullTxt: { ...typography.body, color: colors.mint, fontWeight: '700', fontSize: 14, flexShrink: 1 },
  btnTodos: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'center',
    gap: 10,
    paddingVertical: 14,
    paddingHorizontal: spacing.md,
    borderRadius: radii.lg,
    backgroundColor: 'rgba(75, 36, 108, 0.45)',
    borderWidth: 1,
    borderColor: 'rgba(199, 195, 227, 0.35)',
    marginBottom: spacing.xs,
  },
  btnTodosTxt: { color: colors.accentBright, fontWeight: '800', fontSize: 16 },
  btnTodosHint: {
    ...typography.small,
    color: colors.textFaint,
    marginBottom: spacing.md,
    lineHeight: 18,
  },
  scroll: { maxHeight: 380 },
  seccionTit: {
    fontSize: 11,
    fontWeight: '800',
    letterSpacing: 1,
    color: colors.accent,
    textTransform: 'uppercase',
    marginBottom: spacing.sm,
  },
  row: {
    flexDirection: 'row',
    alignItems: 'center',
    paddingVertical: spacing.sm,
    paddingHorizontal: spacing.sm,
    marginBottom: spacing.xs,
    backgroundColor: colors.surface,
    borderRadius: radii.md,
    borderWidth: 1,
    borderColor: colors.stroke,
  },
  rowTxt: { flex: 1, minWidth: 0, marginRight: spacing.sm },
  rowNombre: { color: colors.text, fontWeight: '600', fontSize: 15 },
  rowUrg: { fontSize: 12, fontWeight: '600', marginTop: 2 },
  btnHecho: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 6,
    paddingVertical: 8,
    paddingHorizontal: 10,
    borderRadius: radii.pill,
    backgroundColor: 'rgba(125, 193, 145, 0.18)',
    borderWidth: 1,
    borderColor: 'rgba(125, 193, 145, 0.45)',
    flexShrink: 0,
  },
  btnHechoTxt: { color: colors.mint, fontWeight: '800', fontSize: 12 },
});
