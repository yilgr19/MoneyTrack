import React, { useMemo, useState, useCallback, useRef, useEffect } from 'react';
import {
  View,
  Text,
  Modal,
  TouchableOpacity,
  ScrollView,
  StyleSheet,
  Pressable,
  Animated,
  Easing,
  Platform,
} from 'react-native';
import { LinearGradient } from 'expo-linear-gradient';
import { Ionicons } from '@expo/vector-icons';
import { useSafeAreaInsets } from 'react-native-safe-area-context';
import { useApp } from '../context/AppContext';
import { formatearNumero } from '../lib/finance';
import { puedeRegistrarCompraPorRegla48h, formatCountdownMs } from '../lib/asistenteComprasLogic';
import { registrarGastoDesdeIntencionConUi, yaNoLoQuieroIntencionConUi } from '../lib/intencionesCompraAcciones';
import { rootNavigationRef } from '../navigation/rootNavigationRef';
import { colors, spacing, radii, typography, shadows } from '../theme';

export function DeseosCompraBell() {
  const insets = useSafeAreaInsets();
  const { state, replaceState } = useApp();
  const [open, setOpen] = useState(false);
  const [ahora, setAhora] = useState(Date.now());
  const moneda = state?.moneda || '';

  useEffect(() => {
    const t = setInterval(() => setAhora(Date.now()), 1000);
    return () => clearInterval(t);
  }, []);

  const pendientes = useMemo(() => {
    const list = state?.intencionesCompra || [];
    return list.filter((x) => x && x.estado === 'pendiente');
  }, [state?.intencionesCompra]);

  const n = pendientes.length;

  const abrirPanel = useCallback(() => setOpen(true), []);

  const badgePulse = useRef(new Animated.Value(1)).current;
  useEffect(() => {
    if (n <= 0) {
      badgePulse.setValue(1);
      return undefined;
    }
    const h = Animated.loop(
      Animated.sequence([
        Animated.timing(badgePulse, { toValue: 1.12, duration: 700, useNativeDriver: true }),
        Animated.timing(badgePulse, { toValue: 1, duration: 700, useNativeDriver: true }),
      ])
    );
    h.start();
    return () => h.stop();
  }, [n, badgePulse]);

  const iconRock = useRef(new Animated.Value(0)).current;
  useEffect(() => {
    if (n <= 0) {
      iconRock.setValue(0);
      return undefined;
    }
    const h = Animated.loop(
      Animated.sequence([
        Animated.timing(iconRock, {
          toValue: 1,
          duration: 280,
          easing: Easing.inOut(Easing.sin),
          useNativeDriver: true,
        }),
        Animated.timing(iconRock, {
          toValue: -1,
          duration: 280,
          easing: Easing.inOut(Easing.sin),
          useNativeDriver: true,
        }),
        Animated.timing(iconRock, {
          toValue: 0,
          duration: 280,
          easing: Easing.inOut(Easing.sin),
          useNativeDriver: true,
        }),
      ])
    );
    h.start();
    return () => h.stop();
  }, [n, iconRock]);

  const rot = iconRock.interpolate({
    inputRange: [-1, 0, 1],
    outputRange: ['-8deg', '0deg', '8deg'],
  });

  const irAsistenteDeseos = useCallback(() => {
    setOpen(false);
    rootNavigationRef.navigate('Mas', { screen: 'AsistenteCompras', params: { tab: 'deseos' } });
  }, []);

  return (
    <>
      <TouchableOpacity
        onPress={abrirPanel}
        hitSlop={12}
        style={[styles.bellWrap, n > 0 && styles.bellWrapActiva]}
        accessibilityLabel={
          n > 0 ? `Deseos de compra, ${n} pendientes` : 'Deseos de compra, sin pendientes'
        }
        accessibilityRole="button"
        activeOpacity={0.88}
      >
        <Animated.View style={[styles.bellHalo, { transform: [{ rotate: rot }] }]}>
          <Ionicons
            name={n > 0 ? 'heart' : 'heart-outline'}
            size={25}
            color={n > 0 ? colors.danger : colors.textSecondary}
          />
        </Animated.View>
        {n > 0 ? (
          <Animated.View style={[styles.badge, { transform: [{ scale: badgePulse }] }]} accessibilityElementsHidden>
            <LinearGradient
              colors={['#e879f9', '#a855f7']}
              start={{ x: 0, y: 0 }}
              end={{ x: 1, y: 1 }}
              style={StyleSheet.absoluteFill}
            />
            <Text style={styles.badgeTxt}>{n > 9 ? '9+' : String(n)}</Text>
          </Animated.View>
        ) : null}
      </TouchableOpacity>

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
            <LinearGradient
              colors={[colors.bgElevated, colors.bg]}
              start={{ x: 0, y: 0 }}
              end={{ x: 0.5, y: 1 }}
              style={StyleSheet.absoluteFill}
            />
            <View style={styles.handle} />
            <View style={styles.sheetHeadRow}>
              <View style={styles.titBloque}>
                <View style={styles.titFilaIcon}>
                  <LinearGradient
                    colors={['#5c2d58', '#4a2468']}
                    start={{ x: 0, y: 0 }}
                    end={{ x: 1, y: 1 }}
                    style={styles.titIconCirc}
                  >
                    <Ionicons name="heart" size={22} color="#e879f9" />
                  </LinearGradient>
                  <View style={styles.titTxtCol}>
                    <Text style={styles.titEtiq}>Deseos de compra</Text>
                    <Text style={styles.titGde}>Lista rápida de pendientes</Text>
                  </View>
                </View>
                {n > 0 ? (
                  <View style={styles.contadorFila}>
                    <View style={styles.contPill}>
                      <View style={styles.puntoOn} />
                      <Text style={styles.contPillTxt}>
                        {n} {n === 1 ? 'pendiente' : 'pendientes'}
                      </Text>
                    </View>
                  </View>
                ) : null}
              </View>
              <TouchableOpacity
                onPress={() => setOpen(false)}
                hitSlop={12}
                style={styles.cerrarCirc}
                accessibilityLabel="Cerrar deseos"
                activeOpacity={0.85}
              >
                <Ionicons name="close" size={22} color={colors.text} />
              </TouchableOpacity>
            </View>

            <TouchableOpacity onPress={irAsistenteDeseos} style={styles.linkAsistente} activeOpacity={0.85}>
              <Ionicons name="open-outline" size={18} color={colors.accent} style={styles.linkAsistenteIco} />
              <Text style={styles.linkAsistenteTxt}>Abrir asistente · Deseos</Text>
            </TouchableOpacity>

            {pendientes.length === 0 ? (
              <View style={styles.vacioBox}>
                <LinearGradient colors={['#2a1f32', '#1a2830']} style={styles.vacioHalo}>
                  <Ionicons name="heart-outline" size={56} color={colors.textSecondary} />
                </LinearGradient>
                <Text style={styles.vacioTit}>Sin deseos pendientes</Text>
                <Text style={styles.vacioSub}>Los deseos guardados aparecen aquí.</Text>
              </View>
            ) : (
              <ScrollView
                style={styles.list}
                contentContainerStyle={{ paddingBottom: spacing.xl }}
                showsVerticalScrollIndicator={false}
                keyboardShouldPersistTaps="handled"
              >
                {pendientes.map((intencion) => {
                  const puede = puedeRegistrarCompraPorRegla48h(intencion, ahora);
                  const hasta = intencion.cooldownHasta;
                  const esperaMs =
                    intencion.aplicabaCooldown && hasta != null && !puede ? Math.max(0, hasta - ahora) : 0;
                  const precio = intencion.precioEstimado;
                  const cat = String(intencion.nombreCategoria || '').trim() || '—';
                  const nom = String(intencion.nombre || '').trim() || 'Sin nombre';

                  return (
                    <View key={intencion.id} style={styles.filaCard}>
                      <LinearGradient
                        colors={['#2e2238', colors.surfaceSolid]}
                        start={{ x: 0, y: 0.5 }}
                        end={{ x: 1, y: 0.5 }}
                        style={styles.filaGrad}
                        pointerEvents="none"
                      />
                      <Text style={styles.filaTit}>{nom}</Text>
                      <Text style={styles.filaSub}>
                        {formatearNumero(precio)} {moneda} · {cat}
                      </Text>
                      {esperaMs > 0 ? (
                        <Text style={styles.cooldownTxt}>⏳ Podrás registrar en {formatCountdownMs(esperaMs)}</Text>
                      ) : null}
                      <View style={styles.accRow}>
                        <TouchableOpacity
                          style={[styles.btnGhost, !puede && styles.btnDisabled]}
                          disabled={!puede}
                          onPress={() =>
                            registrarGastoDesdeIntencionConUi({
                              state,
                              intencion,
                              origenValor: null,
                              replaceState,
                            })
                          }
                          activeOpacity={0.85}
                        >
                          <Ionicons name="bag-check-outline" size={18} color={puede ? colors.mint : colors.textMuted} />
                          <Text style={[styles.btnGhostTxt, !puede && styles.btnGhostTxtDis]}>Lo compré</Text>
                        </TouchableOpacity>
                        <TouchableOpacity
                          style={styles.btnDangerGhost}
                          onPress={() =>
                            yaNoLoQuieroIntencionConUi({
                              state,
                              intencion,
                              moneda,
                              replaceState,
                            })
                          }
                          activeOpacity={0.85}
                        >
                          <Ionicons name="heart-dislike-outline" size={18} color={colors.danger} />
                          <Text style={styles.btnDangerTxt}>Ya no lo quiero</Text>
                        </TouchableOpacity>
                      </View>
                    </View>
                  );
                })}
              </ScrollView>
            )}
          </View>
        </View>
      </Modal>
    </>
  );
}

const styles = StyleSheet.create({
  bellWrap: { position: 'relative', padding: 4, justifyContent: 'center', marginRight: spacing.xs },
  bellWrapActiva: {
    ...shadows.soft,
    borderRadius: radii.pill,
  },
  bellHalo: {
    width: 44,
    height: 44,
    borderRadius: 22,
    backgroundColor: 'rgba(0,0,0,0.2)',
    borderWidth: 1,
    borderColor: 'rgba(199, 195, 227, 0.12)',
    alignItems: 'center',
    justifyContent: 'center',
  },
  badge: {
    position: 'absolute',
    right: -1,
    top: 0,
    minWidth: 20,
    height: 20,
    borderRadius: 10,
    alignItems: 'center',
    justifyContent: 'center',
    paddingHorizontal: 5,
    overflow: 'hidden',
    borderWidth: 2,
    borderColor: colors.bg,
  },
  badgeTxt: { color: '#fff', fontSize: 10, fontWeight: '800', zIndex: 1 },
  modalRoot: { flex: 1, justifyContent: 'flex-end' },
  backdrop: { ...StyleSheet.absoluteFillObject, backgroundColor: 'rgba(0,0,0,0.6)' },
  sheet: {
    maxHeight: '90%',
    backgroundColor: colors.bg,
    borderTopLeftRadius: radii.xl + 4,
    borderTopRightRadius: radii.xl + 4,
    borderWidth: 1,
    borderColor: 'rgba(199, 195, 227, 0.18)',
    paddingHorizontal: spacing.lg,
    paddingTop: spacing.sm,
    overflow: 'hidden',
    ...Platform.select({ ios: shadows.card, android: { elevation: 12 } }),
  },
  handle: {
    alignSelf: 'center',
    width: 44,
    height: 5,
    borderRadius: 3,
    backgroundColor: 'rgba(199, 195, 227, 0.35)',
    marginBottom: spacing.lg,
  },
  sheetHeadRow: {
    flexDirection: 'row',
    alignItems: 'flex-start',
    justifyContent: 'space-between',
    marginBottom: spacing.md,
  },
  titBloque: { flex: 1, minWidth: 0, marginRight: spacing.sm },
  titFilaIcon: { flexDirection: 'row', alignItems: 'flex-start' },
  titTxtCol: { flex: 1, minWidth: 0, flexShrink: 1, paddingRight: 2 },
  titIconCirc: {
    width: 48,
    height: 48,
    borderRadius: 24,
    alignItems: 'center',
    justifyContent: 'center',
    marginRight: spacing.md,
    borderWidth: 1,
    borderColor: 'rgba(199, 195, 227, 0.15)',
    flexShrink: 0,
  },
  titEtiq: { fontSize: 10, fontWeight: '700', color: colors.accent, letterSpacing: 1.4, textTransform: 'uppercase' },
  titGde: {
    ...typography.title,
    fontSize: 19,
    marginTop: 4,
    lineHeight: 25,
    flexShrink: 1,
  },
  contadorFila: { marginTop: spacing.md },
  contPill: {
    flexDirection: 'row',
    alignItems: 'center',
    alignSelf: 'flex-start',
    backgroundColor: 'rgba(199, 195, 227, 0.1)',
    paddingHorizontal: 10,
    paddingVertical: 5,
    borderRadius: radii.pill,
    borderWidth: 1,
    borderColor: 'rgba(199, 195, 227, 0.2)',
  },
  puntoOn: { width: 6, height: 6, borderRadius: 3, backgroundColor: '#e879f9', marginRight: 8 },
  contPillTxt: { fontSize: 12, fontWeight: '700', color: colors.textSecondary },
  cerrarCirc: {
    width: 40,
    height: 40,
    borderRadius: 20,
    backgroundColor: 'rgba(255,255,255,0.08)',
    borderWidth: 1,
    borderColor: 'rgba(255,255,255,0.08)',
    alignItems: 'center',
    justifyContent: 'center',
    flexShrink: 0,
  },
  linkAsistente: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 8,
    marginBottom: spacing.lg,
    paddingVertical: spacing.xs,
    flexWrap: 'wrap',
    alignSelf: 'stretch',
  },
  linkAsistenteIco: { flexShrink: 0 },
  linkAsistenteTxt: {
    ...typography.body,
    color: colors.accent,
    fontWeight: '700',
    fontSize: 14,
    flexShrink: 1,
    flexGrow: 1,
    minWidth: 0,
  },
  list: { maxHeight: 500 },
  vacioBox: { alignItems: 'center', paddingVertical: spacing.xl + spacing.md, paddingHorizontal: spacing.md },
  vacioHalo: {
    width: 120,
    height: 120,
    borderRadius: 60,
    alignItems: 'center',
    justifyContent: 'center',
    borderWidth: 1,
    borderColor: 'rgba(232, 121, 249, 0.35)',
  },
  vacioTit: { fontSize: 20, fontWeight: '800', color: colors.text, marginTop: spacing.lg, textAlign: 'center' },
  vacioSub: { ...typography.body, textAlign: 'center', marginTop: spacing.sm, color: colors.textSecondary },
  filaCard: {
    borderRadius: radii.lg,
    borderWidth: 1,
    borderColor: 'rgba(199, 195, 227, 0.12)',
    overflow: 'hidden',
    padding: spacing.md,
    marginBottom: spacing.md,
    position: 'relative',
  },
  filaGrad: { ...StyleSheet.absoluteFillObject },
  filaTit: {
    color: colors.text,
    fontWeight: '800',
    fontSize: 16,
    lineHeight: 22,
    zIndex: 1,
    flexShrink: 1,
  },
  filaSub: {
    color: colors.textSecondary,
    fontSize: 14,
    lineHeight: 20,
    marginTop: 4,
    zIndex: 1,
    flexShrink: 1,
  },
  cooldownTxt: { ...typography.small, color: colors.warning, marginTop: spacing.sm, zIndex: 1 },
  accRow: { flexDirection: 'row', flexWrap: 'wrap', gap: spacing.sm, marginTop: spacing.md, zIndex: 1 },
  btnGhost: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 6,
    paddingVertical: 10,
    paddingHorizontal: 12,
    borderRadius: radii.pill,
    backgroundColor: 'rgba(125, 193, 145, 0.15)',
    borderWidth: 1,
    borderColor: 'rgba(125, 193, 145, 0.35)',
  },
  btnDisabled: { opacity: 0.55 },
  btnGhostTxt: { color: colors.mint, fontWeight: '800', fontSize: 13 },
  btnGhostTxtDis: { color: colors.textMuted },
  btnDangerGhost: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 6,
    paddingVertical: 10,
    paddingHorizontal: 12,
    borderRadius: radii.pill,
    backgroundColor: 'rgba(248, 113, 131, 0.12)',
    borderWidth: 1,
    borderColor: 'rgba(248, 113, 131, 0.35)',
  },
  btnDangerTxt: { color: colors.danger, fontWeight: '800', fontSize: 13 },
});
