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
import { useNotificacionLectura } from '../context/NotificacionLecturaContext';
import { reunirNotificacionesApp, stateSinAvisosGastoMovimientoEnLista } from '../lib/notificacionesApp';
import { contarNoLeidas, firmaNotificacion } from '../lib/notificacionesLectura';
import ExtractoBancarioModal from './ExtractoBancarioModal';
import { colors, spacing, radii, typography, shadows } from '../theme';

const SEV = {
  danger: {
    icon: 'flame-outline',
    color: colors.danger,
    line: 'rgba(248, 113, 131, 0.95)',
    grad: ['#2d2228', colors.surfaceSolid],
  },
  warning: {
    icon: 'partly-sunny-outline',
    color: colors.warning,
    line: 'rgba(251, 191, 36, 0.9)',
    grad: ['#2c2618', colors.surfaceSolid],
  },
  info: {
    icon: 'sparkles',
    color: colors.chartBlue,
    line: 'rgba(125, 211, 192, 0.55)',
    grad: ['#1a2528', colors.surfaceSolid],
  },
};

const TIPO_ACENTO = {
  pago: { icon: 'calendar-outline', color: colors.accentGold, label: 'Pago' },
  categoria: { icon: 'color-palette-outline', color: colors.accent, label: 'Categoría' },
  presupuesto: { icon: 'speedometer-outline', color: colors.accentGold, label: 'Presupuesto' },
  tc: { icon: 'card-outline', color: colors.chartBlue, label: 'Tarjeta' },
  saldo: { icon: 'wallet-outline', color: colors.mint, label: 'Saldo' },
  listaSuper: { icon: 'basket-outline', color: colors.mint, label: 'Lista súper' },
  gasto_movimiento: { icon: 'receipt-outline', color: colors.accent, label: 'Gasto' },
  meta: { icon: 'flag-outline', color: colors.success, label: 'Meta' },
};

function NotificacionFila({ it, index, open, onTarjeta, esTocable }) {
  const s = SEV[it.severidad] || SEV.info;
  const ac = TIPO_ACENTO[it.tipo] || TIPO_ACENTO.pago;
  const op = useRef(new Animated.Value(0)).current;
  const y = useRef(new Animated.Value(22)).current;

  useEffect(() => {
    if (!open) {
      op.setValue(0);
      y.setValue(22);
      return;
    }
    const delay = 45 + index * 62;
    op.setValue(0);
    y.setValue(16);
    /*
     * No mezclar useNativeDriver true/false en el mismo nodo: en Android la opacidad puede quedar en 0
     * y la lista se ve “vacía” aunque haya ítems. Opacity en vista hija (JS), translate en padre (nativo).
     */
    Animated.parallel([
      Animated.timing(op, {
        toValue: 1,
        duration: 420,
        delay,
        easing: Easing.out(Easing.cubic),
        useNativeDriver: false,
      }),
      Animated.spring(y, {
        toValue: 0,
        delay,
        friction: 8,
        tension: 70,
        useNativeDriver: true,
      }),
    ]).start();
  }, [open, it.id, index, op, y]);

  const cuerpo = (
    <Animated.View style={[styles.filaCard, { borderLeftColor: s.line, transform: [{ translateY: y }] }]}>
      <Animated.View style={[styles.filaCardInner, { opacity: op }]}>
        <LinearGradient
          colors={s.grad}
          start={{ x: 0, y: 0.5 }}
          end={{ x: 1, y: 0.5 }}
          style={styles.filaGrad}
          pointerEvents="none"
        />
        <View style={styles.filaTop}>
          <LinearGradient
            colors={[ac.color + '45', ac.color + '14']}
            start={{ x: 0, y: 0 }}
            end={{ x: 1, y: 1 }}
            style={styles.iconHalo}
          >
            <Ionicons name={ac.icon} size={22} color={ac.color} />
          </LinearGradient>
          <View style={styles.filaTxt}>
            <View style={styles.filaChips}>
              <View style={[styles.tipoChip, { borderColor: ac.color + '55' }]}>
                <Text style={[styles.tipoChipTxt, { color: ac.color }]}>{ac.label}</Text>
              </View>
              <Ionicons name={s.icon} size={15} color={s.color} style={{ marginLeft: 4 }} />
            </View>
            <Text style={styles.filaTit} selectable>
              {it.titulo}
            </Text>
            <Text style={styles.filaSub} selectable>
              {it.detalle}
            </Text>
          </View>
          {esTocable ? (
            <View style={styles.flechaPill}>
              <Ionicons name="chevron-forward" size={18} color={colors.accent} />
            </View>
          ) : (
            <View style={styles.flechaSpacer} />
          )}
        </View>
      </Animated.View>
    </Animated.View>
  );

  if (esTocable) {
    return (
      <Pressable
        onPress={onTarjeta}
        style={({ pressed }) => [styles.filaPress, pressed && styles.filaPressOn]}
        android_ripple={{ color: 'rgba(199, 195, 227, 0.2)', borderless: false }}
      >
        {cuerpo}
      </Pressable>
    );
  }
  return <View style={styles.filaPress}>{cuerpo}</View>;
}

export function NotificacionBell() {
  const insets = useSafeAreaInsets();
  const { state, replaceState } = useApp();
  const { firmasLeidas, marcarVistosAhora } = useNotificacionLectura();
  const [open, setOpen] = useState(false);
  const [extractoTarjetaId, setExtractoTarjetaId] = useState(null);
  const moneda = state?.moneda || '';
  const tarjetaExtracto = useMemo(
    () => (state?.tarjetasCredito || []).find((x) => x && x.id === extractoTarjetaId) || null,
    [state?.tarjetasCredito, extractoTarjetaId]
  );

  const { items, total: totalAvisos } = useMemo(() => {
    if (!state) return { items: [], total: 0 };
    return reunirNotificacionesApp(state, new Date());
  }, [state]);

  const noLeidas = useMemo(
    () => contarNoLeidas(items, firmasLeidas ?? {}),
    [items, firmasLeidas]
  );

  const itemsVisibles = useMemo(() => {
    if (firmasLeidas == null) return items;
    return items.filter((it) => firmasLeidas[it.id] !== firmaNotificacion(it));
  }, [items, firmasLeidas]);

  const openRef = useRef(false);
  useEffect(() => {
    if (openRef.current && !open) {
      if (state) {
        const cur = reunirNotificacionesApp(state, new Date()).items;
        marcarVistosAhora();
        replaceState((s) => stateSinAvisosGastoMovimientoEnLista(s, cur));
      }
    }
    openRef.current = open;
  }, [open, state, marcarVistosAhora, replaceState]);

  const abrirPanel = useCallback(() => {
    setOpen(true);
  }, []);

  const badgePulse = useRef(new Animated.Value(1)).current;
  useEffect(() => {
    if (noLeidas <= 0) {
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
  }, [noLeidas, badgePulse]);

  return (
    <>
      <TouchableOpacity
        onPress={abrirPanel}
        hitSlop={12}
        style={[styles.bellWrap, noLeidas > 0 && styles.bellWrapActiva]}
        accessibilityLabel={
          noLeidas > 0
            ? `Notificaciones, ${noLeidas} no leídas de ${totalAvisos}`
            : `Notificaciones, ${totalAvisos ? 'todo leído' : 'sin avisos'}`
        }
        accessibilityRole="button"
        activeOpacity={0.88}
      >
        <View style={styles.bellHalo}>
          <Ionicons
            name={noLeidas > 0 ? 'notifications' : 'notifications-outline'}
            size={25}
            color={noLeidas > 0 ? colors.accentGold : colors.textSecondary}
          />
        </View>
        {noLeidas > 0 ? (
          <Animated.View style={[styles.badge, { transform: [{ scale: badgePulse }] }]} accessibilityElementsHidden>
            <LinearGradient
              colors={['#f87171', '#dc2626']}
              start={{ x: 0, y: 0 }}
              end={{ x: 1, y: 1 }}
              style={StyleSheet.absoluteFill}
            />
            <Text style={styles.badgeTxt}>{noLeidas > 9 ? '9+' : String(noLeidas)}</Text>
          </Animated.View>
        ) : null}
      </TouchableOpacity>

      <Modal visible={open} animationType="slide" transparent onRequestClose={() => setOpen(false)}>
        <View style={styles.modalRoot}>
          <Pressable style={styles.backdrop} onPress={() => setOpen(false)} android_ripple={null} />
          <View
            style={[
              styles.sheet,
              { paddingTop: insets.top > 0 ? spacing.xs : spacing.md, paddingBottom: Math.max(insets.bottom, spacing.lg) },
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
                    colors={['#4a3d2a', '#3d2858']}
                    start={{ x: 0, y: 0 }}
                    end={{ x: 1, y: 1 }}
                    style={styles.titIconCirc}
                  >
                    <Ionicons name="notifications" size={22} color={colors.accentGold} />
                  </LinearGradient>
                  <View style={styles.titTxtCol}>
                    <Text style={styles.titEtiq}>Centro de avisos</Text>
                    <Text style={styles.titGde}>Avisos recientes</Text>
                  </View>
                </View>
                {itemsVisibles.length > 0 ? (
                  <View style={styles.contadorFila}>
                    <View style={styles.contPill}>
                      <View style={styles.puntoOn} />
                      <Text style={styles.contPillTxt}>
                        {itemsVisibles.length} {itemsVisibles.length === 1 ? 'nuevo' : 'nuevos'}
                      </Text>
                    </View>
                  </View>
                ) : null}
              </View>
              <TouchableOpacity
                onPress={() => setOpen(false)}
                hitSlop={12}
                style={styles.cerrarCirc}
                accessibilityLabel="Cerrar notificaciones"
                activeOpacity={0.85}
              >
                <Ionicons name="close" size={22} color={colors.text} />
              </TouchableOpacity>
            </View>
            <Text style={styles.ayudTxt}>Arriba = más reciente. Al cerrar, se ocultan hasta que haya algo nuevo.</Text>

            {itemsVisibles.length === 0 ? (
              <View style={styles.vacioBox}>
                <LinearGradient colors={['#1e2e24', '#1a2830']} style={styles.vacioHalo}>
                  <Ionicons name="checkmark-circle" size={56} color={colors.mint} />
                </LinearGradient>
                <Text style={styles.vacioTit}>¡Bien! Sin pendientes</Text>
                <Text style={styles.vacioSub}>Aquí verás pagos y avisos cuando existan.</Text>
              </View>
            ) : (
              <ScrollView
                style={styles.list}
                contentContainerStyle={styles.listContent}
                showsVerticalScrollIndicator={false}
                keyboardShouldPersistTaps="handled"
              >
                {itemsVisibles.map((it, index) => (
                  <NotificacionFila
                    key={it.id}
                    it={it}
                    index={index}
                    open={open}
                    esTocable={!!it.tarjetaId}
                    onTarjeta={() => {
                      setOpen(false);
                      setExtractoTarjetaId(it.tarjetaId);
                    }}
                  />
                ))}
              </ScrollView>
            )}
          </View>
        </View>
      </Modal>
      <ExtractoBancarioModal
        visible={!!tarjetaExtracto}
        onClose={() => setExtractoTarjetaId(null)}
        state={state}
        tarjeta={tarjetaExtracto}
        moneda={moneda}
      />
    </>
  );
}

const styles = StyleSheet.create({
  bellWrap: { position: 'relative', padding: 4, justifyContent: 'center' },
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
  titFilaIcon: {
    flexDirection: 'row',
    alignItems: 'flex-start',
    width: '100%',
    minWidth: 0,
    alignSelf: 'stretch',
  },
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
    width: '100%',
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
  puntoOn: { width: 6, height: 6, borderRadius: 3, backgroundColor: colors.mint, marginRight: 8 },
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
  ayudTxt: {
    ...typography.small,
    color: colors.textFaint,
    marginBottom: spacing.lg,
    lineHeight: 19,
    flexShrink: 1,
    width: '100%',
  },
  list: { maxHeight: 500, width: '100%', minHeight: 120 },
  listContent: { paddingBottom: spacing.xl, width: '100%' },
  vacioBox: { alignItems: 'center', paddingVertical: spacing.xl + spacing.md, paddingHorizontal: spacing.md },
  vacioHalo: {
    width: 120,
    height: 120,
    borderRadius: 60,
    alignItems: 'center',
    justifyContent: 'center',
    borderWidth: 1,
    borderColor: 'rgba(125, 193, 145, 0.35)',
  },
  vacioTit: { fontSize: 20, fontWeight: '800', color: colors.text, marginTop: spacing.lg, textAlign: 'center' },
  vacioSub: { ...typography.body, textAlign: 'center', marginTop: spacing.sm, color: colors.textSecondary },
  filaPress: {
    marginBottom: spacing.md,
    borderRadius: radii.lg,
    overflow: 'visible',
    width: '100%',
    alignSelf: 'stretch',
  },
  filaPressOn: { transform: [{ scale: 0.99 }] },
  filaCard: {
    borderRadius: radii.lg,
    borderWidth: 1,
    borderColor: 'rgba(199, 195, 227, 0.12)',
    borderLeftWidth: 4,
    overflow: 'hidden',
    position: 'relative',
    width: '100%',
    alignSelf: 'stretch',
  },
  filaCardInner: {
    padding: spacing.md,
    paddingBottom: spacing.md + 4,
    position: 'relative',
    width: '100%',
  },
  filaGrad: { ...StyleSheet.absoluteFillObject },
  filaTop: {
    flexDirection: 'row',
    alignItems: 'flex-start',
    width: '100%',
    minWidth: 0,
    alignSelf: 'stretch',
  },
  iconHalo: {
    width: 48,
    height: 48,
    borderRadius: 16,
    alignItems: 'center',
    justifyContent: 'center',
    marginRight: spacing.md,
    borderWidth: 1,
    borderColor: 'rgba(255,255,255,0.1)',
    flexShrink: 0,
  },
  filaTxt: {
    flex: 1,
    minWidth: 0,
    flexShrink: 1,
    zIndex: 1,
    alignSelf: 'stretch',
  },
  filaChips: { flexDirection: 'row', alignItems: 'center', marginBottom: 6, flexWrap: 'wrap' },
  tipoChip: {
    borderWidth: 1,
    borderRadius: radii.pill,
    paddingHorizontal: 8,
    paddingVertical: 2,
    backgroundColor: 'rgba(0,0,0,0.15)',
  },
  tipoChipTxt: { fontSize: 10, fontWeight: '800', textTransform: 'uppercase', letterSpacing: 0.6 },
  filaTit: {
    color: colors.text,
    fontWeight: '800',
    fontSize: 16,
    lineHeight: 22,
    flexShrink: 1,
    alignSelf: 'stretch',
  },
  filaSub: {
    color: colors.textSecondary,
    fontSize: 14,
    lineHeight: 21,
    marginTop: 6,
    flexShrink: 1,
    alignSelf: 'stretch',
  },
  flechaPill: {
    width: 32,
    height: 32,
    borderRadius: 16,
    backgroundColor: 'rgba(199, 195, 227, 0.12)',
    alignItems: 'center',
    justifyContent: 'center',
    marginLeft: 4,
    flexShrink: 0,
  },
  flechaSpacer: { width: 8 },
});
