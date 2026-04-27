import React, { useEffect, useMemo, useState, useCallback } from 'react';
import {
  View,
  Text,
  Modal,
  TouchableOpacity,
  ScrollView,
  StyleSheet,
  Pressable,
} from 'react-native';
import { Ionicons } from '@expo/vector-icons';
import { useSafeAreaInsets } from 'react-native-safe-area-context';
import { useApp } from '../context/AppContext';
import { reunirNotificacionesApp } from '../lib/notificacionesApp';
import {
  contarNoLeidas,
  firmaNotificacion,
  loadFirmasLectura,
  marcarAvisosActualesComoVistos,
  saveFirmasLectura,
} from '../lib/notificacionesLectura';
import { colors, spacing, radii, typography } from '../theme';

const SEV = {
  danger: { icon: 'alert-circle', color: colors.danger, bg: 'rgba(199, 123, 136, 0.15)' },
  warning: { icon: 'warning', color: colors.warning, bg: 'rgba(217, 180, 74, 0.12)' },
  info: { icon: 'information-circle', color: colors.chartBlue, bg: 'rgba(167, 216, 222, 0.1)' },
};

export function NotificacionBell() {
  const insets = useSafeAreaInsets();
  const { state } = useApp();
  const [open, setOpen] = useState(false);
  /** null = cargando; Record = firmas leídas por id */
  const [firmasLeidas, setFirmasLeidas] = useState(null);

  useEffect(() => {
    let active = true;
    loadFirmasLectura().then((m) => {
      if (active) setFirmasLeidas(m);
    });
    return () => {
      active = false;
    };
  }, []);

  const { items, total: totalAvisos } = useMemo(() => {
    if (!state) return { items: [], total: 0 };
    return reunirNotificacionesApp(state, new Date());
  }, [state]);

  const noLeidas = useMemo(() => {
    if (firmasLeidas === null) return 0;
    return contarNoLeidas(items, firmasLeidas);
  }, [items, firmasLeidas]);

  const abrirYMarcarComoVistas = useCallback(async () => {
    if (!state) {
      setOpen(true);
      return;
    }
    const cur = reunirNotificacionesApp(state, new Date()).items;
    setFirmasLeidas((prev) => {
      const base = prev === null ? {} : prev;
      const next = marcarAvisosActualesComoVistos(cur, base);
      saveFirmasLectura(next).catch(() => {});
      return next;
    });
    setOpen(true);
  }, [state]);

  return (
    <>
      <TouchableOpacity
        onPress={abrirYMarcarComoVistas}
        hitSlop={12}
        style={styles.bellWrap}
        accessibilityLabel={
          noLeidas > 0
            ? `Notificaciones, ${noLeidas} no leídas de ${totalAvisos}`
            : `Notificaciones, ${totalAvisos ? 'todo leído' : 'sin avisos'}`
        }
        accessibilityRole="button"
      >
        <Ionicons name="notifications-outline" size={26} color={colors.accentBright} />
        {noLeidas > 0 ? (
          <View style={styles.badge} accessibilityElementsHidden>
            <Text style={styles.badgeTxt}>{noLeidas > 9 ? '9+' : String(noLeidas)}</Text>
          </View>
        ) : null}
      </TouchableOpacity>

      <Modal
        visible={open}
        animationType="slide"
        transparent
        onRequestClose={() => setOpen(false)}
      >
        <View style={styles.modalRoot}>
          <Pressable style={styles.backdrop} onPress={() => setOpen(false)} />
          <View style={[styles.sheet, { paddingTop: insets.top > 0 ? 0 : spacing.md, paddingBottom: Math.max(insets.bottom, spacing.lg) }]}>
          <View style={styles.handle} />
          <View style={styles.sheetHead}>
            <Text style={typography.title}>Avisos</Text>
            <TouchableOpacity onPress={() => setOpen(false)} hitSlop={10}>
              <Ionicons name="close" size={28} color={colors.textMuted} />
            </TouchableOpacity>
          </View>
            <Text style={[typography.small, { color: colors.textFaint, marginBottom: spacing.md }]}>
            Al abrir, los avisos actuales se marcan como leídos. El contador solo sube con avisos nuevos o
            actualizados.
          </Text>

          {items.length === 0 ? (
            <View style={styles.empty}>
              <Ionicons name="checkmark-done-outline" size={48} color={colors.mint} />
              <Text style={[typography.body, { marginTop: spacing.md, textAlign: 'center' }]}>
                No hay avisos por ahora.
              </Text>
            </View>
          ) : (
            <ScrollView
              style={styles.list}
              contentContainerStyle={{ paddingBottom: spacing.lg }}
              showsVerticalScrollIndicator={false}
              keyboardShouldPersistTaps="handled"
            >
              {items.map((it) => {
                const s = SEV[it.severidad] || SEV.info;
                const visto = firmasLeidas != null && firmasLeidas[it.id] === firmaNotificacion(it);
                return (
                  <View
                    key={it.id}
                    style={[
                      styles.item,
                      { backgroundColor: s.bg, borderColor: colors.stroke, opacity: visto ? 0.72 : 1 },
                    ]}
                  >
                    <Ionicons name={s.icon} size={22} color={s.color} style={styles.itemIcon} />
                    <View style={{ flex: 1, minWidth: 0 }}>
                      <Text style={styles.itemTit}>{it.titulo}</Text>
                      <Text style={styles.itemSub}>{it.detalle}</Text>
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
  bellWrap: { position: 'relative', padding: 2, justifyContent: 'center' },
  badge: {
    position: 'absolute',
    right: -2,
    top: -4,
    minWidth: 18,
    height: 18,
    borderRadius: 9,
    backgroundColor: colors.danger,
    alignItems: 'center',
    justifyContent: 'center',
    paddingHorizontal: 4,
  },
  badgeTxt: { color: '#fff', fontSize: 10, fontWeight: '800' },
  modalRoot: { flex: 1, justifyContent: 'flex-end' },
  backdrop: { ...StyleSheet.absoluteFillObject, backgroundColor: 'rgba(0,0,0,0.5)' },
  sheet: {
    maxHeight: '88%',
    backgroundColor: colors.bgElevated,
    borderTopLeftRadius: radii.xl,
    borderTopRightRadius: radii.xl,
    borderWidth: 1,
    borderColor: colors.stroke,
    paddingHorizontal: spacing.lg,
    paddingTop: spacing.sm,
  },
  handle: {
    alignSelf: 'center',
    width: 40,
    height: 4,
    borderRadius: 2,
    backgroundColor: colors.strokeStrong,
    marginBottom: spacing.md,
  },
  sheetHead: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    marginBottom: spacing.xs,
  },
  list: { maxHeight: 480 },
  empty: { alignItems: 'center', paddingVertical: spacing.xl },
  item: {
    flexDirection: 'row',
    alignItems: 'flex-start',
    borderRadius: radii.md,
    borderWidth: 1,
    padding: spacing.md,
    marginBottom: spacing.sm,
  },
  itemIcon: { marginRight: spacing.md, marginTop: 2, flexShrink: 0 },
  itemTit: { color: colors.text, fontWeight: '700', fontSize: 15, lineHeight: 20 },
  itemSub: { color: colors.textSecondary, fontSize: 13, lineHeight: 19, marginTop: 4 },
});
