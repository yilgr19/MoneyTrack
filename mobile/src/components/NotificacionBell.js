import React, { useMemo, useState, useCallback } from 'react';
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
import { useNotificacionLectura } from '../context/NotificacionLecturaContext';
import { reunirNotificacionesApp } from '../lib/notificacionesApp';
import { contarNoLeidas, firmaNotificacion } from '../lib/notificacionesLectura';
import ExtractoBancarioModal from './ExtractoBancarioModal';
import { colors, spacing, radii, typography } from '../theme';

const SEV = {
  danger: { icon: 'heart-outline', color: colors.danger, bg: 'rgba(199, 123, 136, 0.18)' },
  warning: { icon: 'partly-sunny-outline', color: colors.warning, bg: 'rgba(217, 180, 74, 0.14)' },
  info: { icon: 'sparkles', color: colors.chartBlue, bg: 'rgba(167, 216, 222, 0.14)' },
};

const TIPO_ACENTO = {
  pago: { icon: 'calendar-outline', color: colors.mint },
  categoria: { icon: 'color-palette-outline', color: colors.accent },
  tc: { icon: 'card-outline', color: colors.mint },
  saldo: { icon: 'wallet-outline', color: colors.mint },
};

export function NotificacionBell() {
  const insets = useSafeAreaInsets();
  const { state } = useApp();
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

  const noLeidas = useMemo(() => {
    if (firmasLeidas === null) return 0;
    return contarNoLeidas(items, firmasLeidas);
  }, [items, firmasLeidas]);

  const abrirYMarcarComoVistas = useCallback(() => {
    if (!state) {
      setOpen(true);
      return;
    }
    marcarVistosAhora();
    setOpen(true);
  }, [state, marcarVistosAhora]);

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
            <View style={{ flex: 1, minWidth: 0, paddingRight: spacing.sm }}>
              <Text style={typography.label}>Avisos y consejos</Text>
              <Text style={typography.title}>Para que lleves el mes con calma</Text>
            </View>
            <TouchableOpacity onPress={() => setOpen(false)} hitSlop={10}>
              <Ionicons name="close" size={28} color={colors.textMuted} />
            </TouchableOpacity>
          </View>
            <Text style={[typography.small, { color: colors.textFaint, marginBottom: spacing.md, lineHeight: 20 }]}>
            Arriba lo más reciente/urgente. Al abrir, lo de ahora cuenta como leído; el badge vuelve si cambia
            algo.
          </Text>

          {items.length === 0 ? (
            <View style={styles.empty}>
              <Ionicons name="cafe-outline" size={50} color={colors.mint} />
              <Text style={[typography.body, { marginTop: spacing.md, textAlign: 'center', lineHeight: 24 }]}>
                Todo en orden por ahora. Cuando pase algo que te interese, lo verás aquí.
              </Text>
              <Text style={[typography.small, { marginTop: spacing.sm, textAlign: 'center', color: colors.textFaint }]}>
                Es como un recordatorio de un amigo: solo cuando toca.
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
                const ac = TIPO_ACENTO[it.tipo] || TIPO_ACENTO.pago;
                const visto = firmasLeidas != null && firmasLeidas[it.id] === firmaNotificacion(it);
                const body = (
                  <>
                    <View style={styles.dobleIcono}>
                      <Ionicons name={ac.icon} size={20} color={ac.color} />
                      <Ionicons name={s.icon} size={20} color={s.color} style={{ marginTop: 4 }} />
                    </View>
                    <View style={{ flex: 1, minWidth: 0 }}>
                      <Text style={styles.itemTit}>{it.titulo}</Text>
                      <Text style={styles.itemSub}>{it.detalle}</Text>
                    </View>
                    {it.tarjetaId ? (
                      <Ionicons name="chevron-forward" size={20} color={colors.textFaint} style={{ marginTop: 2 }} />
                    ) : null}
                  </>
                );
                const wrapStyle = [styles.item, { backgroundColor: s.bg, borderColor: colors.stroke, opacity: visto ? 0.7 : 1 }];
                if (it.tarjetaId) {
                  return (
                    <TouchableOpacity
                      key={it.id}
                      activeOpacity={0.8}
                      onPress={() => {
                        setOpen(false);
                        setExtractoTarjetaId(it.tarjetaId);
                      }}
                      style={wrapStyle}
                    >
                      {body}
                    </TouchableOpacity>
                  );
                }
                return (
                  <View key={it.id} style={wrapStyle}>
                    {body}
                  </View>
                );
              })}
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
  dobleIcono: { marginRight: spacing.md, marginTop: 2, flexShrink: 0, alignItems: 'center' },
  itemTit: { color: colors.text, fontWeight: '700', fontSize: 15, lineHeight: 20 },
  itemSub: { color: colors.textSecondary, fontSize: 13, lineHeight: 19, marginTop: 4 },
});
