import React, { useMemo, useState, useCallback, useEffect, useRef } from 'react';
import {
  View,
  Text,
  Modal,
  TouchableOpacity,
  Pressable,
  ScrollView,
  StyleSheet,
  TextInput,
  Alert,
  Platform,
  Animated,
  Easing,
} from 'react-native';
import { LinearGradient } from 'expo-linear-gradient';
import { Ionicons } from '@expo/vector-icons';
import { useSafeAreaInsets } from 'react-native-safe-area-context';
import { useApp } from '../context/AppContext';
import { formatearNumero, obtenerMesAño, montoGastoCuentaParaPresupuestoEnMes } from '../lib/finance';
import { PrimaryButton } from './Buttons';
import { colors, spacing, radii, typography, shadows } from '../theme';

function calcularResumenPresupuesto(state) {
  if (!state) {
    return {
      moneda: '',
      presupuestoMensual: 0,
      gastosMesActual: 0,
      ingresosMesActual: 0,
      flujoMes: 0,
      estadoMsg: '',
      estadoDetalle: '',
      estadoKind: 'info',
    };
  }
  const moneda = state.moneda || '';
  const presupuestoMensual = parseFloat(state.presupuestoMensual) || 0;
  const ahora = new Date();
  const mesActual = ahora.getMonth();
  const añoActual = ahora.getFullYear();
  const gastos = state.gastos || [];
  const ingresos = state.ingresos || [];

  const gastosMesActual = gastos.reduce(
    (s, g) => s + montoGastoCuentaParaPresupuestoEnMes(g, state, mesActual, añoActual),
    0
  );
  const ingresosMesActual = ingresos
    .filter((i) => {
      if (i.esRetiroBolsillo) return false;
      const { mes, año } = obtenerMesAño(i.fecha);
      return mes === mesActual && año === añoActual;
    })
    .reduce((s, i) => s + (parseFloat(i.cantidad) || 0), 0);
  const flujoMes = ingresosMesActual - gastosMesActual;

  let estadoMsg = '';
  let estadoDetalle = '';
  let estadoKind = 'info';
  if (presupuestoMensual <= 0) {
    estadoMsg = 'Sin tope mensual no hay semáforo (abajo o Saldo).';
    if (ingresosMesActual > 0 || gastosMesActual > 0) {
      estadoDetalle = `+${formatearNumero(ingresosMesActual)} / flujo ${formatearNumero(flujoMes)} ${moneda}`;
    }
  } else {
    const disponible = presupuestoMensual - gastosMesActual;
    const pctUsado = (gastosMesActual / presupuestoMensual) * 100;
    if (disponible > 0 && pctUsado < 80) {
      estadoMsg = '¡Dentro del tope!';
      estadoDetalle = `Quedan ${formatearNumero(disponible)} ${moneda} de tu límite de gasto.`;
      estadoKind = 'ok';
    } else if (disponible > 0 && pctUsado >= 80) {
      estadoMsg = 'Cerca del tope del mes';
      estadoDetalle = `Quedan ${formatearNumero(disponible)} ${moneda}.`;
      estadoKind = 'cuidado';
    } else if (disponible === 0) {
      estadoMsg = 'Límite de gasto alcanzado';
      estadoDetalle = `${formatearNumero(presupuestoMensual)} ${moneda} este mes.`;
      estadoKind = 'alerta';
    } else {
      estadoMsg = 'Sobre el tope fijado';
      estadoDetalle = `+${formatearNumero(Math.abs(disponible))} ${moneda} sobre límite.`;
      estadoKind = 'superado';
    }
  }

  return {
    moneda,
    presupuestoMensual,
    gastosMesActual,
    ingresosMesActual,
    flujoMes,
    estadoMsg,
    estadoDetalle,
    estadoKind,
  };
}

/**
 * Icono de medidor junto a deseos/campana: verde si vas bien, rojo si estás cerca o pasado del tope.
 * Abre el mismo resumen que la tarjeta de presupuesto en Inicio.
 */
export function PresupuestoMedidorBell() {
  const insets = useSafeAreaInsets();
  const { state, replaceState } = useApp();
  const [open, setOpen] = useState(false);
  const [draftTope, setDraftTope] = useState('');

  const r = useMemo(() => calcularResumenPresupuesto(state), [state]);

  const pctReal =
    r.presupuestoMensual > 0 ? (r.gastosMesActual / r.presupuestoMensual) * 100 : 0;
  const barColors =
    pctReal >= 100
      ? ['#fb7185', '#9f1239']
      : pctReal >= 80
        ? ['#fbbf24', '#c2410c']
        : ['#4ade80', '#14b8a6'];
  const tituloColor =
    r.estadoKind === 'ok'
      ? colors.mint
      : r.estadoKind === 'cuidado'
        ? colors.warning
        : r.estadoKind === 'alerta'
          ? colors.orange
          : r.estadoKind === 'superado'
            ? colors.danger
            : colors.textSecondary;

  const medidorOk = r.presupuestoMensual > 0 && r.estadoKind === 'ok';
  /** Límite alcanzado o superado: icono en rojo que debe parpadear */
  const presupuestoEnRojoCritico =
    r.presupuestoMensual > 0 && (r.estadoKind === 'alerta' || r.estadoKind === 'superado');

  const parpadeo = useRef(new Animated.Value(1)).current;
  useEffect(() => {
    if (!presupuestoEnRojoCritico) {
      parpadeo.setValue(1);
      return undefined;
    }
    const loop = Animated.loop(
      Animated.sequence([
        Animated.timing(parpadeo, {
          toValue: 0.38,
          duration: 520,
          easing: Easing.inOut(Easing.quad),
          useNativeDriver: false,
        }),
        Animated.timing(parpadeo, {
          toValue: 1,
          duration: 520,
          easing: Easing.inOut(Easing.quad),
          useNativeDriver: false,
        }),
      ])
    );
    loop.start();
    return () => loop.stop();
  }, [presupuestoEnRojoCritico, parpadeo]);

  const iconColor = r.presupuestoMensual <= 0 ? colors.textSecondary : medidorOk ? colors.mint : colors.danger;
  const bordeHalo =
    r.presupuestoMensual <= 0
      ? 'rgba(199, 195, 227, 0.12)'
      : medidorOk
        ? 'rgba(125, 193, 145, 0.45)'
        : 'rgba(248, 113, 131, 0.55)';

  const guardarTope = useCallback(() => {
    const n = parseFloat(String(draftTope).replace(',', '.')) || 0;
    replaceState((s) => ({ ...s, presupuestoMensual: n > 0 ? n : 0 }));
    setDraftTope('');
    if (n > 0) Alert.alert('Listo', 'Tope mensual guardado.');
  }, [draftTope, replaceState]);

  const disponibleGrid = r.presupuestoMensual - r.gastosMesActual;

  return (
    <>
      <TouchableOpacity
        onPress={() => setOpen(true)}
        hitSlop={10}
        style={styles.wrap}
        accessibilityRole="button"
        accessibilityLabel={
          r.presupuestoMensual <= 0
            ? 'Presupuesto del mes, sin tope configurado'
            : medidorOk
              ? 'Presupuesto del mes, dentro del tope'
              : 'Presupuesto del mes, cerca o sobre el tope'
        }
        activeOpacity={0.88}
      >
        <Animated.View style={[styles.halo, { borderColor: bordeHalo, opacity: parpadeo }]}>
          <Ionicons name="speedometer-outline" size={24} color={iconColor} />
        </Animated.View>
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
            <View style={styles.handle} />
            <View style={styles.sheetHead}>
              <Text style={styles.sheetTit} accessibilityRole="header">
                Presupuesto del mes
              </Text>
              <TouchableOpacity onPress={() => setOpen(false)} hitSlop={12} style={styles.cerrar}>
                <Ionicons name="close" size={22} color={colors.text} />
              </TouchableOpacity>
            </View>

            <ScrollView
              showsVerticalScrollIndicator={false}
              keyboardShouldPersistTaps="handled"
              contentContainerStyle={styles.scrollContent}
            >
              <LinearGradient
                colors={['rgba(91, 33, 182, 0.35)', 'rgba(12, 8, 18, 0.97)', 'rgba(8, 20, 28, 0.98)']}
                locations={[0, 0.45, 1]}
                start={{ x: 0, y: 0 }}
                end={{ x: 1, y: 1 }}
                style={styles.cardGrad}
              >
                <View style={styles.cardHeadCol}>
                  <View style={styles.cardHeadRow}>
                    <LinearGradient
                      colors={['rgba(167, 139, 250, 0.5)', 'rgba(45, 212, 191, 0.25)']}
                      style={styles.cardIcon}
                    >
                      <Ionicons name="speedometer-outline" size={26} color={colors.accentBright} />
                    </LinearGradient>
                    <View style={styles.cardHeadTxt}>
                      <Text style={styles.cardEyebrow}>Resumen del mes</Text>
                      {r.presupuestoMensual > 0 ? (
                        <>
                          <Text style={[styles.cardEstadoTit, { color: tituloColor }]}>{r.estadoMsg}</Text>
                          {r.estadoDetalle ? (
                            <Text style={styles.cardEstadoSub}>{r.estadoDetalle}</Text>
                          ) : null}
                        </>
                      ) : (
                        <Text style={styles.cardEstadoSub}>{r.estadoMsg}</Text>
                      )}
                    </View>
                  </View>
                  {r.presupuestoMensual > 0 ? (
                    <View style={styles.pctRow}>
                      <View style={styles.pctRing}>
                        <Text style={styles.pctBig}>{Math.min(999, Math.round(pctReal))}</Text>
                        <Text style={styles.pctSuf}>%</Text>
                      </View>
                    </View>
                  ) : null}
                </View>

                {r.presupuestoMensual > 0 ? (
                  <>
                    <View style={styles.barOuter}>
                      <LinearGradient
                        colors={barColors}
                        start={{ x: 0, y: 0.5 }}
                        end={{ x: 1, y: 0.5 }}
                        style={[
                          styles.barFill,
                          {
                            width: `${Math.min(100, pctReal)}%`,
                            minWidth: pctReal > 0.5 ? 4 : 0,
                          },
                        ]}
                      />
                    </View>
                    <View style={styles.grid}>
                      <View style={styles.cell}>
                        <Text style={styles.cellLab}>Tope</Text>
                        <Text style={styles.cellVal} numberOfLines={2} adjustsFontSizeToFit minimumFontScale={0.85}>
                          {formatearNumero(r.presupuestoMensual)} {r.moneda}
                        </Text>
                      </View>
                      <View style={[styles.cell, styles.cellSecond]}>
                        <Text style={styles.cellLab}>Queda</Text>
                        <Text
                          style={[
                            styles.cellVal,
                            {
                              color:
                                disponibleGrid > 0
                                  ? colors.mint
                                  : disponibleGrid < 0
                                    ? colors.danger
                                    : colors.textSecondary,
                            },
                          ]}
                          numberOfLines={2}
                          adjustsFontSizeToFit
                          minimumFontScale={0.85}
                        >
                          {formatearNumero(disponibleGrid)} {r.moneda}
                        </Text>
                      </View>
                    </View>
                    <Text style={styles.ingresoFlujo}>
                      Ingresos {formatearNumero(r.ingresosMesActual)} {r.moneda} · Flujo{' '}
                      <Text style={{ color: r.flujoMes >= 0 ? colors.mint : colors.danger }}>
                        {r.flujoMes >= 0 ? '+' : ''}
                        {formatearNumero(r.flujoMes)} {r.moneda}
                      </Text>
                    </Text>
                    <TouchableOpacity
                      onPress={() => {
                        Alert.alert('Presupuesto', '¿Eliminar presupuesto?', [
                          { text: 'Cancelar', style: 'cancel' },
                          {
                            text: 'Eliminar',
                            style: 'destructive',
                            onPress: () => replaceState((s) => ({ ...s, presupuestoMensual: 0 })),
                          },
                        ]);
                      }}
                      style={styles.eliminarTouch}
                    >
                      <Text style={styles.eliminarTxt}>Eliminar presupuesto</Text>
                    </TouchableOpacity>
                  </>
                ) : (
                  <View style={styles.sinTope}>
                    <Text style={styles.sinTopeLab}>Fijar tope rápido</Text>
                    <TextInput
                      style={styles.input}
                      value={draftTope}
                      onChangeText={setDraftTope}
                      placeholder="Ej. 2000"
                      placeholderTextColor={colors.textFaint}
                      keyboardType="decimal-pad"
                    />
                    <PrimaryButton title="Guardar tope" onPress={guardarTope} />
                  </View>
                )}
              </LinearGradient>
            </ScrollView>
          </View>
        </View>
      </Modal>
    </>
  );
}

const styles = StyleSheet.create({
  wrap: { marginRight: spacing.xs },
  halo: {
    width: 44,
    height: 44,
    borderRadius: 22,
    backgroundColor: 'rgba(0,0,0,0.2)',
    borderWidth: 1.5,
    alignItems: 'center',
    justifyContent: 'center',
  },
  modalRoot: { flex: 1, justifyContent: 'flex-end' },
  backdrop: { ...StyleSheet.absoluteFillObject, backgroundColor: 'rgba(0,0,0,0.55)' },
  sheet: {
    maxHeight: '88%',
    width: '100%',
    minWidth: 0,
    backgroundColor: colors.surfaceSolid,
    borderTopLeftRadius: radii.xl + 4,
    borderTopRightRadius: radii.xl + 4,
    borderWidth: 1,
    borderColor: 'rgba(199, 195, 227, 0.2)',
    paddingHorizontal: spacing.lg,
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
  sheetHead: {
    flexDirection: 'row',
    alignItems: 'flex-start',
    justifyContent: 'space-between',
    marginBottom: spacing.md,
    width: '100%',
    minWidth: 0,
  },
  sheetTit: {
    ...typography.title,
    fontSize: 18,
    flex: 1,
    minWidth: 0,
    marginRight: spacing.sm,
    lineHeight: 24,
    flexShrink: 1,
  },
  cerrar: {
    width: 40,
    height: 40,
    borderRadius: 20,
    backgroundColor: 'rgba(255,255,255,0.08)',
    alignItems: 'center',
    justifyContent: 'center',
    flexShrink: 0,
  },
  scrollContent: { paddingBottom: spacing.md, width: '100%' },
  cardGrad: {
    padding: spacing.lg,
    borderRadius: radii.xl,
    borderWidth: 1,
    borderColor: 'rgba(167, 139, 250, 0.35)',
    width: '100%',
    minWidth: 0,
  },
  cardHeadCol: {
    width: '100%',
    minWidth: 0,
    marginBottom: spacing.md,
  },
  cardHeadRow: {
    flexDirection: 'row',
    alignItems: 'flex-start',
    gap: spacing.md,
    width: '100%',
    minWidth: 0,
  },
  pctRow: {
    flexDirection: 'row',
    justifyContent: 'flex-end',
    width: '100%',
    marginTop: spacing.sm,
    minWidth: 0,
  },
  cardIcon: {
    width: 52,
    height: 52,
    borderRadius: 18,
    alignItems: 'center',
    justifyContent: 'center',
    borderWidth: 1,
    borderColor: 'rgba(255,255,255,0.14)',
    flexShrink: 0,
  },
  cardHeadTxt: { flex: 1, minWidth: 0, alignSelf: 'stretch' },
  cardEyebrow: {
    fontSize: 10,
    fontWeight: '800',
    letterSpacing: 1.2,
    textTransform: 'uppercase',
    color: colors.accent,
    marginBottom: 6,
  },
  cardEstadoTit: {
    fontSize: 19,
    fontWeight: '800',
    letterSpacing: -0.3,
    lineHeight: 24,
    flexShrink: 1,
    alignSelf: 'stretch',
  },
  cardEstadoSub: {
    fontSize: 13,
    color: colors.textSecondary,
    marginTop: 4,
    lineHeight: 20,
    fontWeight: '500',
    flexShrink: 1,
    alignSelf: 'stretch',
  },
  pctRing: {
    flexDirection: 'row',
    alignItems: 'baseline',
    flexShrink: 0,
    paddingHorizontal: 10,
    paddingVertical: 6,
    borderRadius: radii.lg,
    backgroundColor: 'rgba(0,0,0,0.35)',
    borderWidth: 1,
    borderColor: 'rgba(255,255,255,0.1)',
  },
  pctBig: { fontSize: 26, fontWeight: '900', color: colors.text, fontVariant: ['tabular-nums'] },
  pctSuf: { fontSize: 13, fontWeight: '800', color: colors.textMuted, marginLeft: 1 },
  barOuter: {
    height: 11,
    borderRadius: radii.pill,
    backgroundColor: 'rgba(0,0,0,0.35)',
    overflow: 'hidden',
    marginBottom: spacing.md,
    borderWidth: 1,
    borderColor: 'rgba(255,255,255,0.08)',
  },
  barFill: { height: '100%', borderRadius: radii.pill },
  grid: { flexDirection: 'row', marginBottom: spacing.sm, gap: spacing.md },
  cell: { flex: 1, minWidth: 0 },
  cellSecond: {
    borderLeftWidth: 1,
    borderColor: 'rgba(255,255,255,0.08)',
    paddingLeft: spacing.md,
  },
  cellLab: {
    fontSize: 10,
    fontWeight: '700',
    color: colors.textFaint,
    textTransform: 'uppercase',
    letterSpacing: 0.6,
    marginBottom: 4,
  },
  cellVal: { fontSize: 14, fontWeight: '700', color: colors.text, fontVariant: ['tabular-nums'] },
  ingresoFlujo: { fontSize: 12, color: colors.textFaint, textAlign: 'center', marginBottom: spacing.sm },
  eliminarTouch: { alignSelf: 'center', paddingVertical: spacing.xs },
  eliminarTxt: { color: colors.accentBright, fontWeight: '700', fontSize: 14 },
  sinTope: { marginTop: spacing.xs },
  sinTopeLab: { ...typography.small, color: colors.textMuted, marginBottom: spacing.sm },
  input: {
    borderWidth: 1,
    borderColor: colors.stroke,
    borderRadius: radii.md,
    padding: spacing.md,
    color: colors.text,
    fontSize: 16,
    marginBottom: spacing.sm,
    backgroundColor: 'rgba(0,0,0,0.2)',
  },
});
