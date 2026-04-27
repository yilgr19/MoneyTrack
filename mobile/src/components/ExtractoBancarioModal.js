import React, { useMemo } from 'react';
import {
  View,
  Text,
  Modal,
  TouchableOpacity,
  ScrollView,
  Pressable,
  StyleSheet,
} from 'react-native';
import { Ionicons } from '@expo/vector-icons';
import { useSafeAreaInsets } from 'react-native-safe-area-context';
import { formatearNumero, construirExtractoBancarioTarjeta } from '../lib/finance';
import { colors, spacing, radii, typography } from '../theme';

/**
 * Extracto tipo banco: fechas, cupo, movimientos del corte, totales y proyección 3/6 cuotas.
 */
export default function ExtractoBancarioModal({ visible, onClose, state, tarjeta, moneda }) {
  const insets = useSafeAreaInsets();

  const ext = useMemo(() => {
    if (!visible || !tarjeta || !state) return null;
    return construirExtractoBancarioTarjeta(tarjeta, state, new Date());
  }, [visible, tarjeta, state]);

  if (!ext) return null;

  return (
    <Modal visible={visible} animationType="slide" transparent onRequestClose={onClose}>
      <View style={styles.root}>
        <Pressable style={styles.backdrop} onPress={onClose} />
        <View
          style={[
            styles.sheet,
            {
              paddingTop: spacing.md,
              paddingBottom: Math.max(insets.bottom, spacing.lg),
            },
          ]}
        >
          <View style={styles.headRow}>
            <View style={{ flex: 1, minWidth: 0 }}>
              <Text style={typography.label}>Extracto de corte</Text>
              <Text style={[typography.title, { marginTop: 4 }]} numberOfLines={2}>
                {ext.nombre}
              </Text>
            </View>
            <TouchableOpacity onPress={onClose} hitSlop={12} accessibilityLabel="Cerrar">
              <Ionicons name="close" size={28} color={colors.textMuted} />
            </TouchableOpacity>
          </View>

          <ScrollView
            style={styles.scroll}
            contentContainerStyle={{ paddingBottom: spacing.xl }}
            showsVerticalScrollIndicator={false}
            keyboardShouldPersistTaps="handled"
          >
            <Text style={styles.sectionTitle}>1. Fechas (control de tiempos)</Text>
            <View style={styles.block}>
              <FilaL label="Fecha de corte" value={ext.etiquetaCorte} />
              <FilaL label="Fecha límite de pago" value={ext.etiquetaProxPago} />
            </View>

            <Text style={styles.sectionTitle}>2. Resumen de disponibilidad (cupo)</Text>
            <View style={styles.block}>
              <FilaM label="Cupo total aprobado" monto={ext.cupoTotal} moneda={moneda} />
              <FilaM label="Cupo disponible" monto={ext.cupoDisponible} moneda={moneda} positive />
              <FilaM label="Cupo utilizado (deuda total estimada)" monto={ext.cupoUtilizado} moneda={moneda} />
            </View>
            <Text style={styles.footNote}>
              Incluye cupo anotado en Saldo y movimientos de la app. Si unifica con el banco, evita duplicar en Saldo
              y en Gastos.
            </Text>

            <Text style={styles.sectionTitle}>3. Detalle de movimientos (periodo de corte)</Text>
            {ext.lineas.length > 0 ? (
              <FilaL
                label="Suma capital del periodo (movimientos que cierran hoy)"
                value={`${formatearNumero(ext.capitalPeriodo)} ${moneda}`}
              />
            ) : null}
            {ext.lineas.length === 0 ? (
              <Text style={[typography.small, { color: colors.textFaint, marginBottom: spacing.sm }]}>
                Sin cargos con fecha de corte en este día (revisa otras fechas o Saldo).
              </Text>
            ) : (
              ext.lineas.map((ln, i) => (
                <View key={i} style={styles.movBlock}>
                  <Text style={styles.movTitle}>{ln.descripcion}</Text>
                  <Text style={styles.movSub}>{ln.cuotaLabel} · {ln.categoria}</Text>
                  <FilaL label="Capital del periodo" value={`${formatearNumero(ln.capitalMes)} ${moneda}`} />
                </View>
              ))
            )}
            <FilaL
              label="Intereses del periodo (estim., tasa E.A.)"
              value={`${formatearNumero(ext.intereses)} ${moneda}`}
            />
            {ext.lineas.length > 0 ? (
              <Text style={[styles.footNote, { marginTop: 4, marginBottom: spacing.sm }]}>
                Con movimientos hoy, el interés se calcula como r. mensual × suma de capital de este corte (igual
                que en Pagos programados de Gastos), no sobre toda la deuda del extracto.
              </Text>
            ) : null}
            <FilaL label="Costos fijos" value={`${formatearNumero(ext.costosFijos)} ${moneda}`} />

            <Text style={styles.sectionTitle}>4. Totales a pagar (obligación)</Text>
            <View style={styles.block}>
              <FilaM label="Pago mínimo (aprox. 3%)" monto={ext.pagoMinimo} moneda={moneda} />
              <FilaM
                label="Pago total (evitar intereses, saldo del periodo)"
                monto={ext.pagoTotalObl}
                moneda={moneda}
                bold
              />
            </View>

            <Text style={styles.sectionTitle}>5. Proyección (ahorro vs plazos)</Text>
            <View style={styles.block}>
              <FilaM label="Costo total proyectado a 6 cuotas" monto={ext.proy6} moneda={moneda} />
              <FilaM label="Costo total proyectado a 3 cuotas" monto={ext.proy3} moneda={moneda} positive />
              <FilaL
                label="Ahorro efectivo (pagar a 3 vs 6 en cuotas, estimación)"
                value={`${formatearNumero(ext.ahorro6vs3)} ${moneda}`}
                valueStyle={{ color: colors.mint, fontWeight: '700' }}
              />
            </View>
            {ext.tasaEA > 0 ? (
              <Text style={styles.footNote}>Tasa E.A. {formatearNumero(ext.tasaEA, 2)}% (Saldo). Proyecciones aprox.</Text>
            ) : (
              <Text style={styles.footNote}>Añade tasa E.A. en Saldo → Tarjeta para ver proyección con interés.</Text>
            )}
          </ScrollView>
        </View>
      </View>
    </Modal>
  );
}

function FilaL({ label, value, valueStyle }) {
  return (
    <View style={styles.filaL}>
      <Text style={styles.filaLlab}>{label}</Text>
      <Text style={[styles.filaLval, valueStyle]}>{value}</Text>
    </View>
  );
}

function FilaM({ label, monto, moneda, positive, bold }) {
  return (
    <View style={styles.filaL}>
      <Text style={styles.filaLlab}>{label}</Text>
      <Text
        style={[
          styles.filaLvalMono,
          positive && { color: colors.mint },
          bold && { fontWeight: '800' },
        ]}
      >
        {formatearNumero(monto)} {moneda}
      </Text>
    </View>
  );
}

const styles = StyleSheet.create({
  root: { flex: 1, justifyContent: 'flex-end' },
  backdrop: { ...StyleSheet.absoluteFillObject, backgroundColor: 'rgba(0,0,0,0.55)' },
  sheet: {
    maxHeight: '92%',
    backgroundColor: colors.bgElevated,
    borderTopLeftRadius: radii.xl,
    borderTopRightRadius: radii.xl,
    borderWidth: 1,
    borderColor: colors.stroke,
    paddingHorizontal: spacing.lg,
  },
  headRow: { flexDirection: 'row', alignItems: 'flex-start', marginBottom: spacing.md },
  scroll: { maxHeight: '100%' },
  sectionTitle: {
    ...typography.label,
    color: colors.accent,
    marginTop: spacing.lg,
    marginBottom: spacing.sm,
  },
  block: {
    borderWidth: 1,
    borderColor: colors.stroke,
    borderRadius: radii.md,
    padding: spacing.md,
    marginBottom: spacing.sm,
  },
  filaL: { flexDirection: 'row', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: spacing.sm },
  filaLlab: { ...typography.small, color: colors.textSecondary, flex: 1, paddingRight: spacing.sm },
  filaLval: { ...typography.body, fontWeight: '600', flexShrink: 0, textAlign: 'right', maxWidth: '58%' },
  filaLvalMono: { ...typography.monoAmount, fontSize: 15, flexShrink: 0, textAlign: 'right' },
  movBlock: { marginBottom: spacing.md, borderLeftWidth: 3, borderLeftColor: colors.accentDeep, paddingLeft: spacing.sm },
  movTitle: { color: colors.text, fontWeight: '700', fontSize: 15 },
  movSub: { color: colors.textFaint, fontSize: 12, marginTop: 2, marginBottom: 4 },
  footNote: { ...typography.small, color: colors.textFaint, lineHeight: 20, marginTop: spacing.xs, marginBottom: spacing.sm },
});
