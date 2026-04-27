import React, { useEffect, useMemo, useState } from 'react';
import {
  View,
  Text,
  StyleSheet,
  TouchableOpacity,
  Alert,
  Modal,
  Pressable,
  Platform,
} from 'react-native';
import { Picker } from '@react-native-picker/picker';
import { Ionicons } from '@expo/vector-icons';
import ScreenWrap from '../components/ScreenWrap';
import ExtractoBancarioModal from '../components/ExtractoBancarioModal';
import { PrimaryButton, GhostButton } from '../components/Buttons';
import UICard from '../components/UICard';
import { useApp } from '../context/AppContext';
import { construirExtractoBancarioTarjeta, formatearNumero, refUltimaHoraDiaEnMes } from '../lib/finance';
import { colors, spacing, radii, typography, iconSemantic } from '../theme';

const MESES_ATRAS = 36;

function generarIdExtracto() {
  return `ext-hist-${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
}

function opcionesMeses() {
  const out = [];
  const now = new Date();
  for (let i = 0; i < MESES_ATRAS; i++) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    const y = d.getFullYear();
    const m = d.getMonth() + 1;
    const value = `${y}-${String(m).padStart(2, '0')}`;
    const label = d.toLocaleDateString('es', { month: 'long', year: 'numeric' });
    out.push({ value, label: label.charAt(0).toUpperCase() + label.slice(1) });
  }
  return out;
}

function formatearEtiquetaMes(ym) {
  const [a, b] = String(ym).split('-');
  const y = parseInt(a, 10);
  const m0 = parseInt(b, 10) - 1;
  if (!Number.isFinite(y) || m0 < 0) return String(ym);
  return new Date(y, m0, 1).toLocaleDateString('es', { month: 'long', year: 'numeric' });
}

function nNumSeguro(v) {
  const x = Number(v);
  return Number.isFinite(x) ? x : 0;
}

export default function ExtractosTarjetasScreen() {
  const { state, replaceState } = useApp();
  const moneda = (state?.moneda && String(state.moneda).trim()) || '';
  const tarjetas = state?.tarjetasCredito || [];
  const mesesOpt = useMemo(() => opcionesMeses(), []);

  const [modalAdd, setModalAdd] = useState(false);
  const [mesSel, setMesSel] = useState(mesesOpt[0]?.value || '');
  const [tarjetaIdSel, setTarjetaIdSel] = useState('');

  const [verExtracto, setVerExtracto] = useState(null);

  const tarjetaDeExtractoAbierto = useMemo(() => {
    if (!verExtracto) return null;
    return (state?.tarjetasCredito || []).find((x) => x && String(x.id) === String(verExtracto.tarjetaId)) || null;
  }, [verExtracto, state?.tarjetasCredito]);

  const listaOrdenada = useMemo(() => {
    const arr = [...(state?.extractosTarjetasHistorial || [])];
    arr.sort((a, b) => {
      const m = (b.mes || '').localeCompare(a.mes || '');
      if (m !== 0) return m;
      return (a.nombreEntidad || '').localeCompare(b.nombreEntidad || '', 'es');
    });
    return arr;
  }, [state?.extractosTarjetasHistorial]);

  /** Deuda mostrada en fila: recalculada con movimientos actuales (mismo criterio que al abrir el extracto). */
  const filasVista = useMemo(() => {
    return listaOrdenada.map((it) => {
      const t = tarjetas.find((x) => x && String(x.id) === String(it.tarjetaId));
      let cupoVivo = nNumSeguro(it.snapshot?.cupoUtilizado);
      if (t && it.mes) {
        try {
          const refD = refUltimaHoraDiaEnMes(it.mes);
          const ex = construirExtractoBancarioTarjeta(t, state || {}, refD);
          cupoVivo = nNumSeguro(ex.cupoUtilizado);
        } catch {
          /* no-op */
        }
      }
      return { it, cupoVivo };
    });
  }, [listaOrdenada, tarjetas, state]);

  useEffect(() => {
    if (tarjetas.length && !tarjetaIdSel) {
      setTarjetaIdSel(String(tarjetas[0].id));
    }
  }, [tarjetas, tarjetaIdSel]);

  function guardarExtracto() {
    if (!tarjetaIdSel) {
      Alert.alert('Tarjeta', 'Elige una tarjeta en Saldo primero.');
      return;
    }
    const t = tarjetas.find((x) => x && String(x.id) === String(tarjetaIdSel));
    if (!t) {
      Alert.alert('Tarjeta', 'No se encontró la tarjeta seleccionada.');
      return;
    }
    if (!mesSel) {
      Alert.alert('Mes', 'Elige un mes de referencia.');
      return;
    }
    const refD = refUltimaHoraDiaEnMes(mesSel);
    const snapshot = construirExtractoBancarioTarjeta(t, state || {}, refD);
    const nombreEntidad = String(t.nombreEntidad || 'Tarjeta').trim() || 'Tarjeta';
    const nuevo = {
      id: generarIdExtracto(),
      tarjetaId: String(t.id),
      mes: mesSel,
      nombreEntidad,
      refISO: refD.toISOString(),
      creadoEn: new Date().toISOString(),
      snapshot,
    };
    replaceState((s) => {
      const prev = s.extractosTarjetasHistorial || [];
      const sinDup = prev.filter(
        (x) => !((x.mes === nuevo.mes && String(x.tarjetaId) === String(nuevo.tarjetaId)))
      );
      return { ...s, extractosTarjetasHistorial: [nuevo, ...sinDup] };
    });
    setModalAdd(false);
    Alert.alert('Guardado', 'El extracto quedó en el historial para ese mes y tarjeta.');
  }

  function eliminarItem(item) {
    Alert.alert('Quitar extracto', '¿Eliminar esta copia del historial?', [
      { text: 'Cancelar', style: 'cancel' },
      {
        text: 'Eliminar',
        style: 'destructive',
        onPress: () => {
          replaceState((s) => ({
            ...s,
            extractosTarjetasHistorial: (s.extractosTarjetasHistorial || []).filter((x) => x.id !== item.id),
          }));
        },
      },
    ]);
  }

  return (
    <ScreenWrap contentStyle={{ paddingTop: spacing.xs }}>
      <Text style={styles.hint}>
        Cada registro archiva un mes y tarjeta. La lista y el detalle se recalculan con tus movimientos y cupos
        actuales; la fecha “guardado” es solo la copia de respaldo. Puedes volver a guardar el mismo mes para
        sustituir el archivo.
      </Text>

      {tarjetas.length === 0 ? (
        <UICard>
          <Text style={typography.body}>
            Añade tarjetas de crédito en Saldo → Tarjeta para generar y archivar extractos.
          </Text>
        </UICard>
      ) : (
        <PrimaryButton title="Añadir extracto al historial" onPress={() => setModalAdd(true)} style={{ marginBottom: spacing.md }} />
      )}

      <Text style={typography.label} accessibilityRole="header">
        Historial (más reciente primero)
      </Text>
      {filasVista.length === 0 ? (
        <UICard>
          <Text style={[typography.body, { color: colors.textMuted }]}>Aún no hay extractos guardados.</Text>
        </UICard>
      ) : (
        filasVista.map(({ it, cupoVivo }) => (
          <View key={it.id} style={styles.rowCard}>
            <TouchableOpacity
              style={styles.rowMain}
              onPress={() => setVerExtracto(it)}
              onLongPress={() => eliminarItem(it)}
              activeOpacity={0.75}
            >
              <View
                style={[
                  styles.rowIcon,
                  { backgroundColor: iconSemantic.moreMenu.ExtractosTarjetas.bg },
                ]}
              >
                <Ionicons
                  name="receipt-outline"
                  size={24}
                  color={iconSemantic.moreMenu.ExtractosTarjetas.fg}
                />
              </View>
              <View style={styles.rowBody}>
                <Text style={styles.rowTitle}>
                  {formatearEtiquetaMes(it.mes)} · {it.nombreEntidad || 'Tarjeta'}
                </Text>
                <Text style={styles.rowSub} numberOfLines={2}>
                  Deuda est. (al día) {formatearNumero(cupoVivo, 0)} {moneda}
                  {it.creadoEn
                    ? ` · guardado ${new Date(it.creadoEn).toLocaleDateString('es', { day: 'numeric', month: 'short' })}`
                    : ''}
                </Text>
                <Text style={styles.rowHint}>Mantén pulsado para eliminar</Text>
              </View>
              <Ionicons name="chevron-forward" size={20} color={colors.textFaint} />
            </TouchableOpacity>
          </View>
        ))
      )}

      <Modal visible={modalAdd} animationType="slide" transparent onRequestClose={() => setModalAdd(false)}>
        <View style={styles.mOverlay}>
          <Pressable style={StyleSheet.absoluteFill} onPress={() => setModalAdd(false)} />
          <View style={styles.mSheet}>
            <Text style={typography.title}>Nuevo extracto archivado</Text>
            <Text style={styles.mHint}>
              Se usa el último día del mes como fecha de corte de referencia para el cálculo.
            </Text>
            <Text style={styles.modalLab}>Mes (periodo de referencia)</Text>
            <View style={styles.pickerWrap}>
              <Picker
                selectedValue={mesSel}
                onValueChange={setMesSel}
                style={{ color: colors.text }}
                dropdownIconColor={colors.text}
              >
                {mesesOpt.map((m) => (
                  <Picker.Item key={m.value} label={m.label} value={m.value} />
                ))}
              </Picker>
            </View>
            <Text style={styles.modalLab}>Tarjeta</Text>
            <View style={styles.pickerWrap}>
              <Picker
                selectedValue={tarjetaIdSel}
                onValueChange={setTarjetaIdSel}
                style={{ color: colors.text }}
                dropdownIconColor={colors.text}
              >
                {tarjetas.map((t) => (
                  <Picker.Item
                    key={t.id}
                    label={String(t.nombreEntidad || 'Tarjeta').trim() || 'Tarjeta'}
                    value={t.id}
                  />
                ))}
              </Picker>
            </View>
            <PrimaryButton title="Guardar en historial" onPress={guardarExtracto} style={{ marginTop: spacing.lg }} />
            <GhostButton title="Cerrar" onPress={() => setModalAdd(false)} style={{ marginTop: spacing.sm }} />
          </View>
        </View>
      </Modal>

      <ExtractoBancarioModal
        visible={verExtracto != null}
        onClose={() => setVerExtracto(null)}
        state={state || {}}
        tarjeta={tarjetaDeExtractoAbierto}
        moneda={moneda}
        extSnapshot={!tarjetaDeExtractoAbierto && verExtracto ? verExtracto.snapshot : null}
        refMesHistorial={verExtracto?.mes}
        historialCreadoEn={verExtracto?.creadoEn}
      />
    </ScreenWrap>
  );
}

const styles = StyleSheet.create({
  hint: { ...typography.small, color: colors.textSecondary, marginBottom: spacing.lg, lineHeight: 20 },
  rowCard: {
    backgroundColor: colors.surface,
    borderRadius: radii.lg,
    borderWidth: 1,
    borderColor: colors.stroke,
    marginBottom: spacing.sm,
  },
  rowMain: { flexDirection: 'row', alignItems: 'center', padding: spacing.md },
  rowIcon: {
    width: 44,
    height: 44,
    borderRadius: radii.md,
    backgroundColor: colors.surfaceHighlight,
    alignItems: 'center',
    justifyContent: 'center',
    marginRight: spacing.md,
    borderWidth: 1,
    borderColor: colors.stroke,
  },
  rowBody: { flex: 1, minWidth: 0 },
  rowTitle: { color: colors.text, fontSize: 16, fontWeight: '700' },
  rowSub: { color: colors.textMuted, fontSize: 13, marginTop: 4, lineHeight: 19 },
  rowHint: { color: colors.textFaint, fontSize: 11, marginTop: 6 },
  mOverlay: {
    flex: 1,
    backgroundColor: 'rgba(0,0,0,0.5)',
    justifyContent: 'flex-end',
  },
  mSheet: {
    backgroundColor: colors.bgElevated,
    borderTopLeftRadius: radii.xl,
    borderTopRightRadius: radii.xl,
    padding: spacing.lg,
    paddingBottom: Platform.OS === 'ios' ? 32 : spacing.lg,
    borderWidth: 1,
    borderColor: colors.stroke,
  },
  mHint: { ...typography.small, color: colors.textSecondary, marginBottom: spacing.md, lineHeight: 20 },
  modalLab: { ...typography.label, marginTop: spacing.md, marginBottom: spacing.xs },
  pickerWrap: {
    borderWidth: 1,
    borderColor: colors.stroke,
    borderRadius: radii.md,
    overflow: 'hidden',
    backgroundColor: 'rgba(0,0,0,0.12)',
  },
});
