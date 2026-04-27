import React, { useMemo, useState } from 'react';
import { View, Text, StyleSheet, TextInput, TouchableOpacity, Alert, Platform, Switch } from 'react-native';
import DateTimePicker from '@react-native-community/datetimepicker';
import { Picker } from '@react-native-picker/picker';
import ScreenWrap from '../components/ScreenWrap';
import UICard from '../components/UICard';
import { PrimaryButton } from '../components/Buttons';
import { useApp } from '../context/AppContext';
import {
  CUENTAS,
  formatearNumero,
  normalizarCategoria,
  generarIdPagoProgramado,
  fechaALocalISO,
  parseFechaHoraLocal,
} from '../lib/finance';
import { colors, spacing, radii, typography } from '../theme';

function pad(n) {
  return String(n).padStart(2, '0');
}

export default function PagosScreen() {
  const { state, replaceState } = useApp();
  const moneda = state.moneda || '';
  const categorias = useMemo(() => (state.categorias || []).map(normalizarCategoria), [state.categorias]);

  const [concepto, setConcepto] = useState('');
  const [monto, setMonto] = useState('');
  const [frecuencia, setFrecuencia] = useState('mensual');
  const [diaPago, setDiaPago] = useState(1);
  const [cuenta, setCuenta] = useState('');
  const [categoria, setCategoria] = useState('');
  const [fechaInicio, setFechaInicio] = useState(new Date());
  const [showPicker, setShowPicker] = useState(false);
  const [activo, setActivo] = useState(true);
  const [nota, setNota] = useState('');

  const diasOpciones = useMemo(() => {
    if (frecuencia === 'mensual') {
      return Array.from({ length: 28 }, (_, i) => i + 1);
    }
    if (frecuencia === 'quincenal') {
      return [1, 15];
    }
    return [1];
  }, [frecuencia]);

  function guardar() {
    const m = parseFloat(monto) || 0;
    if (!concepto.trim() || m <= 0 || !cuenta || !categoria) {
      Alert.alert('Datos', 'Completa concepto, monto, cuenta y categoría.');
      return;
    }
    const fechaStr = fechaALocalISO(fechaInicio);
    const diaFromFecha = Math.min(28, fechaInicio.getDate());
    const p = {
      id: generarIdPagoProgramado(),
      concepto: concepto.trim(),
      monto: m,
      frecuencia,
      fechaInicio: fechaStr,
      diaPago: frecuencia === 'semanal' ? fechaInicio.getDate() : frecuencia === 'mensual' ? diaFromFecha : diaPago,
      cuenta,
      categoria,
      activo,
      nota: nota.trim() || '',
    };
    replaceState((s) => ({ ...s, pagosProgramados: [...(s.pagosProgramados || []), p] }));
    Alert.alert('Listo', 'Guardado. Regístralo desde Gastos cuando corresponda.');
    setConcepto('');
    setMonto('');
    setNota('');
  }

  function eliminar(id) {
    Alert.alert('Eliminar', '¿Quitar este pago programado?', [
      { text: 'Cancelar', style: 'cancel' },
      {
        text: 'Eliminar',
        style: 'destructive',
        onPress: () =>
          replaceState((s) => ({
            ...s,
            pagosProgramados: (s.pagosProgramados || []).filter((x) => x.id !== id),
          })),
      },
    ]);
  }

  return (
    <ScreenWrap includeTopInset={false} contentStyle={{ paddingTop: spacing.xs }}>
      <Text style={typography.label}>Automatiza</Text>
      <Text style={typography.hero}>Pagos programados</Text>
      <Text style={[typography.subtitle, { marginBottom: spacing.lg }]}>
        Recordatorios; confirma en Gastos
      </Text>

      <UICard>
        <Text style={typography.label}>Nuevo</Text>
        <Lab>Concepto</Lab>
        <TextInput style={styles.input} value={concepto} onChangeText={setConcepto} placeholderTextColor={colors.textFaint} />
        <Lab>Monto</Lab>
        <TextInput style={styles.input} value={monto} onChangeText={setMonto} keyboardType="decimal-pad" placeholderTextColor={colors.textFaint} />
        <Lab>Frecuencia</Lab>
        <View style={styles.pickerWrap}>
          <Picker selectedValue={frecuencia} onValueChange={setFrecuencia} style={{ color: colors.text }}>
            <Picker.Item label="Mensual" value="mensual" />
            <Picker.Item label="Quincenal" value="quincenal" />
            <Picker.Item label="Semanal" value="semanal" />
          </Picker>
        </View>
        {(frecuencia === 'mensual' || frecuencia === 'quincenal') && (
          <>
            <Lab>Día de pago</Lab>
            <View style={styles.pickerWrap}>
              <Picker selectedValue={diaPago} onValueChange={setDiaPago} style={{ color: colors.text }}>
                {diasOpciones.map((d) => (
                  <Picker.Item key={d} label={String(d)} value={d} />
                ))}
              </Picker>
            </View>
          </>
        )}
        <Lab>Cuenta</Lab>
        <View style={styles.pickerWrap}>
          <Picker selectedValue={cuenta} onValueChange={setCuenta} style={{ color: colors.text }}>
            <Picker.Item label="Selecciona…" value="" />
            {CUENTAS.map((c) => (
              <Picker.Item key={c.id} label={c.nombre} value={c.id} />
            ))}
          </Picker>
        </View>
        <Lab>Categoría</Lab>
        <View style={styles.pickerWrap}>
          <Picker selectedValue={categoria} onValueChange={setCategoria} style={{ color: colors.text }}>
            <Picker.Item label="Selecciona…" value="" />
            {categorias.map((c) => (
              <Picker.Item key={c.nombre} label={`${c.icono} ${c.nombre}`} value={c.nombre} />
            ))}
          </Picker>
        </View>
        <Lab>Fecha de inicio (mensual/quincenal: define el día del vencimiento en el mes)</Lab>
        <TouchableOpacity style={styles.input} onPress={() => setShowPicker(true)}>
          <Text style={{ color: colors.text, fontSize: 16 }}>
            {fechaInicio.toLocaleDateString('es', { dateStyle: 'short' })}
          </Text>
        </TouchableOpacity>
        {showPicker && (
          <DateTimePicker
            value={fechaInicio}
            mode="date"
            display={Platform.OS === 'ios' ? 'spinner' : 'default'}
            onChange={(ev, d) => {
              if (Platform.OS !== 'ios') setShowPicker(false);
              if (ev.type === 'dismissed') setShowPicker(false);
              if (d) setFechaInicio(d);
            }}
          />
        )}
        <View style={styles.rowSwitch}>
          <Text style={[typography.body, { flex: 1, minWidth: 0, paddingRight: spacing.sm }]}>Activo</Text>
          <Switch
            value={activo}
            onValueChange={setActivo}
            trackColor={{ false: colors.barTrack, true: colors.accentDeep }}
            thumbColor={activo ? colors.text : colors.textMuted}
          />
        </View>
        <Lab>Nota (opcional)</Lab>
        <TextInput style={styles.input} value={nota} onChangeText={setNota} placeholderTextColor={colors.textFaint} />
        <PrimaryButton title="Guardar pago programado" onPress={guardar} style={{ marginTop: spacing.md }} />
      </UICard>

      <UICard style={{ marginBottom: 0 }}>
        <Text style={typography.label}>Lista</Text>
        {(state.pagosProgramados || []).length === 0 ? (
          <Text style={typography.small}>Sin pagos programados.</Text>
        ) : (
          (state.pagosProgramados || []).map((p) => (
            <View key={p.id} style={styles.item}>
              <Text style={styles.itemTit}>{p.concepto}</Text>
              <Text style={typography.small}>
                {formatearNumero(p.monto)} {moneda} · {p.frecuencia} ·{' '}
                {p.activo !== false ? 'activo' : 'inactivo'}
                {p.fechaInicio
                  ? ` · ${parseFechaHoraLocal(p.fechaInicio)?.toLocaleDateString('es', { dateStyle: 'short' }) ?? p.fechaInicio}`
                  : ''}
              </Text>
              <TouchableOpacity onPress={() => eliminar(p.id)} hitSlop={8}>
                <Text style={styles.del}>Eliminar</Text>
              </TouchableOpacity>
            </View>
          ))
        )}
      </UICard>
    </ScreenWrap>
  );
}

function Lab({ children }) {
  return <Text style={styles.lab}>{children}</Text>;
}

const styles = StyleSheet.create({
  lab: {
    ...typography.label,
    marginTop: spacing.md,
    marginBottom: spacing.xs,
    color: colors.textMuted,
    letterSpacing: 0.8,
  },
  input: {
    borderWidth: 1,
    borderColor: colors.stroke,
    borderRadius: radii.md,
    padding: spacing.md,
    color: colors.text,
    fontSize: 16,
    backgroundColor: 'rgba(0,0,0,0.18)',
  },
  pickerWrap: {
    borderWidth: 1,
    borderColor: colors.stroke,
    borderRadius: radii.md,
    overflow: 'hidden',
    backgroundColor: 'rgba(0,0,0,0.12)',
  },
  rowSwitch: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    marginTop: spacing.lg,
    paddingVertical: spacing.sm,
    gap: spacing.md,
  },
  item: {
    marginTop: spacing.md,
    paddingBottom: spacing.md,
    borderBottomWidth: 1,
    borderBottomColor: colors.stroke,
  },
  itemTit: {
    color: colors.text,
    fontWeight: '700',
    fontSize: 15,
    marginBottom: 4,
    flexShrink: 1,
  },
  del: { color: colors.danger, marginTop: 6, fontWeight: '600', fontSize: 13 },
});
