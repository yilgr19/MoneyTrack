import React, { useState } from 'react';
import { View, Text, StyleSheet, TextInput, TouchableOpacity, Alert, Platform } from 'react-native';
import DateTimePicker from '@react-native-community/datetimepicker';
import { Picker } from '@react-native-picker/picker';
import ScreenWrap from '../components/ScreenWrap';
import UICard from '../components/UICard';
import { PrimaryButton } from '../components/Buttons';
import { useApp } from '../context/AppContext';
import { CUENTAS } from '../lib/finance';
import { colors, spacing, radii, typography } from '../theme';

function pad(n) {
  return String(n).padStart(2, '0');
}

export default function IngresosScreen() {
  const { state, replaceState } = useApp();
  const [cantidad, setCantidad] = useState('');
  const [fecha, setFecha] = useState(new Date());
  const [showPicker, setShowPicker] = useState(false);
  const [origen, setOrigen] = useState('');
  const [nota, setNota] = useState('');

  function guardar() {
    const c = parseFloat(cantidad) || 0;
    if (c <= 0 || !origen) {
      Alert.alert('Datos', 'Cantidad y cuenta obligatorios.');
      return;
    }
    const fechaStr = `${fecha.getFullYear()}-${pad(fecha.getMonth() + 1)}-${pad(fecha.getDate())}`;
    const ing = { cantidad: c, fecha: fechaStr, origen, nota: nota.trim() || null };
    replaceState((s) => ({ ...s, ingresos: [...(s.ingresos || []), ing] }));
    Alert.alert('Listo', 'Ingreso registrado.');
    setCantidad('');
    setNota('');
  }

  return (
    <ScreenWrap includeTopInset={false} contentStyle={{ paddingTop: spacing.xs }}>
      <Text style={typography.label}>Entradas</Text>
      <Text style={typography.hero}>Ingresos</Text>
      <Text style={[typography.subtitle, { marginBottom: spacing.lg }]}>Sueldo, ventas y otros</Text>

      <UICard style={{ marginBottom: 0 }}>
        <Text style={typography.label}>Registro</Text>
        <Text style={styles.lab}>Cantidad</Text>
        <TextInput
          style={styles.input}
          value={cantidad}
          onChangeText={setCantidad}
          keyboardType="decimal-pad"
          placeholderTextColor={colors.textFaint}
        />
        <Text style={styles.lab}>Fecha</Text>
        <TouchableOpacity style={styles.input} onPress={() => setShowPicker(true)}>
          <Text style={{ color: colors.text, fontSize: 16 }}>{fecha.toLocaleDateString('es')}</Text>
        </TouchableOpacity>
        {showPicker && (
          <DateTimePicker
            value={fecha}
            mode="date"
            display={Platform.OS === 'ios' ? 'spinner' : 'default'}
            onChange={(ev, d) => {
              if (Platform.OS !== 'ios') setShowPicker(false);
              if (ev.type === 'dismissed') setShowPicker(false);
              if (d) setFecha(d);
            }}
          />
        )}
        <Text style={styles.lab}>Cuenta destino</Text>
        <View style={styles.pickerWrap}>
          <Picker selectedValue={origen} onValueChange={setOrigen} style={{ color: colors.text }}>
            <Picker.Item label="Selecciona…" value="" />
            {CUENTAS.map((c) => (
              <Picker.Item key={c.id} label={c.nombre} value={c.id} />
            ))}
          </Picker>
        </View>
        <Text style={styles.lab}>Nota (opcional)</Text>
        <TextInput style={styles.input} value={nota} onChangeText={setNota} placeholderTextColor={colors.textFaint} />
        <PrimaryButton title="Guardar ingreso" onPress={guardar} style={{ marginTop: spacing.lg }} />
      </UICard>
    </ScreenWrap>
  );
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
});
