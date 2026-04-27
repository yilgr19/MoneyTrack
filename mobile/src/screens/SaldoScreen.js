import React, { useState, useEffect } from 'react';
import { View, Text, StyleSheet, TextInput, Alert } from 'react-native';
import { Picker } from '@react-native-picker/picker';
import ScreenWrap from '../components/ScreenWrap';
import UICard from '../components/UICard';
import { PrimaryButton } from '../components/Buttons';
import { useApp } from '../context/AppContext';
import { CUENTAS, formatearNumero, calcularSaldosPorCuenta } from '../lib/finance';
import { emptySaldosCuentas } from '../lib/storage';
import { colors, spacing, radii, typography, layoutStyles } from '../theme';

const MONEDAS = [
  { value: '', label: 'Selecciona…' },
  { value: 'USD', label: 'USD' },
  { value: 'EUR', label: 'EUR' },
  { value: 'MXN', label: 'MXN' },
  { value: 'COP', label: 'COP' },
  { value: 'ARS', label: 'ARS' },
  { value: 'CLP', label: 'CLP' },
  { value: 'PEN', label: 'PEN' },
  { value: 'GBP', label: 'GBP' },
  { value: 'JPY', label: 'JPY' },
  { value: 'BRL', label: 'BRL' },
  { value: 'CAD', label: 'CAD' },
  { value: 'GTQ', label: 'GTQ' },
];

export default function SaldoScreen() {
  const { state, replaceState } = useApp();
  const [moneda, setMoneda] = useState(state.moneda || '');
  const [saldos, setSaldos] = useState(() => ({ ...emptySaldosCuentas(), ...state.saldosCuentas }));
  const [limiteTc, setLimiteTc] = useState(String(state.limiteTarjetaCredito || 0));
  const [presupuesto, setPresupuesto] = useState(String(state.presupuestoMensual || 0));
  const [nota, setNota] = useState(state.saldoInicialNota || '');

  useEffect(() => {
    setMoneda(state.moneda || '');
    setSaldos({ ...emptySaldosCuentas(), ...state.saldosCuentas });
    setLimiteTc(String(state.limiteTarjetaCredito || 0));
    setPresupuesto(String(state.presupuestoMensual || 0));
    setNota(state.saldoInicialNota || '');
  }, [state.moneda, state.saldosCuentas, state.limiteTarjetaCredito, state.presupuestoMensual, state.saldoInicialNota]);

  const saldosActuales = calcularSaldosPorCuenta(state);

  function guardar() {
    if (!moneda) {
      Alert.alert('Moneda', 'Selecciona un tipo de moneda.');
      return;
    }
    const sc = { ...emptySaldosCuentas() };
    CUENTAS.forEach((c) => {
      sc[c.id] = parseFloat(saldos[c.id]) || 0;
    });
    replaceState((s) => ({
      ...s,
      moneda,
      saldosCuentas: sc,
      limiteTarjetaCredito: parseFloat(limiteTc) || 0,
      presupuestoMensual: parseFloat(presupuesto) || 0,
      saldoInicialNota: nota.trim(),
    }));
    Alert.alert('Guardado', 'Saldo inicial actualizado.');
  }

  return (
    <ScreenWrap contentStyle={{ paddingTop: spacing.xs }}>
      <Text style={typography.label}>Cuentas</Text>
      <Text style={typography.hero}>Saldo inicial</Text>
      <Text style={[typography.subtitle, { marginBottom: spacing.lg }]}>
        Montos por cuenta y moneda base
      </Text>

      <UICard>
        <Text style={typography.label}>Configuración</Text>
        <Text style={styles.lab}>Moneda</Text>
        <View style={styles.pickerWrap}>
          <Picker selectedValue={moneda} onValueChange={setMoneda} style={{ color: colors.text }}>
            {MONEDAS.map((m) => (
              <Picker.Item key={m.value || 'x'} label={m.label} value={m.value} />
            ))}
          </Picker>
        </View>

        {CUENTAS.map((c) => (
          <View key={c.id}>
            <Text style={styles.lab}>{c.nombre}</Text>
            <TextInput
              style={styles.input}
              keyboardType="decimal-pad"
              value={String(saldos[c.id] ?? '')}
              onChangeText={(t) => setSaldos((prev) => ({ ...prev, [c.id]: t }))}
              placeholder="0.00"
              placeholderTextColor={colors.textFaint}
            />
            {c.id === 'tarjetaCredito' && (
              <>
                <Text style={styles.lab}>Límite tarjeta (alertas)</Text>
                <TextInput
                  style={styles.input}
                  keyboardType="decimal-pad"
                  value={limiteTc}
                  onChangeText={setLimiteTc}
                  placeholderTextColor={colors.textFaint}
                />
              </>
            )}
          </View>
        ))}

        <Text style={styles.lab}>Presupuesto mensual (opcional)</Text>
        <TextInput
          style={styles.input}
          keyboardType="decimal-pad"
          value={presupuesto}
          onChangeText={setPresupuesto}
          placeholderTextColor={colors.textFaint}
        />

        <Text style={styles.lab}>Nota (opcional)</Text>
        <TextInput
          style={styles.input}
          value={nota}
          onChangeText={setNota}
          placeholderTextColor={colors.textFaint}
        />

        <PrimaryButton title="Guardar saldo inicial" onPress={guardar} style={{ marginTop: spacing.lg }} />
      </UICard>

      <UICard style={{ marginBottom: 0 }}>
        <Text style={typography.label}>Vista previa</Text>
        <Text style={[typography.small, { marginBottom: spacing.md }]}>
          Saldos calculados con movimientos actuales
        </Text>
        {CUENTAS.map((c) => (
          <View key={c.id} style={layoutStyles.rowBetween}>
            <Text style={[typography.body, layoutStyles.rowLabel]}>{c.nombre}</Text>
            <Text style={[typography.monoAmount, layoutStyles.rowValue]}>
              {formatearNumero(saldosActuales[c.id] || 0)} {state.moneda}
            </Text>
          </View>
        ))}
        <Text style={styles.total} numberOfLines={2} adjustsFontSizeToFit minimumFontScale={0.75}>
          Total · {formatearNumero(saldosActuales.total || 0)} {state.moneda}
        </Text>
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
  total: {
    fontSize: 18,
    fontWeight: '700',
    color: colors.mint,
    marginTop: spacing.md,
    letterSpacing: -0.3,
    maxWidth: '100%',
  },
});
