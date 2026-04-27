import React, { useMemo, useState } from 'react';
import { View, Text, StyleSheet, TextInput, TouchableOpacity, Alert, Platform } from 'react-native';
import DateTimePicker from '@react-native-community/datetimepicker';
import { Picker } from '@react-native-picker/picker';
import ScreenWrap from '../components/ScreenWrap';
import UICard from '../components/UICard';
import { PrimaryButton } from '../components/Buttons';
import { useApp } from '../context/AppContext';
import {
  formatearNumero,
  CUENTAS,
  calcularSaldosPorCuenta,
  normalizarOrigenCuenta,
  normalizarCategoria,
  montoGastoAfectaSaldo,
  pagoDebeMostrarseParaPagar,
} from '../lib/finance';
import { colors, spacing, radii, typography } from '../theme';

const CUOTAS_OPTS = [1, 2, 3, 6, 12, 24];

function pad(n) {
  return String(n).padStart(2, '0');
}

export default function GastosScreen() {
  const { state, replaceState } = useApp();
  const moneda = state.moneda || '';

  const [nombre, setNombre] = useState('');
  const [cantidad, setCantidad] = useState('');
  const [fecha, setFecha] = useState(new Date());
  const [showPicker, setShowPicker] = useState(false);
  const [categoria, setCategoria] = useState('');
  const [origen, setOrigen] = useState('');
  const [cuotas, setCuotas] = useState(1);
  const [nota, setNota] = useState('');
  const [pagoProgramadoEnUso, setPagoProgramadoEnUso] = useState(null);

  const categorias = useMemo(
    () => (state.categorias || []).map(normalizarCategoria),
    [state.categorias]
  );

  const saldos = useMemo(() => calcularSaldosPorCuenta(state), [state]);
  const cantNum = parseFloat(cantidad) || 0;
  const cuotaMensualTc = cantNum > 0 && cuotas > 0 ? cantNum / cuotas : cantNum;

  const cuentasDisponibles = useMemo(() => {
    const list = [];
    CUENTAS.forEach((c) => {
      let saldo = saldos[c.id] || 0;
      let puede = false;
      if (c.id === 'tarjetaCredito') {
        saldo = Math.max(0, saldos.tarjetaCredito || 0);
        puede = saldo > 0 && (cantNum === 0 || saldo >= cuotaMensualTc);
      } else {
        puede = saldo > 0 && (cantNum === 0 || saldo >= cantNum);
      }
      if (puede) list.push({ ...c, saldo });
    });
    return list;
  }, [saldos, cantNum, cuotaMensualTc]);

  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);
  const pagosPendientes = (state.pagosProgramados || []).filter(
    (p) => p.activo !== false && pagoDebeMostrarseParaPagar(p, hoy)
  );

  function aplicarPagoProgramado(p) {
    setNombre(p.concepto || '');
    setCantidad(String(p.monto ?? ''));
    setCategoria(p.categoria || (categorias[0]?.nombre ?? ''));
    setOrigen(normalizarOrigenCuenta(p.cuenta) || p.cuenta || '');
    setCuotas(1);
    setNota(p.nota || '');
    setPagoProgramadoEnUso(p.id);
  }

  function onSubmit() {
    if (!nombre.trim() || cantNum <= 0 || !categoria || !origen) {
      Alert.alert('Datos incompletos', 'Nombre, cantidad, categoría y cuenta son obligatorios.');
      return;
    }

    const saldosAct = calcularSaldosPorCuenta(state);
    const cuotasVal = origen === 'tarjetaCredito' ? cuotas : 1;
    const cuotaMensualVal = origen === 'tarjetaCredito' ? cantNum / cuotasVal : cantNum;
    const saldoOrigen = Math.max(0, saldosAct[origen] || 0);
    const montoAValidar = origen === 'tarjetaCredito' ? cuotaMensualVal : cantNum;
    const saldoTotal = saldosAct.total || 0;

    if (origen !== 'tarjetaCredito' && cantNum > saldoTotal) {
      Alert.alert('Saldo', 'No hay saldo suficiente en total.');
      return;
    }
    if (montoAValidar > saldoOrigen) {
      Alert.alert('Saldo', 'No hay suficiente saldo en la cuenta seleccionada.');
      return;
    }

    const catObj = categorias.find((c) => c.nombre === categoria);
    if (catObj && catObj.limite) {
      const lim = parseFloat(catObj.limite);
      const ah = new Date();
      const gastosCategoria = (state.gastos || []).filter((g) => {
        const d = new Date(g.fecha);
        return g.categoria === categoria && d.getMonth() === ah.getMonth() && d.getFullYear() === ah.getFullYear();
      });
      const gastadoMes = gastosCategoria.reduce((s, g) => s + montoGastoAfectaSaldo(g), 0);
      if (gastadoMes + montoAValidar > lim) {
        Alert.alert('Límite categoría', 'Este gasto supera el límite mensual de la categoría. ¿Continuar?', [
          { text: 'Cancelar', style: 'cancel' },
          { text: 'Sí', onPress: () => guardarGasto(cuotasVal, cuotaMensualVal) },
        ]);
        return;
      }
    }

    guardarGasto(cuotasVal, cuotaMensualVal);
  }

  function guardarGasto(cuotasVal, cuotaMensualVal) {
    const fechaStr = `${fecha.getFullYear()}-${pad(fecha.getMonth() + 1)}-${pad(fecha.getDate())}T${pad(fecha.getHours())}:${pad(fecha.getMinutes())}:00`;

    const nuevo = {
      nombre: nombre.trim(),
      cantidad: cantNum,
      fecha: fechaStr,
      categoria,
      origen,
      nota: nota.trim() || null,
      cuotas: origen === 'tarjetaCredito' ? cuotasVal : 1,
      cuotaMensual: origen === 'tarjetaCredito' ? cuotaMensualVal : cantNum,
    };

    replaceState((s) => {
      let gastos = [...(s.gastos || []), nuevo];
      let pagos = [...(s.pagosProgramados || [])];

      if (origen === 'tarjetaCredito' && cuotasVal > 1) {
        const d = new Date(fecha);
        const diaCompra = Math.min(28, d.getDate());
        const año = d.getFullYear();
        const mes = d.getMonth();
        for (let i = 0; i < cuotasVal - 1; i++) {
          const nextDate = new Date(año, mes + i + 1, diaCompra);
          const fechaCuota = nextDate.toISOString().slice(0, 10);
          pagos.push({
            id: `cuota-${Date.now()}-${i}-${Math.random().toString(36).slice(2, 6)}`,
            concepto: `${nombre.trim()} - Cuota ${i + 2} de ${cuotasVal}`,
            monto: cuotaMensualVal,
            frecuencia: 'unico',
            fechaInicio: fechaCuota,
            diaPago: nextDate.getDate(),
            cuenta: origen,
            categoria,
            activo: true,
            nota: `${nota || ''}${nota ? ' | ' : ''}Cuota diferida automática`,
            esCuotaDiferida: true,
          });
        }
      }

      if (pagoProgramadoEnUso) {
        pagos = pagos.filter((p) => p.id !== pagoProgramadoEnUso);
      }

      return { ...s, gastos, pagosProgramados: pagos };
    });

    Alert.alert('Listo', 'Gasto registrado.');
    setNombre('');
    setCantidad('');
    setNota('');
    setCuotas(1);
    setPagoProgramadoEnUso(null);
    setFecha(new Date());
  }

  return (
    <ScreenWrap contentStyle={{ paddingTop: spacing.xs }}>
      <Text style={typography.label}>Movimientos</Text>
      <Text style={typography.hero}>Registrar gasto</Text>
      <Text style={[typography.subtitle, { marginBottom: spacing.lg }]}>
        Registra cada salida de dinero
      </Text>

      {pagosPendientes.length > 0 && (
        <UICard accent>
          <Text style={typography.label}>Pagos programados</Text>
          {pagosPendientes.map((p) => (
            <TouchableOpacity key={p.id} style={styles.pagoRow} onPress={() => aplicarPagoProgramado(p)}>
              <Text style={[typography.body, styles.pagoConcepto]}>
                {p.concepto} — {formatearNumero(p.monto)} {moneda}
              </Text>
              <Text style={styles.link}>Usar en formulario →</Text>
            </TouchableOpacity>
          ))}
        </UICard>
      )}

      <UICard style={{ marginBottom: 0 }}>
        <Text style={typography.label}>Detalle</Text>

        <FieldLabel>Nombre</FieldLabel>
        <TextInput
          style={styles.input}
          value={nombre}
          onChangeText={setNombre}
          placeholder="Ej: Supermercado"
          placeholderTextColor={colors.textFaint}
        />

        <FieldLabel>Cantidad</FieldLabel>
        <TextInput
          style={styles.input}
          value={cantidad}
          onChangeText={setCantidad}
          keyboardType="decimal-pad"
          placeholder="0.00"
          placeholderTextColor={colors.textFaint}
        />

        <FieldLabel>Fecha y hora</FieldLabel>
        <TouchableOpacity style={styles.input} onPress={() => setShowPicker(true)}>
          <Text style={{ color: colors.text, fontSize: 16 }}>{fecha.toLocaleString('es')}</Text>
        </TouchableOpacity>
        {showPicker && (
          <DateTimePicker
            value={fecha}
            mode="datetime"
            display={Platform.OS === 'ios' ? 'spinner' : 'default'}
            onChange={(ev, d) => {
              if (Platform.OS !== 'ios') setShowPicker(false);
              if (ev.type === 'dismissed') setShowPicker(false);
              if (d) setFecha(d);
            }}
          />
        )}

        <FieldLabel>Categoría</FieldLabel>
        {categorias.length === 0 ? (
          <Text style={styles.warn}>Crea categorías en Más → Categorías.</Text>
        ) : (
          <View style={styles.pickerWrap}>
            <Picker
              selectedValue={categoria}
              onValueChange={(v) => setCategoria(v)}
              dropdownIconColor={colors.text}
              style={{ color: colors.text }}
            >
              <Picker.Item label="Selecciona…" value="" color={colors.textMuted} />
              {categorias.map((c) => (
                <Picker.Item key={c.nombre} label={`${c.icono} ${c.nombre}`} value={c.nombre} />
              ))}
            </Picker>
          </View>
        )}

        <FieldLabel>Cuenta</FieldLabel>
        {cuentasDisponibles.length === 0 ? (
          <Text style={styles.warn}>No hay saldo suficiente en una cuenta para este monto.</Text>
        ) : (
          <View style={styles.pickerWrap}>
            <Picker selectedValue={origen} onValueChange={setOrigen} style={{ color: colors.text }}>
              <Picker.Item label="Selecciona…" value="" />
              {cuentasDisponibles.map((c) => (
                <Picker.Item
                  key={c.id}
                  label={`${c.nombre} (${formatearNumero(c.saldo)} ${moneda})`}
                  value={c.id}
                />
              ))}
            </Picker>
          </View>
        )}

        {origen === 'tarjetaCredito' && (
          <>
            <FieldLabel>Cuotas</FieldLabel>
            <View style={styles.pickerWrap}>
              <Picker selectedValue={cuotas} onValueChange={(v) => setCuotas(v)} style={{ color: colors.text }}>
                {CUOTAS_OPTS.map((n) => (
                  <Picker.Item key={n} label={n === 1 ? '1 (contado)' : `${n} cuotas`} value={n} />
                ))}
              </Picker>
            </View>
            {cuotas > 1 && (
              <Text style={typography.small}>
                Cuota mensual aprox.: {formatearNumero(cantNum / cuotas)} {moneda}
              </Text>
            )}
          </>
        )}

        <FieldLabel>Nota (opcional)</FieldLabel>
        <TextInput
          style={styles.input}
          value={nota}
          onChangeText={setNota}
          placeholderTextColor={colors.textFaint}
        />

        <PrimaryButton title="Guardar gasto" onPress={onSubmit} style={{ marginTop: spacing.lg }} />
      </UICard>
    </ScreenWrap>
  );
}

function FieldLabel({ children }) {
  return <Text style={styles.fieldLab}>{children}</Text>;
}

const styles = StyleSheet.create({
  fieldLab: {
    ...typography.label,
    marginTop: spacing.md,
    marginBottom: spacing.xs,
    color: colors.textMuted,
    letterSpacing: 0.8,
  },
  pagoRow: {
    marginTop: spacing.md,
    paddingBottom: spacing.md,
    borderBottomWidth: 1,
    borderBottomColor: colors.stroke,
  },
  pagoConcepto: { flexShrink: 1, minWidth: 0 },
  link: { color: colors.accentBright, marginTop: 6, fontWeight: '600', fontSize: 13 },
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
  warn: { color: colors.danger, marginVertical: spacing.sm, fontSize: 14 },
});
