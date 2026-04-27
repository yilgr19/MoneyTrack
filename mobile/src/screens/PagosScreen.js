import React, { useEffect, useMemo, useRef, useState } from 'react';
import { View, Text, StyleSheet, TextInput, TouchableOpacity, Alert, Platform, Switch } from 'react-native';
import { useSafeAreaInsets } from 'react-native-safe-area-context';
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
  claveRecordatorioPagoCumplido,
} from '../lib/finance';
import { colors, spacing, radii, typography } from '../theme';

const FLASH_DURACION_MS = 2800;

export default function PagosScreen() {
  const insets = useSafeAreaInsets();
  const flashTimer = useRef(null);
  const { state, replaceState } = useApp();
  const moneda = state.moneda || '';
  const categorias = useMemo(() => (state.categorias || []).map(normalizarCategoria), [state.categorias]);

  /** Alineada con Gastos: no mostrar recordatorios TC ya atendidos (evita duplicar tras reemplazar). */
  const pagosListaVisible = useMemo(() => {
    const excl = new Set(state?.recordatoriosPagoRegistrado || []);
    return (state?.pagosProgramados || []).filter((p) => {
      if (!p) return false;
      if (p.esRecordatorioTarjeta) {
        const k = claveRecordatorioPagoCumplido(p);
        if (k && excl.has(k)) return false;
      }
      return true;
    });
  }, [state?.pagosProgramados, state?.recordatoriosPagoRegistrado]);

  const [concepto, setConcepto] = useState('');
  const [monto, setMonto] = useState('');
  const [frecuencia, setFrecuencia] = useState('mensual');
  /** Solo quincenal: día 1 o 15. Mensual toma el día (1–28) de la fecha del primer pago. */
  const [diaPago, setDiaPago] = useState(1);
  const [cuenta, setCuenta] = useState('');
  const [categoria, setCategoria] = useState('');
  const [fechaInicio, setFechaInicio] = useState(new Date());
  const [showPicker, setShowPicker] = useState(false);
  const [activo, setActivo] = useState(true);
  const [nota, setNota] = useState('');
  const [flashMsg, setFlashMsg] = useState(null);

  const diasQuincenal = useMemo(() => [1, 15], []);

  useEffect(
    () => () => {
      if (flashTimer.current) clearTimeout(flashTimer.current);
    },
    []
  );

  function mostrarFlash(texto) {
    if (flashTimer.current) clearTimeout(flashTimer.current);
    setFlashMsg(texto);
    flashTimer.current = setTimeout(() => {
      setFlashMsg(null);
      flashTimer.current = null;
    }, FLASH_DURACION_MS);
  }

  function guardar() {
    const m = parseFloat(monto) || 0;
    if (!concepto.trim() || m <= 0 || !cuenta || !categoria) {
      Alert.alert('Datos', 'Completa concepto, monto, cuenta y categoría.');
      return;
    }
    const fi = new Date(
      fechaInicio.getFullYear(),
      fechaInicio.getMonth(),
      fechaInicio.getDate(),
      12,
      0,
      0
    );
    if (frecuencia === 'quincenal') {
      fi.setDate(diaPago);
    }
    const fechaStr = fechaALocalISO(fi);
    const diaMensual = Math.min(28, fi.getDate());
    const p = {
      id: generarIdPagoProgramado(),
      concepto: concepto.trim(),
      monto: m,
      frecuencia,
      fechaInicio: fechaStr,
      diaPago: frecuencia === 'semanal' ? fi.getDate() : frecuencia === 'mensual' ? diaMensual : diaPago,
      cuenta,
      categoria,
      activo,
      nota: nota.trim() || '',
    };
    replaceState((s) => ({ ...s, pagosProgramados: [...(s.pagosProgramados || []), p] }));
    if (frecuencia === 'mensual') {
      mostrarFlash('Mensual: mismo día cada mes. Listo.');
    } else if (frecuencia === 'quincenal') {
      mostrarFlash('Quincenal guardado.');
    } else {
      mostrarFlash('Semanal guardado.');
    }
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
          replaceState((s) => {
            const list = s.pagosProgramados || [];
            const p = list.find((x) => x && String(x.id) === String(id));
            let rpr = [...(s.recordatoriosPagoRegistrado || [])];
            if (p && p.esRecordatorioTarjeta) {
              const k = claveRecordatorioPagoCumplido(p);
              if (k && !rpr.includes(k)) rpr.push(k);
            }
            return {
              ...s,
              pagosProgramados: list.filter((x) => x && String(x.id) !== String(id)),
              recordatoriosPagoRegistrado: rpr,
            };
          }),
      },
    ]);
  }

  return (
    <View style={styles.pantalla}>
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
        {frecuencia === 'quincenal' && (
          <>
            <Lab>Pagas cada quincena el</Lab>
            <View style={styles.pickerWrap}>
              <Picker selectedValue={diaPago} onValueChange={setDiaPago} style={{ color: colors.text }}>
                {diasQuincenal.map((d) => (
                  <Picker.Item key={d} label={d === 1 ? 'Día 1' : 'Día 15'} value={d} />
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
        <Lab>
          {frecuencia === 'quincenal' ? 'Fecha del primer pago (mes; el día se alinea a 1 o 15 al guardar)' : 'Fecha del primer pago (día exacto y mes)'}
        </Lab>
        {frecuencia === 'mensual' ? (
          <Text style={styles.ayudaFecha}>
            <Text style={styles.ayudaNegrita}>Solo hace falta la fecha</Text>: el{' '}
            <Text style={styles.ayudaNegrita}>día del 1 al 28</Text> fija en qué número de cada mes se
            programará el pago, y el <Text style={styles.ayudaNegrita}>mes</Text> en qué empieza. No
            hace falta elegir de nuevo el &quot;día de pago&quot; — ya queda fijado por esta fecha.
          </Text>
        ) : frecuencia === 'quincenal' ? (
          <Text style={styles.ayudaFecha}>
            El <Text style={styles.ayudaNegrita}>día 1 o 15</Text> (selector arriba) es en qué quincena
            cae el pago. La <Text style={styles.ayudaNegrita}>fecha</Text> fija <Text style={styles.ayudaNegrita}>partir de qué mes</Text> se registra; al guardar ajustamos el día
            a 1 o 15.
          </Text>
        ) : (
          <Text style={styles.ayudaFecha}>
            Elige <Text style={styles.ayudaNegrita}>qué día de la semana</Text> con la primera fecha: el
            gasto se repetirá cada 7 días.
          </Text>
        )}
        <Text style={styles.ayudaCampana}>
          Campana: al cerrar el panel, el aviso se oculta; vuelve al toque o si cambia el mensaje.
        </Text>
        <TouchableOpacity
          style={styles.input}
          onPress={() => setShowPicker(true)}
          accessibilityLabel="Fecha del primer pago, día y mes"
          accessibilityRole="button"
        >
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
        {pagosListaVisible.length === 0 ? (
          <Text style={typography.small}>Sin pagos programados.</Text>
        ) : (
          pagosListaVisible.map((p) => (
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
    {flashMsg ? (
      <View
        style={[styles.flashWrap, { paddingBottom: Math.max(insets.bottom, spacing.md) + spacing.xs }]}
        pointerEvents="none"
        accessibilityLiveRegion="polite"
      >
        <View style={styles.flashBox}>
          <Text style={styles.flashText}>{flashMsg}</Text>
        </View>
      </View>
    ) : null}
    </View>
  );
}

function Lab({ children }) {
  return <Text style={styles.lab}>{children}</Text>;
}

const styles = StyleSheet.create({
  pantalla: { flex: 1 },
  flashWrap: {
    position: 'absolute',
    left: spacing.lg,
    right: spacing.lg,
    bottom: 0,
    zIndex: 100,
    alignItems: 'center',
  },
  flashBox: {
    maxWidth: '100%',
    backgroundColor: colors.bgElevated,
    borderWidth: 1,
    borderColor: colors.mint,
    borderRadius: radii.md,
    paddingVertical: spacing.md,
    paddingHorizontal: spacing.lg,
    shadowColor: '#000',
    shadowOffset: { width: 0, height: 4 },
    shadowOpacity: 0.25,
    shadowRadius: 8,
    elevation: 6,
  },
  flashText: {
    ...typography.small,
    color: colors.text,
    lineHeight: 20,
    textAlign: 'center',
    fontWeight: '600',
  },
  ayudaFecha: {
    ...typography.small,
    color: colors.textSecondary,
    marginBottom: spacing.sm,
    lineHeight: 20,
  },
  ayudaNegrita: { fontWeight: '700', color: colors.text },
  ayudaCampana: {
    ...typography.small,
    color: colors.textFaint,
    marginBottom: spacing.sm,
    lineHeight: 18,
  },
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
