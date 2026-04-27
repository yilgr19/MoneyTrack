import React, { useMemo, useState } from 'react';
import { View, Text, StyleSheet, TextInput, TouchableOpacity, Alert, Platform, ScrollView } from 'react-native';
import { LinearGradient } from 'expo-linear-gradient';
import { Ionicons } from '@expo/vector-icons';
import DateTimePicker from '@react-native-community/datetimepicker';
import { Picker } from '@react-native-picker/picker';
import ScreenWrap from '../components/ScreenWrap';
import UICard from '../components/UICard';
import { PrimaryButton } from '../components/Buttons';
import { useApp } from '../context/AppContext';
import {
  CUENTAS,
  formatearNumero,
  calcularSaldosPorCuenta,
  generarIdMeta,
  normalizarMeta,
  ICONO_META_POR_DEFECTO,
} from '../lib/finance';
import { colors, spacing, radii, typography } from '../theme';

/**
 * Iconos vectoriales (Ionicons, estilo contorno) — criterio visual distinto a la rejilla de emojis de categorías.
 * Secuencia en scroll horizontal con anillo menta al elegir.
 */
const ICONOS_META_ION = [
  'trophy-outline',
  'flag-outline',
  'airplane-outline',
  'car-outline',
  'home-outline',
  'school-outline',
  'heart-outline',
  'gift-outline',
  'wallet-outline',
  'pie-chart-outline',
  'trending-up-outline',
  'rocket-outline',
  'boat-outline',
  'fitness-outline',
  'restaurant-outline',
  'bicycle-outline',
  'bus-outline',
  'cash-outline',
  'diamond-outline',
  'globe-outline',
  'sunny-outline',
  'medical-outline',
  'paw-outline',
  'book-outline',
  'leaf-outline',
  'umbrella-outline',
];

export default function MetasScreen() {
  const { state, replaceState } = useApp();
  const moneda = state.moneda || '';

  const [nombre, setNombre] = useState('');
  const [iconoMeta, setIconoMeta] = useState(ICONO_META_POR_DEFECTO);
  const [objetivo, setObjetivo] = useState('');
  const [plazo, setPlazo] = useState('');
  const [aporte, setAporte] = useState('');
  const [origen, setOrigen] = useState('');
  const [fecha, setFecha] = useState(new Date());
  const [showPicker, setShowPicker] = useState(false);

  const [aportarId, setAportarId] = useState(null);
  const [aportarCant, setAportarCant] = useState('');
  const [aportarOrigen, setAportarOrigen] = useState('');
  const [aportarFecha, setAportarFecha] = useState(new Date());
  const [showPicker2, setShowPicker2] = useState(false);

  const saldos = useMemo(() => calcularSaldosPorCuenta(state), [state]);
  const aporteNum = parseFloat(aporte) || 0;
  const cantAportar = parseFloat(aportarCant) || 0;

  const cuentasNuevaMeta = useMemo(() => {
    return CUENTAS.filter((c) => {
      const s = saldos[c.id] || 0;
      return s > 0 && (aporteNum === 0 || s >= aporteNum);
    });
  }, [saldos, aporteNum]);

  function crearMeta() {
    const obj = parseFloat(objetivo) || 0;
    if (!nombre.trim() || obj <= 0 || aporteNum <= 0 || !origen) {
      Alert.alert('Datos', 'Completa nombre, objetivo, aporte y origen.');
      return;
    }
    if (aporteNum > saldos.total) {
      Alert.alert('Saldo', 'No tienes saldo suficiente en total.');
      return;
    }
    if (aporteNum > (saldos[origen] || 0)) {
      Alert.alert('Saldo', 'No tienes suficiente en la cuenta seleccionada.');
      return;
    }
    const metaId = generarIdMeta();
    const pl = plazo.trim() ? parseInt(plazo, 10) : null;
    const fechaStr = `${fecha.getFullYear()}-${String(fecha.getMonth() + 1).padStart(2, '0')}-${String(fecha.getDate()).padStart(2, '0')}`;
    replaceState((s) => ({
      ...s,
      metas: [
        ...(s.metas || []),
        { id: metaId, nombre: nombre.trim(), objetivo: obj, plazo: pl || null, icono: iconoMeta },
      ],
      contribucionesMetas: [
        ...(s.contribucionesMetas || []),
        { metaId, cantidad: aporteNum, fecha: fechaStr, origen },
      ],
    }));
    Alert.alert('Listo', 'Meta creada con primer aporte.');
    setNombre('');
    setIconoMeta(ICONO_META_POR_DEFECTO);
    setObjetivo('');
    setPlazo('');
    setAporte('');
    setOrigen('');
  }

  function ejecutarAportar() {
    if (!aportarId || cantAportar <= 0 || !aportarOrigen) return;
    if (cantAportar > saldos.total) {
      Alert.alert('Saldo', 'Saldo total insuficiente.');
      return;
    }
    if (cantAportar > (saldos[aportarOrigen] || 0)) {
      Alert.alert('Saldo', 'Saldo en cuenta insuficiente.');
      return;
    }
    const fechaStr = `${aportarFecha.getFullYear()}-${String(aportarFecha.getMonth() + 1).padStart(2, '0')}-${String(aportarFecha.getDate()).padStart(2, '0')}`;
    replaceState((s) => ({
      ...s,
      contribucionesMetas: [
        ...(s.contribucionesMetas || []),
        { metaId: aportarId, cantidad: cantAportar, fecha: fechaStr, origen: aportarOrigen },
      ],
    }));
    Alert.alert('Listo', 'Aporte registrado.');
    setAportarId(null);
    setAportarCant('');
    setAportarOrigen('');
  }

  const metas = state.metas || [];

  return (
    <ScreenWrap includeTopInset={false} contentStyle={{ paddingTop: spacing.xs }}>
      <Text style={typography.label}>Ahorro</Text>
      <Text style={typography.hero}>Metas</Text>
      <Text style={[typography.subtitle, { marginBottom: spacing.lg }]}>Objetivos y aportes</Text>

      <UICard accent>
        <Text style={typography.label}>Disponible</Text>
        <Text
          style={styles.heroSaldo}
          adjustsFontSizeToFit
          minimumFontScale={0.6}
          numberOfLines={2}
          maxFontSizeMultiplier={1.2}
        >
          {formatearNumero(saldos.total)} <Text style={styles.heroMon}>{moneda}</Text>
        </Text>
        {saldos.totalReservado > 0 && (
          <Text style={typography.small}>
            Reservado en metas: {formatearNumero(saldos.totalReservado)} {moneda}
          </Text>
        )}
      </UICard>

      <UICard>
        <Text style={typography.label}>Nueva meta</Text>
        <Lab>Nombre</Lab>
        <TextInput style={styles.input} value={nombre} onChangeText={setNombre} placeholderTextColor={colors.textFaint} />
        <Lab>Icono</Lab>
        <Text style={styles.hintIcono}>
          Elige un símbolo de estilo contorno (distinto a las categorías). Desliza para ver más.
        </Text>
        <ScrollView
          horizontal
          showsHorizontalScrollIndicator={false}
          contentContainerStyle={styles.iconoMetaScroll}
          style={{ maxHeight: 72, marginTop: spacing.xs }}
        >
          {ICONOS_META_ION.map((ionName, i) => {
            const sel = ionName === iconoMeta;
            return (
              <TouchableOpacity
                key={ionName}
                onPress={() => setIconoMeta(ionName)}
                activeOpacity={0.88}
                style={[styles.iconoMetaChip, sel && styles.iconoMetaChipSel]}
                accessibilityLabel={`Elegir icono de meta, opción ${i + 1} de ${ICONOS_META_ION.length}`}
                accessibilityState={{ selected: sel }}
              >
                <Ionicons
                  name={ionName}
                  size={24}
                  color={sel ? colors.mint : colors.textMuted}
                />
              </TouchableOpacity>
            );
          })}
        </ScrollView>
        <View style={styles.metaVistaPrevia} accessibilityRole="text">
          <View style={styles.metaVistaPreviaRing}>
            <Ionicons name={iconoMeta} size={26} color={colors.mint} />
          </View>
          <View style={{ flex: 1, minWidth: 0 }}>
            <Text style={typography.label}>Vista previa</Text>
            <Text style={styles.hintIcono} numberOfLines={2}>
              Así se verá en Inicio y en esta pantalla.
            </Text>
          </View>
        </View>
        <Lab>Objetivo</Lab>
        <TextInput style={styles.input} value={objetivo} onChangeText={setObjetivo} keyboardType="decimal-pad" placeholderTextColor={colors.textFaint} />
        <Lab>Plazo (meses, opcional)</Lab>
        <TextInput style={styles.input} value={plazo} onChangeText={setPlazo} keyboardType="number-pad" placeholderTextColor={colors.textFaint} />
        <Lab>Primer aporte</Lab>
        <TextInput style={styles.input} value={aporte} onChangeText={setAporte} keyboardType="decimal-pad" placeholderTextColor={colors.textFaint} />
        <Lab>Origen</Lab>
        <View style={styles.pickerWrap}>
          <Picker selectedValue={origen} onValueChange={setOrigen} style={{ color: colors.text }}>
            <Picker.Item label="Selecciona…" value="" />
            {cuentasNuevaMeta.map((c) => (
              <Picker.Item
                key={c.id}
                label={`${c.nombre} (${formatearNumero(saldos[c.id])} ${moneda})`}
                value={c.id}
              />
            ))}
          </Picker>
        </View>
        <Lab>Fecha</Lab>
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
        <PrimaryButton title="Crear meta y aportar" onPress={crearMeta} style={{ marginTop: spacing.md }} />
      </UICard>

      <UICard>
        <Text style={typography.label}>Tus metas</Text>
        {metas.length === 0 ? (
          <Text style={typography.small}>Sin metas aún.</Text>
        ) : (
          metas.map((meta) => {
            const m = normalizarMeta(meta);
            const acum = (state.contribucionesMetas || [])
              .filter((c) => c.metaId === meta.id)
              .reduce((s, c) => s + c.cantidad, 0);
            const obj = parseFloat(meta.objetivo) || 0;
            const pct = obj > 0 ? Math.min(100, (acum / obj) * 100) : 0;
            return (
              <View key={meta.id} style={styles.metaRow}>
                <View style={styles.metaFilaTit}>
                  <View style={styles.metaIconoLista}>
                    <Ionicons name={m.icono} size={22} color={colors.mint} />
                  </View>
                  <Text style={styles.metaTit}>{m.nombre}</Text>
                </View>
                <Text style={typography.small}>
                  {formatearNumero(acum)} / {formatearNumero(obj)} {moneda} · {formatearNumero(pct, 0)}%
                </Text>
                <View style={styles.barBg}>
                  <LinearGradient
                    colors={[colors.mint, colors.success]}
                    start={{ x: 0, y: 0 }}
                    end={{ x: 1, y: 0 }}
                    style={[styles.barFill, { width: `${pct}%` }]}
                  />
                </View>
                <TouchableOpacity onPress={() => setAportarId(meta.id)} hitSlop={8}>
                  <Text style={styles.link}>Aportar →</Text>
                </TouchableOpacity>
              </View>
            );
          })
        )}
      </UICard>

      {aportarId && (
        <UICard style={{ marginBottom: 0 }}>
          <Text style={typography.label}>Nuevo aporte</Text>
          <TextInput
            style={styles.input}
            value={aportarCant}
            onChangeText={setAportarCant}
            keyboardType="decimal-pad"
            placeholder="Cantidad"
            placeholderTextColor={colors.textFaint}
          />
          <Lab>Cuenta</Lab>
          <View style={styles.pickerWrap}>
            <Picker selectedValue={aportarOrigen} onValueChange={setAportarOrigen} style={{ color: colors.text }}>
              <Picker.Item label="Selecciona…" value="" />
              {CUENTAS.filter((c) => {
                const s = saldos[c.id] || 0;
                return s > 0 && (cantAportar === 0 || s >= cantAportar);
              }).map((c) => (
                <Picker.Item key={c.id} label={c.nombre} value={c.id} />
              ))}
            </Picker>
          </View>
          <Lab>Fecha</Lab>
          <TouchableOpacity style={styles.input} onPress={() => setShowPicker2(true)}>
            <Text style={{ color: colors.text, fontSize: 16 }}>{aportarFecha.toLocaleDateString('es')}</Text>
          </TouchableOpacity>
          {showPicker2 && (
            <DateTimePicker
              value={aportarFecha}
              mode="date"
              display={Platform.OS === 'ios' ? 'spinner' : 'default'}
              onChange={(ev, d) => {
                if (Platform.OS !== 'ios') setShowPicker2(false);
                if (ev.type === 'dismissed') setShowPicker2(false);
                if (d) setAportarFecha(d);
              }}
            />
          )}
          <PrimaryButton title="Registrar aporte" onPress={ejecutarAportar} style={{ marginTop: spacing.md }} />
          <TouchableOpacity onPress={() => setAportarId(null)}>
            <Text style={[styles.link, { marginTop: spacing.md, textAlign: 'center' }]}>Cancelar</Text>
          </TouchableOpacity>
        </UICard>
      )}
    </ScreenWrap>
  );
}

function Lab({ children }) {
  return <Text style={styles.lab}>{children}</Text>;
}

const styles = StyleSheet.create({
  heroSaldo: {
    fontSize: 26,
    fontWeight: '700',
    color: colors.text,
    letterSpacing: -0.6,
    marginTop: 4,
    maxWidth: '100%',
  },
  heroMon: { fontSize: 16, color: colors.mint, fontWeight: '600' },
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
  metaRow: {
    marginBottom: spacing.lg,
    paddingBottom: spacing.md,
    borderBottomWidth: 1,
    borderBottomColor: colors.stroke,
  },
  metaTit: { color: colors.text, fontWeight: '700', fontSize: 16, flex: 1, minWidth: 0, flexShrink: 1 },
  barBg: {
    height: 7,
    backgroundColor: colors.barTrack,
    borderRadius: radii.pill,
    marginVertical: spacing.sm,
    overflow: 'hidden',
  },
  barFill: { height: '100%', borderRadius: radii.pill },
  link: { color: colors.accentBright, fontWeight: '600', fontSize: 14 },
  hintIcono: {
    ...typography.small,
    color: colors.textFaint,
    lineHeight: 18,
    marginBottom: 2,
  },
  iconoMetaScroll: {
    flexDirection: 'row',
    alignItems: 'center',
    paddingVertical: 2,
    paddingRight: spacing.sm,
  },
  iconoMetaChip: {
    width: 52,
    height: 52,
    borderRadius: 26,
    marginRight: spacing.sm,
    alignItems: 'center',
    justifyContent: 'center',
    borderWidth: 1,
    borderColor: colors.stroke,
    backgroundColor: 'rgba(125, 193, 145, 0.06)',
  },
  iconoMetaChipSel: {
    borderColor: colors.mint,
    borderWidth: 2,
    backgroundColor: 'rgba(125, 193, 145, 0.2)',
  },
  metaVistaPrevia: {
    flexDirection: 'row',
    alignItems: 'center',
    marginTop: spacing.md,
    marginBottom: spacing.xs,
    padding: spacing.md,
    borderRadius: radii.md,
    backgroundColor: 'rgba(125, 193, 145, 0.08)',
    borderWidth: 1,
    borderColor: 'rgba(125, 193, 145, 0.25)',
  },
  metaVistaPreviaRing: {
    width: 50,
    height: 50,
    borderRadius: 25,
    backgroundColor: 'rgba(0,0,0,0.2)',
    borderWidth: 1,
    borderColor: colors.mint,
    alignItems: 'center',
    justifyContent: 'center',
    marginRight: spacing.md,
  },
  metaFilaTit: { flexDirection: 'row', alignItems: 'center', marginBottom: 4 },
  metaIconoLista: {
    width: 40,
    height: 40,
    borderRadius: 20,
    backgroundColor: 'rgba(125, 193, 145, 0.12)',
    borderWidth: 1,
    borderColor: 'rgba(125, 193, 145, 0.35)',
    alignItems: 'center',
    justifyContent: 'center',
    marginRight: spacing.sm,
  },
});
