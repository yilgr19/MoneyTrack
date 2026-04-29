import React, { useMemo, useState } from 'react';
import {
  View,
  Text,
  StyleSheet,
  TextInput,
  TouchableOpacity,
  Alert,
  ScrollView,
  Modal,
  Pressable,
  Platform,
} from 'react-native';
import { Picker } from '@react-native-picker/picker';
import ScreenWrap from '../components/ScreenWrap';
import { HeaderConCampana } from '../components/HeaderConCampana';
import UICard from '../components/UICard';
import { PrimaryButton } from '../components/Buttons';
import { useApp } from '../context/AppContext';
import {
  formatearNumero,
  generarIdBolsillo,
  obtenerCuentasDestinoIngreso,
  obtenerCuentasOrigenGastoElegible,
  obtenerSaldoDisponibleParaOrigenMovimiento,
  totalSaldoBolsillos,
} from '../lib/finance';
import { colors, spacing, radii, typography, COLORES_BOLSILLO } from '../theme';

function colorBolsillo(b, i) {
  if (b && b.color && String(b.color).trim()) return String(b.color).trim();
  return COLORES_BOLSILLO[i % COLORES_BOLSILLO.length];
}

function pad(n) {
  return String(n).padStart(2, '0');
}

function categoriaParaMovimiento(state) {
  const c = state?.categorias?.[0];
  if (!c) return 'Otros';
  return typeof c === 'string' ? c : c.nombre || 'Otros';
}

export default function MisBolsillosScreen() {
  const { state, replaceState } = useApp();
  const moneda = state?.moneda || '';
  const bolsillos = state?.bolsillos || [];

  const [nombreNuevo, setNombreNuevo] = useState('');
  const [bolsilloEnviar, setBolsilloEnviar] = useState('');
  const [montoEnviar, setMontoEnviar] = useState('');
  const [origenEnviar, setOrigenEnviar] = useState('');
  const [bolsilloSacar, setBolsilloSacar] = useState('');
  const [montoSacar, setMontoSacar] = useState('');
  const [destinoSacar, setDestinoSacar] = useState('');
  const [colorNuevo, setColorNuevo] = useState(COLORES_BOLSILLO[0]);
  /** `nuevo` | bolsilloId */
  const [colorModal, setColorModal] = useState(null);

  const totalBols = useMemo(() => totalSaldoBolsillos(state || {}), [state]);

  const mEnviar = parseFloat(montoEnviar) || 0;
  const mSacar = parseFloat(montoSacar) || 0;

  const cuentasOrigen = useMemo(
    () =>
      obtenerCuentasOrigenGastoElegible(state || {}, mEnviar, mEnviar, {
        excluirTarjetaComoOrigen: true,
      }),
    [state, mEnviar]
  );

  const cuentasDestino = useMemo(() => obtenerCuentasDestinoIngreso(state || {}), [state]);

  function crearBolsillo() {
    const n = nombreNuevo.trim();
    if (!n) {
      Alert.alert('Nombre', 'Escribe un nombre para el bolsillo.');
      return;
    }
    replaceState((s) => ({
      ...s,
      bolsillos: [
        ...(s.bolsillos || []),
        { id: generarIdBolsillo(), nombre: n, saldo: 0, color: colorNuevo || COLORES_BOLSILLO[0] },
      ],
    }));
    setNombreNuevo('');
    Alert.alert('Listo', 'Bolsillo creado. Puedes enviarle dinero desde una caja.');
  }

  function enviarABolsillo() {
    if (!bolsilloEnviar || mEnviar <= 0 || !origenEnviar) {
      Alert.alert('Datos', 'Elige bolsillo, monto y cuenta de origen.');
      return;
    }
    const saldoOr = obtenerSaldoDisponibleParaOrigenMovimiento(state || {}, origenEnviar);
    if (mEnviar > saldoOr) {
      Alert.alert('Saldo', 'No hay suficiente saldo en la cuenta elegida.');
      return;
    }
    const cat = categoriaParaMovimiento(state);
    const bol = bolsillos.find((b) => b && String(b.id) === String(bolsilloEnviar));
    const nombreBol = (bol && bol.nombre) || 'Bolsillo';
    const fechaStr = `${new Date().getFullYear()}-${pad(new Date().getMonth() + 1)}-${pad(new Date().getDate())}T${pad(new Date().getHours())}:${pad(new Date().getMinutes())}:00`;
    const nuevoGasto = {
      nombre: `A bolsillo · ${nombreBol}`,
      cantidad: mEnviar,
      fecha: fechaStr,
      categoria: cat,
      origen: origenEnviar,
      esTransferenciaBolsillo: true,
      bolsilloId: String(bolsilloEnviar),
      nota: 'Ahorro en bolsillo (no cuenta como gasto del presupuesto)',
    };
    replaceState((s) => {
      const list = s.bolsillos || [];
      const next = list.map((b) =>
        b && String(b.id) === String(bolsilloEnviar)
          ? { ...b, saldo: (parseFloat(b.saldo) || 0) + mEnviar }
          : b
      );
      return {
        ...s,
        gastos: [...(s.gastos || []), nuevoGasto],
        bolsillos: next,
      };
    });
    setMontoEnviar('');
    setOrigenEnviar('');
    Alert.alert('Listo', 'Dinero movido al bolsillo. Sigue fuera del patrimonio estimado en Inicio.');
  }

  function sacarDeBolsillo() {
    if (!bolsilloSacar || mSacar <= 0 || !destinoSacar) {
      Alert.alert('Datos', 'Elige bolsillo, monto y cuenta destino.');
      return;
    }
    const bol = bolsillos.find((b) => b && String(b.id) === String(bolsilloSacar));
    const saldoBol = parseFloat(bol?.saldo) || 0;
    if (mSacar > saldoBol) {
      Alert.alert('Saldo', 'No hay tanto en ese bolsillo.');
      return;
    }
    const nombreBol = (bol && bol.nombre) || 'Bolsillo';
    const fechaStr = `${new Date().getFullYear()}-${pad(new Date().getMonth() + 1)}-${pad(new Date().getDate())}`;
    const ing = {
      cantidad: mSacar,
      fecha: fechaStr,
      origen: destinoSacar,
      esRetiroBolsillo: true,
      nota: `Desde bolsillo: ${nombreBol}`,
    };
    replaceState((s) => {
      const list = s.bolsillos || [];
      const next = list.map((b) =>
        b && String(b.id) === String(bolsilloSacar)
          ? { ...b, saldo: Math.max(0, (parseFloat(b.saldo) || 0) - mSacar) }
          : b
      );
      return {
        ...s,
        ingresos: [...(s.ingresos || []), ing],
        bolsillos: next,
      };
    });
    setMontoSacar('');
    setDestinoSacar('');
    Alert.alert('Listo', 'Dinero devuelto a la caja. Ya suma en patrimonio estimado.');
  }

  function guardarColorElegido(hex) {
    const h = String(hex || '').trim();
    if (!h) {
      setColorModal(null);
      return;
    }
    if (colorModal === 'nuevo') {
      setColorNuevo(h);
      setColorModal(null);
      return;
    }
    if (colorModal) {
      replaceState((s) => ({
        ...s,
        bolsillos: (s.bolsillos || []).map((b) =>
          b && String(b.id) === String(colorModal) ? { ...b, color: h } : b
        ),
      }));
    }
    setColorModal(null);
  }

  function eliminarBolsillo(b) {
    const s = parseFloat(b.saldo) || 0;
    if (s > 0.01) {
      Alert.alert(
        'No se puede borrar',
        'Primero saca el saldo del bolsillo hacia una caja, o márcalo en cero.'
      );
      return;
    }
    Alert.alert('Eliminar bolsillo', `¿Quitar «${b.nombre}»?`, [
      { text: 'Cancelar', style: 'cancel' },
      {
        text: 'Eliminar',
        style: 'destructive',
        onPress: () =>
          replaceState((st) => ({
            ...st,
            bolsillos: (st.bolsillos || []).filter((x) => x && x.id !== b.id),
          })),
      },
    ]);
  }

  return (
    <ScreenWrap includeTopInset={false} scrollEnabled={false} contentStyle={{ paddingTop: spacing.xs }}>
      <HeaderConCampana label="Ahorro" title="Mis bolsillos" subtitle="Separa dinero sin sumarlo al patrimonio" />

      <ScrollView
        style={styles.scrollMain}
        keyboardShouldPersistTaps="handled"
        keyboardDismissMode="on-drag"
        automaticallyAdjustKeyboardInsets={Platform.OS !== 'web'}
        showsVerticalScrollIndicator={false}
      >
        <UICard style={{ marginBottom: spacing.md }}>
          <Text style={[typography.small, { color: colors.textMuted, lineHeight: 20, marginBottom: spacing.sm }]}>
            El dinero que envíes a un bolsillo sale de tus cajas (efectivo, bancos, apps) y deja de contar en el
            «Patrimonio estimado» de Inicio: es ahorro aparte. Al sacarlo de vuelta, vuelve a sumar como saldo
            disponible.
          </Text>
          <Text style={typography.label}>
            Total en bolsillos: {formatearNumero(totalBols)} {moneda}
          </Text>
        </UICard>

        <UICard style={{ marginBottom: spacing.md }}>
          <Text style={typography.label}>Nuevo bolsillo</Text>
          <Text style={styles.lab}>Nombre</Text>
          <TextInput
            style={styles.input}
            value={nombreNuevo}
            onChangeText={setNombreNuevo}
            placeholder="Ej. Viaje, arreglo del carro…"
            placeholderTextColor={colors.textFaint}
          />
          <Text style={styles.lab}>Color (gama viva, distinta a la de categorías)</Text>
          <TouchableOpacity
            style={styles.colorPreviewRow}
            onPress={() => setColorModal('nuevo')}
            activeOpacity={0.8}
            accessibilityLabel="Elegir color para el nuevo bolsillo"
          >
            <View style={[styles.colorSwatchLg, { backgroundColor: colorNuevo, borderColor: colors.stroke }]} />
            <Text style={[typography.small, { color: colors.textSecondary, flex: 1, marginLeft: spacing.md }]}>
              Toca para elegir un tono más vivo
            </Text>
            <Text style={styles.chevronRow}>›</Text>
          </TouchableOpacity>
          <PrimaryButton title="Crear bolsillo" onPress={crearBolsillo} style={{ marginTop: spacing.md }} />
        </UICard>

        {bolsillos.length > 0 ? (
          <UICard style={{ marginBottom: spacing.md }}>
            <Text style={typography.label}>Tus bolsillos</Text>
            {bolsillos.map((b, i) => (
              <View
                key={b.id}
                style={[
                  styles.bolRow,
                  { borderLeftColor: colorBolsillo(b, i), borderLeftWidth: 4, paddingLeft: spacing.sm },
                ]}
              >
                <TouchableOpacity
                  onPress={() => setColorModal(b.id)}
                  style={styles.bolColorBtn}
                  hitSlop={{ top: 4, bottom: 4, left: 4, right: 4 }}
                  accessibilityLabel={`Cambiar color de ${b.nombre}`}
                >
                  <View
                    style={[
                      styles.colorSwatchSm,
                      { backgroundColor: colorBolsillo(b, i), borderColor: colors.stroke },
                    ]}
                  />
                </TouchableOpacity>
                <View style={{ flex: 1, minWidth: 0, marginLeft: spacing.sm }}>
                  <Text style={[typography.body, { fontWeight: '700' }]}>{b.nombre}</Text>
                  <Text style={[typography.small, { color: colors.textSecondary, marginTop: 4 }]}>
                    {formatearNumero(parseFloat(b.saldo) || 0)} {moneda}
                  </Text>
                </View>
                <TouchableOpacity
                  onPress={() => eliminarBolsillo(b)}
                  hitSlop={{ top: 8, bottom: 8, left: 8, right: 8 }}
                >
                  <Text style={styles.linkEliminar}>Quitar</Text>
                </TouchableOpacity>
              </View>
            ))}
          </UICard>
        ) : (
          <UICard style={{ marginBottom: spacing.md }}>
            <Text style={[typography.body, { color: colors.textMuted }]}>Aún no tienes bolsillos. Crea uno arriba.</Text>
          </UICard>
        )}

        {bolsillos.length > 0 ? (
          <>
            <UICard style={{ marginBottom: spacing.md }}>
              <Text style={typography.label}>Enviar dinero a un bolsillo</Text>
              <Text style={styles.lab}>Bolsillo</Text>
              <View style={styles.pickerWrap}>
                <Picker selectedValue={bolsilloEnviar} onValueChange={setBolsilloEnviar} style={{ color: colors.text }}>
                  <Picker.Item label="Selecciona…" value="" />
                  {bolsillos.map((b) => (
                    <Picker.Item key={b.id} label={b.nombre} value={b.id} />
                  ))}
                </Picker>
              </View>
              <Text style={styles.lab}>Monto</Text>
              <TextInput
                style={styles.input}
                value={montoEnviar}
                onChangeText={setMontoEnviar}
                keyboardType="decimal-pad"
                placeholder="0.00"
                placeholderTextColor={colors.textFaint}
              />
              <Text style={styles.lab}>Desde la cuenta</Text>
              {cuentasOrigen.length === 0 ? (
                <Text style={styles.warn}>Ajusta el monto o revisa saldos en Saldo.</Text>
              ) : (
                <View style={styles.pickerWrap}>
                  <Picker selectedValue={origenEnviar} onValueChange={setOrigenEnviar} style={{ color: colors.text }}>
                    <Picker.Item label="Selecciona…" value="" />
                    {cuentasOrigen.map((c) => (
                      <Picker.Item key={c.value} label={c.label} value={c.value} />
                    ))}
                  </Picker>
                </View>
              )}
              <PrimaryButton title="Enviar al bolsillo" onPress={enviarABolsillo} style={{ marginTop: spacing.md }} />
            </UICard>

            <UICard style={{ marginBottom: spacing.lg }}>
              <Text style={typography.label}>Sacar dinero hacia una caja</Text>
              <Text style={styles.lab}>Bolsillo</Text>
              <View style={styles.pickerWrap}>
                <Picker selectedValue={bolsilloSacar} onValueChange={setBolsilloSacar} style={{ color: colors.text }}>
                  <Picker.Item label="Selecciona…" value="" />
                  {bolsillos.map((b) => (
                    <Picker.Item key={b.id} label={b.nombre} value={b.id} />
                  ))}
                </Picker>
              </View>
              <Text style={styles.lab}>Monto</Text>
              <TextInput
                style={styles.input}
                value={montoSacar}
                onChangeText={setMontoSacar}
                keyboardType="decimal-pad"
                placeholder="0.00"
                placeholderTextColor={colors.textFaint}
              />
              <Text style={styles.lab}>A la cuenta</Text>
              {cuentasDestino.length === 0 ? (
                <Text style={styles.warn}>Configura cajas en Saldo.</Text>
              ) : (
                <View style={styles.pickerWrap}>
                  <Picker selectedValue={destinoSacar} onValueChange={setDestinoSacar} style={{ color: colors.text }}>
                    <Picker.Item label="Selecciona…" value="" />
                    {cuentasDestino.map((c) => (
                      <Picker.Item key={c.value} label={c.label} value={c.value} />
                    ))}
                  </Picker>
                </View>
              )}
              <PrimaryButton title="Sacar a la caja" onPress={sacarDeBolsillo} style={{ marginTop: spacing.md }} />
            </UICard>
          </>
        ) : null}
      </ScrollView>

      <Modal
        visible={colorModal != null}
        transparent
        animationType="fade"
        onRequestClose={() => setColorModal(null)}
      >
        <View style={styles.modalBackdrop}>
          <Pressable
            style={StyleSheet.absoluteFill}
            onPress={() => setColorModal(null)}
            accessibilityRole="button"
            accessibilityLabel="Cerrar selector de color"
          />
          <View style={styles.modalCard}>
            <Text style={typography.label}>Color del bolsillo</Text>
            <Text style={[typography.small, { color: colors.textSecondary, marginBottom: spacing.md }]}>
              Colores vivos y claros, pensados para no confundirse con los tonos de las categorías.
            </Text>
            <View style={styles.colorGrid}>
              {COLORES_BOLSILLO.map((hex) => (
                <TouchableOpacity
                  key={hex}
                  onPress={() => guardarColorElegido(hex)}
                  style={[styles.colorDot, { backgroundColor: hex, borderColor: colors.stroke }]}
                  accessibilityLabel={`Elegir color ${hex}`}
                />
              ))}
            </View>
            <TouchableOpacity onPress={() => setColorModal(null)} style={styles.modalCancel}>
              <Text style={styles.linkEliminar}>Cerrar</Text>
            </TouchableOpacity>
          </View>
        </View>
      </Modal>
    </ScreenWrap>
  );
}

const styles = StyleSheet.create({
  scrollMain: { flex: 1, backgroundColor: 'transparent' },
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
  warn: { color: colors.danger, marginVertical: spacing.sm, fontSize: 14 },
  bolRow: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    paddingVertical: spacing.md,
    borderBottomWidth: 1,
    borderBottomColor: colors.stroke,
  },
  linkEliminar: { color: colors.danger, fontSize: 14, fontWeight: '600' },
  chevronRow: {
    fontSize: 22,
    lineHeight: 24,
    color: colors.textFaint,
    fontWeight: '300',
  },
  colorPreviewRow: {
    flexDirection: 'row',
    alignItems: 'center',
    marginTop: spacing.xs,
    paddingVertical: spacing.sm,
  },
  colorSwatchLg: {
    width: 40,
    height: 40,
    borderRadius: 20,
    borderWidth: 1,
  },
  colorSwatchSm: {
    width: 28,
    height: 28,
    borderRadius: 14,
    borderWidth: 1,
  },
  bolColorBtn: { justifyContent: 'center' },
  colorGrid: {
    flexDirection: 'row',
    flexWrap: 'wrap',
    gap: 12,
    marginBottom: spacing.md,
  },
  colorDot: {
    width: 36,
    height: 36,
    borderRadius: 18,
    borderWidth: 1,
  },
  modalBackdrop: {
    flex: 1,
    backgroundColor: 'rgba(0,0,0,0.5)',
    justifyContent: 'center',
    padding: spacing.lg,
  },
  /** Capa a pantalla completa detrás; la tarjeta queda encima y recibe toques. */
  modalCard: {
    maxWidth: 400,
    alignSelf: 'center',
    width: '100%',
    backgroundColor: colors.surfaceSolid,
    borderRadius: radii.lg,
    padding: spacing.lg,
    borderWidth: 1,
    borderColor: colors.stroke,
  },
  modalCancel: { alignSelf: 'flex-end', marginTop: spacing.xs },
});
