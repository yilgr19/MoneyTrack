import React, { useEffect, useMemo, useRef, useState } from 'react';
import { View, Text, StyleSheet, TextInput, TouchableOpacity, Alert, Platform, Switch } from 'react-native';
import { useSafeAreaInsets } from 'react-native-safe-area-context';
import DateTimePicker from '@react-native-community/datetimepicker';
import { Picker } from '@react-native-picker/picker';
import ScreenWrap from '../components/ScreenWrap';
import { HeaderConCampana } from '../components/HeaderConCampana';
import UICard from '../components/UICard';
import { PrimaryButton } from '../components/Buttons';
import { useApp } from '../context/AppContext';
import {
  formatearNumero,
  calcularSaldosPorCuenta,
  normalizarOrigenCuenta,
  normalizarCategoria,
  montoGastoAfectaSaldoEnMes,
  fechaALocalISO,
  fechasCortesParaGastoTarjeta,
  pagoDebeMostrarseParaPagar,
  obtenerCuentasOrigenGastoElegible,
  obtenerCuentasDestinoIngreso,
  obtenerSaldoDisponibleParaOrigenMovimiento,
  totalSaldoLiquido,
  agregarOFusionarPagoProgramadoCuotaCorte,
  reemplazarPagosRecordatorioTarjetas,
  filtrarPagosProgramadosCumplidosPorGasto,
  pagoProgramadoCumplidoPorGasto,
} from '../lib/finance';
import { colors, spacing, radii, typography } from '../theme';

/** 5 títulos y 5 textos: se eligen al azar al mostrar el aviso (mismo tono que la campana). */
const LIMITE_CAT_ALERT_TITULOS = [
  'Uff, límite de categoría\n😱 😉',
  'Uff, te pasas del tope de categoría\n😱 😉',
  'Uff, este registro pasa el límite de la categoría\n😱 😉',
  'Uff, con esto te pasas del tope en esta categoría\n😱 😉',
  'Uff, ojo: este gasto pasa el límite de la categoría\n😱 😉',
];
const LIMITE_CAT_ALERT_MSGS = [
  'Con este gasto te pasas del tope del mes. Puedes cambiar el límite en Más → Categorías o bajar un poco los gastos. ¿Seguir y registrarlo igual?',
  'Pasas el tope con este registro. Cambia el límite en Categorías o gasta menos en lo que te queda. ¿Lo guardas igual?',
  'Este registro pasa el límite. Cambia el tope en Más → Categorías o baja un poco lo que sigues gastando. ¿Continuar?',
  'Te pasas del tope: puedes subirlo en Categorías o cuidar lo que aún te falta de mes. ¿Registrar igual?',
  'El límite se quedó chico. Más → Categorías o menos gasto en el mes, tú eliges. ¿Sigo con el guardado?',
];

const CUOTAS_OPTS = [1, 2, 3, 6, 12, 24];
const FLASH_PAGO_OK_MS = 3200;

function pad(n) {
  return String(n).padStart(2, '0');
}

export default function GastosScreen() {
  const insets = useSafeAreaInsets();
  const { state, replaceState } = useApp();
  const moneda = state.moneda || '';
  const flashPagoRef = useRef(null);
  const [flashPagoExito, setFlashPagoExito] = useState(false);

  const [nombre, setNombre] = useState('');
  const [cantidad, setCantidad] = useState('');
  const [fecha, setFecha] = useState(new Date());
  const [showPicker, setShowPicker] = useState(false);
  const [categoria, setCategoria] = useState('');
  const [origen, setOrigen] = useState('');
  const [cuotas, setCuotas] = useState(1);
  const [nota, setNota] = useState('');
  const [pagoProgramadoEnUso, setPagoProgramadoEnUso] = useState(null);
  /** Abono/liquidación de tarjeta: solo cajas con dinero, no cargo nuevo a la TC. */
  const [abonoDeudaTarjeta, setAbonoDeudaTarjeta] = useState(false);
  /** Con varias filas de TC en Saldo, el cargo a cuál aplica (extracto e importe por entidad). */
  const [tarjetaCreditoElegida, setTarjetaCreditoElegida] = useState('');

  const categorias = useMemo(
    () => (state.categorias || []).map(normalizarCategoria),
    [state.categorias]
  );

  const filasTarjeta = useMemo(() => state.tarjetasCredito || [], [state.tarjetasCredito]);

  useEffect(() => {
    if (normalizarOrigenCuenta(origen) !== 'tarjetaCredito' || abonoDeudaTarjeta) return;
    if (filasTarjeta.length === 1) {
      setTarjetaCreditoElegida(String(filasTarjeta[0].id));
      return;
    }
    if (filasTarjeta.length > 1 && !filasTarjeta.some((t) => t && t.id === tarjetaCreditoElegida)) {
      setTarjetaCreditoElegida(String(filasTarjeta[0].id));
    }
  }, [origen, abonoDeudaTarjeta, filasTarjeta, tarjetaCreditoElegida]);

  const cantNum = parseFloat(cantidad) || 0;
  const cuotaMensualTc = cantNum > 0 && cuotas > 0 ? cantNum / cuotas : cantNum;

  const cuentasDisponibles = useMemo(
    () =>
      obtenerCuentasOrigenGastoElegible(state || {}, cantNum, cuotaMensualTc, {
        excluirTarjetaComoOrigen: abonoDeudaTarjeta,
      }),
    [state, cantNum, cuotaMensualTc, abonoDeudaTarjeta]
  );

  const avisoCuentaContexto = useMemo(() => {
    if (cuentasDisponibles.length > 0 || cantNum <= 0) return null;
    const liq = totalSaldoLiquido(state || {});
    const tTxt = formatearNumero(liq, 0);
    if (liq >= cantNum) {
      if (abonoDeudaTarjeta) {
        return `En efectivo, bancos y billeteras tenías unos ${tTxt} en total: alcanza el monto, pero en ninguna línea suelta hay el valor completo. Mueve plata o divide el pago.`;
      }
      return `En efectivo, bancos y billeteras (sin cupo de tarjeta) tenías unos ${tTxt} en total: alcanza el monto, pero en ninguna cuenta o línea suelta hay el valor completo. Mueve plata, usa tarjeta, o anota el gasto en dos partes.`;
    }
    if (abonoDeudaTarjeta) {
      return 'No alcanza el monto con el saldo en efectivo, bancos y billeteras. Revisa Saldo, agrega un ingreso o ajusta el monto o la cuenta (no aplica abonar con la propia tarjeta).';
    }
    return 'No alcanza el monto con el saldo que la app en efectivo, bancos y billeteras. Revisa Saldo, suma un ingreso o ajusta el monto o la cuenta (incluida la tarjeta en cuota).';
  }, [state, cantNum, cuentasDisponibles.length, abonoDeudaTarjeta]);

  /** Saldos por línea (incl. 0,00) para ver cajas aunque no alcance el gasto. */
  const lineasSaldosReferencia = useMemo(
    () => obtenerCuentasDestinoIngreso(state || {}),
    [state]
  );

  useEffect(() => {
    if (origen && !cuentasDisponibles.some((c) => c.value === origen)) {
      setOrigen('');
    }
  }, [cuentasDisponibles, origen]);

  useEffect(() => {
    if (abonoDeudaTarjeta && normalizarOrigenCuenta(origen) === 'tarjetaCredito') {
      setOrigen('');
    }
    if (abonoDeudaTarjeta) {
      setCuotas(1);
    }
  }, [abonoDeudaTarjeta, origen]);

  useEffect(
    () => () => {
      if (flashPagoRef.current) clearTimeout(flashPagoRef.current);
    },
    []
  );

  const ahora = new Date();
  const pagosPendientes = (state.pagosProgramados || []).filter(
    (p) => p.activo !== false && pagoDebeMostrarseParaPagar(p, ahora)
  );

  function aplicarPagoProgramado(p) {
    setNombre(p.concepto || '');
    setCantidad(String(p.monto ?? ''));
    setCategoria(p.categoria || (categorias[0]?.nombre ?? ''));
    setCuotas(1);
    setNota(p.nota || '');
    setPagoProgramadoEnUso(p.id);
    const esPagoDeuda =
      !!p.esRecordatorioTarjeta ||
      !!p.esCuotaDiferida ||
      (p.concepto && /l[ií]mite pago|corte tc|pago corte/i.test(String(p.concepto)));
    setAbonoDeudaTarjeta(esPagoDeuda);
    if (esPagoDeuda) {
      setOrigen('');
    } else {
      setOrigen(normalizarOrigenCuenta(p.cuenta) || p.cuenta || '');
    }
  }

  function onSubmit() {
    if (!nombre.trim() || cantNum <= 0 || !categoria || !origen) {
      Alert.alert('Datos incompletos', 'Nombre, cantidad, categoría y cuenta son obligatorios.');
      return;
    }

    const saldosAct = calcularSaldosPorCuenta(state);
    const cuotasVal = origen === 'tarjetaCredito' ? cuotas : 1;
    const cuotaMensualVal = origen === 'tarjetaCredito' ? cantNum / cuotasVal : cantNum;
    const saldoOrigen = obtenerSaldoDisponibleParaOrigenMovimiento(state, origen);
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
    if (abonoDeudaTarjeta && normalizarOrigenCuenta(origen) === 'tarjetaCredito') {
      Alert.alert('Cuenta', 'Para abonar o pagar la deuda, elige efectivo, bancos o apps; no un nuevo cargo a la tarjeta.');
      return;
    }
    if (
      origen === 'tarjetaCredito' &&
      (state.tarjetasCredito || []).length > 1 &&
      !String(tarjetaCreditoElegida || '').trim()
    ) {
      Alert.alert('Tarjeta', 'Selecciona con qué entidad o tarjeta hiciste la compra.');
      return;
    }

    const catObj = categorias.find((c) => c.nombre === categoria);
    if (catObj && catObj.limite) {
      const lim = parseFloat(catObj.limite);
      const ah = new Date();
      const m0 = ah.getMonth();
      const y0 = ah.getFullYear();
      const gastosCategoria = (state.gastos || []).filter(
        (g) =>
          g.categoria === categoria && montoGastoAfectaSaldoEnMes(g, state, m0, y0) > 0
      );
      const gastadoMes = gastosCategoria.reduce(
        (s, g) => s + montoGastoAfectaSaldoEnMes(g, state, m0, y0),
        0
      );
      if (gastadoMes + montoAValidar > lim) {
        Alert.alert(
          LIMITE_CAT_ALERT_TITULOS[Math.floor(Math.random() * LIMITE_CAT_ALERT_TITULOS.length)],
          LIMITE_CAT_ALERT_MSGS[Math.floor(Math.random() * LIMITE_CAT_ALERT_MSGS.length)],
          [
            { text: 'Cancelar', style: 'cancel' },
            { text: 'Sí', onPress: () => guardarGasto(cuotasVal, cuotaMensualVal) },
          ]
        );
        return;
      }
    }

    guardarGasto(cuotasVal, cuotaMensualVal);
  }

  function guardarGasto(cuotasVal, cuotaMensualVal) {
    const fechaStr = `${fecha.getFullYear()}-${pad(fecha.getMonth() + 1)}-${pad(fecha.getDate())}T${pad(fecha.getHours())}:${pad(fecha.getMinutes())}:00`;
    const tcsGuard = state.tarjetasCredito || [];
    const tarjetaIdGasto =
      origen === 'tarjetaCredito' && tcsGuard.length === 1
        ? String(tcsGuard[0].id)
        : origen === 'tarjetaCredito' && tcsGuard.length > 1
          ? String(tarjetaCreditoElegida)
          : undefined;

    const nuevo = {
      nombre: nombre.trim(),
      cantidad: cantNum,
      fecha: fechaStr,
      categoria,
      origen,
      nota: nota.trim() || null,
      cuotas: origen === 'tarjetaCredito' ? cuotasVal : 1,
      cuotaMensual: origen === 'tarjetaCredito' ? cuotaMensualVal : cantNum,
      ...(abonoDeudaTarjeta ? { esAbonoDeudaTarjeta: true } : {}),
      ...(origen === 'tarjetaCredito' && tarjetaIdGasto
        ? { tarjetaCreditoId: tarjetaIdGasto }
        : {}),
    };

    const pagosPrev = state.pagosProgramados || [];
    const habiaCierreProgramado =
      !!pagoProgramadoEnUso || pagosPrev.some((p) => pagoProgramadoCumplidoPorGasto(nuevo, p));

    replaceState((s) => {
      let gastos = [...(s.gastos || []), nuevo];
      let pagos = [...(s.pagosProgramados || [])];

      if (origen === 'tarjetaCredito') {
        const tcs = s.tarjetasCredito || [];
        const tTarjeta =
          (tarjetaIdGasto && tcs.find((x) => x && String(x.id) === String(tarjetaIdGasto))) ||
          tcs.find((x) => (parseFloat(x.tasaEA) || 0) > 0) ||
          tcs[0];
        const tasaEaVal = tTarjeta ? parseFloat(tTarjeta.tasaEA) || 0 : 0;
        const fechasC = fechasCortesParaGastoTarjeta(fechaStr, cuotasVal, s, tarjetaIdGasto);
        fechasC.forEach((nextDate, i) => {
          const fechaCuota = fechaALocalISO(nextDate);
          if (!fechaCuota) return;
          pagos = agregarOFusionarPagoProgramadoCuotaCorte(pagos, {
            fechaCorteDate: nextDate,
            monto: cuotaMensualVal,
            nombre: nombre.trim(),
            iCuota: i + 1,
            nCuotas: cuotasVal,
            categoria,
            cuenta: origen,
            notaUsuario: nota.trim(),
            tasaEA: tasaEaVal,
          });
        });
      }

      if (pagoProgramadoEnUso) {
        pagos = pagos.filter((p) => p.id !== pagoProgramadoEnUso);
      } else {
        pagos = filtrarPagosProgramadosCumplidosPorGasto(nuevo, pagos);
      }

      const st = { ...s, gastos, pagosProgramados: pagos };
      return {
        ...st,
        pagosProgramados: reemplazarPagosRecordatorioTarjetas(st.pagosProgramados, st, new Date()),
      };
    });

    if (habiaCierreProgramado) {
      if (flashPagoRef.current) clearTimeout(flashPagoRef.current);
      setFlashPagoExito(true);
      flashPagoRef.current = setTimeout(() => {
        setFlashPagoExito(false);
        flashPagoRef.current = null;
      }, FLASH_PAGO_OK_MS);
    } else {
      Alert.alert('Listo', 'Gasto registrado.');
    }
    setNombre('');
    setCantidad('');
    setNota('');
    setCuotas(1);
    setPagoProgramadoEnUso(null);
    setAbonoDeudaTarjeta(false);
    setFecha(new Date());
  }

  return (
    <View style={styles.pantalla}>
    <ScreenWrap contentStyle={{ paddingTop: spacing.xs }}>
      <HeaderConCampana
        label="Movimientos"
        title="Registrar gasto"
        subtitle="Registra cada salida de dinero"
      />

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

        <View style={styles.switchRow}>
          <View style={{ flex: 1, minWidth: 0, paddingRight: spacing.md }}>
            <Text style={styles.fieldLab}>Pago o abono a la tarjeta (sin cargo nuevo a la TC)</Text>
            <Text style={typography.small}>
              Actívalo para liquidar o abonar deuda: solo verás cajas con dinero (efectivo, bancos, Nequi, etc.), no
              la tarjeta.
            </Text>
          </View>
          <Switch
            value={abonoDeudaTarjeta}
            onValueChange={setAbonoDeudaTarjeta}
            trackColor={{ false: colors.stroke, true: colors.accentDeep }}
            thumbColor={abonoDeudaTarjeta ? colors.accentBright : colors.textFaint}
          />
        </View>

        <FieldLabel>Cuenta</FieldLabel>
        {cuentasDisponibles.length === 0 ? (
          <Text style={styles.warn}>
            {avisoCuentaContexto || 'Escribe un monto o ajusta: no hay saldo en una sola línea o cuenta que alcance, o aún faltan datos en Saldo.'}
          </Text>
        ) : (
          <View style={styles.pickerWrap}>
            <Picker selectedValue={origen} onValueChange={setOrigen} style={{ color: colors.text }}>
              <Picker.Item label="Selecciona…" value="" />
              {cuentasDisponibles.map((c) => (
                <Picker.Item key={c.value} label={c.label} value={c.value} />
              ))}
            </Picker>
          </View>
        )}

        {!abonoDeudaTarjeta && origen === 'tarjetaCredito' && (
          <>
            {filasTarjeta.length > 1 ? (
              <>
                <FieldLabel>¿Con qué tarjeta?</FieldLabel>
                <View style={styles.pickerWrap}>
                  <Picker
                    selectedValue={tarjetaCreditoElegida}
                    onValueChange={setTarjetaCreditoElegida}
                    dropdownIconColor={colors.text}
                    style={{ color: colors.text }}
                  >
                    <Picker.Item label="Selecciona…" value="" color={colors.textMuted} />
                    {filasTarjeta.map((t) => (
                      <Picker.Item
                        key={t.id}
                        label={String(t.nombreEntidad || 'Tarjeta').trim() || 'Tarjeta'}
                        value={t.id}
                      />
                    ))}
                  </Picker>
                </View>
              </>
            ) : null}
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
                Cuota mensual aprox.: {formatearNumero(cantNum / cuotas)} {moneda}. Cada cuota (incl. la 1) se
                contabiliza en el mes de la fecha de corte definida en Saldo → Tarjeta.
              </Text>
            )}
            {origen === 'tarjetaCredito' && cuotas === 1 && (
              <Text style={typography.small}>
                Un solo pago: se imputa al mes de tu próximo corte (según Saldo → Tarjeta).
              </Text>
            )}
          </>
        )}

        {lineasSaldosReferencia.length > 0 ? (
          <View style={styles.refSaldosBlock}>
            <Text style={typography.label}>Saldos actuales (referencia)</Text>
            <Text style={styles.refSaldosHint}>
              Aunque una caja esté en 0,00 o no alcance el monto, sigues viendo dónde está cada parte; añade ingresos
              en Más → Ingresos o ajusta en Saldo.
            </Text>
            {lineasSaldosReferencia.map((row) => (
              <Text key={row.value} style={styles.refSaldosLinea}>
                • {row.label}
              </Text>
            ))}
            <Text style={styles.refSaldosLinea}>
              Líquido (sin cupo de tarjeta en esta suma): {formatearNumero(totalSaldoLiquido(state || {}), 0)}{' '}
              {moneda}
            </Text>
          </View>
        ) : null}

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
    {flashPagoExito ? (
      <View
        style={[styles.flashPagoWrap, { paddingBottom: Math.max(insets.bottom, spacing.md) + spacing.xs }]}
        pointerEvents="none"
        accessibilityLiveRegion="polite"
      >
        <View style={styles.flashPagoBox}>
          <Text style={styles.flashPagoTitulo}>¡Excelente!</Text>
          <Text style={styles.flashPagoSub}>Has registrado tu pago con éxito.</Text>
        </View>
      </View>
    ) : null}
    </View>
  );
}

function FieldLabel({ children }) {
  return <Text style={styles.fieldLab}>{children}</Text>;
}

const styles = StyleSheet.create({
  pantalla: { flex: 1 },
  flashPagoWrap: {
    position: 'absolute',
    left: spacing.lg,
    right: spacing.lg,
    bottom: 0,
    zIndex: 100,
  },
  flashPagoBox: {
    maxWidth: '100%',
    backgroundColor: colors.bgElevated,
    borderWidth: 2,
    borderColor: colors.mint,
    borderRadius: radii.lg,
    paddingVertical: spacing.lg,
    paddingHorizontal: spacing.lg,
    ...Platform.select({
      ios: {
        shadowColor: '#000',
        shadowOffset: { width: 0, height: 6 },
        shadowOpacity: 0.28,
        shadowRadius: 12,
      },
      android: { elevation: 8 },
    }),
  },
  flashPagoTitulo: {
    color: colors.mint,
    fontSize: 20,
    fontWeight: '800',
    textAlign: 'center',
  },
  flashPagoSub: {
    ...typography.body,
    color: colors.textSecondary,
    textAlign: 'center',
    marginTop: spacing.sm,
    lineHeight: 24,
  },
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
  switchRow: {
    flexDirection: 'row',
    alignItems: 'center',
    marginTop: spacing.md,
    padding: spacing.md,
    borderRadius: radii.md,
    borderWidth: 1,
    borderColor: colors.stroke,
    backgroundColor: 'rgba(0,0,0,0.08)',
  },
  warn: { color: colors.danger, marginVertical: spacing.sm, fontSize: 14 },
  refSaldosBlock: {
    marginTop: spacing.md,
    padding: spacing.md,
    borderRadius: radii.md,
    backgroundColor: 'rgba(0,0,0,0.1)',
    borderWidth: 1,
    borderColor: colors.stroke,
  },
  refSaldosHint: { ...typography.small, color: colors.textFaint, marginTop: spacing.xs, lineHeight: 18 },
  refSaldosLinea: { ...typography.small, color: colors.textSecondary, marginTop: 4, lineHeight: 19 },
});
