import React, { useEffect, useMemo, useRef, useState } from 'react';
import { View, Text, StyleSheet, TextInput, TouchableOpacity, Alert, Platform, Switch } from 'react-native';
import { useSafeAreaInsets } from 'react-native-safe-area-context';
import { LinearGradient } from 'expo-linear-gradient';
import { Ionicons } from '@expo/vector-icons';
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
  fechasCortesGastoConFallback,
  pagoDebeMostrarseParaPagar,
  obtenerCuentasOrigenGastoElegible,
  obtenerCuentasDestinoIngreso,
  obtenerSaldoDisponibleParaOrigenMovimiento,
  totalSaldoLiquido,
  agregarOFusionarPagoProgramadoCuotaCorte,
  reemplazarPagosRecordatorioTarjetas,
  filtrarPagosProgramadosCumplidosPorGasto,
  pagoProgramadoCumplidoPorGasto,
  abonoCoindiceCorteMensual,
  claveRecordatorioPagoCumplido,
} from '../lib/finance';
import { colors, spacing, radii, typography, shadows } from '../theme';

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

/** Colores vivos que rotan por fila; si el concepto sugiere TC/corte, se prioriza el morado. */
const PAGO_VISUAL_ROTATE = [
  {
    icon: 'calendar-outline',
    accent: '#fb923c',
    rowBg: 'rgba(251, 146, 60, 0.12)',
    border: '#fb923c',
    iconBg: 'rgba(251, 146, 60, 0.2)',
  },
  {
    icon: 'wallet-outline',
    accent: '#4ade80',
    rowBg: 'rgba(74, 222, 128, 0.1)',
    border: '#4ade80',
    iconBg: 'rgba(74, 222, 128, 0.2)',
  },
  {
    icon: 'flash-outline',
    accent: colors.chartBlue,
    rowBg: 'rgba(167, 216, 222, 0.1)',
    border: '#7dd3fc',
    iconBg: 'rgba(167, 216, 222, 0.2)',
  },
  {
    icon: 'receipt-outline',
    accent: '#a78bfa',
    rowBg: 'rgba(167, 139, 250, 0.1)',
    border: '#8b5cf6',
    iconBg: 'rgba(167, 139, 250, 0.18)',
  },
];

const PAGO_VISUAL_TARJETA = {
  icon: 'card-outline',
  accent: '#c084fc',
  rowBg: 'rgba(192, 132, 252, 0.12)',
  border: '#a855f7',
  iconBg: 'rgba(192, 132, 252, 0.2)',
};

function estiloPagoProgramadoFila(p, index) {
  const t = String(p.concepto || '');
  if (/corte|tarjeta|\btc\b|abono|cr[eé]dito|pago corte|l[ií]mite pago/i.test(t)) {
    return PAGO_VISUAL_TARJETA;
  }
  return PAGO_VISUAL_ROTATE[index % PAGO_VISUAL_ROTATE.length];
}

function pad(n) {
  return String(n).padStart(2, '0');
}

export default function GastosScreen() {
  const insets = useSafeAreaInsets();
  const { state, replaceState } = useApp();
  const moneda = state?.moneda || '';
  const flashPagoRef = useRef(null);
  const [flashPagoExito, setFlashPagoExito] = useState(false);
  /** Tras abonar el monto al corte (lista o monto que coincide con el corte sugerido). */
  const [cortePagoMensaje, setCortePagoMensaje] = useState(null);

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
  /** Abono/pago programado: a qué fila de tarjeta aplica (independiente por entidad). */
  const [tarjetaAbonoElegida, setTarjetaAbonoElegida] = useState('');

  const categorias = useMemo(
    () => (state?.categorias || []).map(normalizarCategoria),
    [state?.categorias]
  );

  const filasTarjeta = useMemo(() => state?.tarjetasCredito || [], [state?.tarjetasCredito]);

  const cantNum = parseFloat(cantidad) || 0;
  const origenCuenta = normalizarOrigenCuenta(origen) || String(origen || '').trim();
  const cuotasNum = Math.max(1, parseInt(cuotas, 10) || 1);
  const cuotaMensualTc = cantNum > 0 && cuotasNum > 0 ? cantNum / cuotasNum : cantNum;

  useEffect(() => {
    if (origenCuenta !== 'tarjetaCredito' || abonoDeudaTarjeta) return;
    if (filasTarjeta.length === 1) {
      setTarjetaCreditoElegida(String(filasTarjeta[0].id));
      return;
    }
    if (filasTarjeta.length > 1 && !filasTarjeta.some((t) => t && t.id === tarjetaCreditoElegida)) {
      setTarjetaCreditoElegida(String(filasTarjeta[0].id));
    }
  }, [origen, origenCuenta, abonoDeudaTarjeta, filasTarjeta, tarjetaCreditoElegida]);

  useEffect(() => {
    if (!abonoDeudaTarjeta) return;
    if (filasTarjeta.length === 1) {
      setTarjetaAbonoElegida(String(filasTarjeta[0].id));
      return;
    }
    if (filasTarjeta.length > 1 && !filasTarjeta.some((x) => x && x.id === tarjetaAbonoElegida)) {
      setTarjetaAbonoElegida(String(filasTarjeta[0].id));
    }
  }, [abonoDeudaTarjeta, filasTarjeta, tarjetaAbonoElegida]);

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
  const pagosPendientes = (state?.pagosProgramados || []).filter(
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
      if (p.tarjetaId) setTarjetaAbonoElegida(String(p.tarjetaId));
    } else {
      setOrigen(normalizarOrigenCuenta(p.cuenta) || p.cuenta || '');
    }
  }

  function onSubmit() {
    if (!nombre.trim() || cantNum <= 0 || !categoria || !origen) {
      Alert.alert('Datos incompletos', 'Nombre, cantidad, categoría y cuenta son obligatorios.');
      return;
    }

    const saldosAct = calcularSaldosPorCuenta(state || {});
    const esTcCarga = origenCuenta === 'tarjetaCredito';
    const cuotasVal = esTcCarga ? cuotasNum : 1;
    const cuotaMensualVal = esTcCarga ? cantNum / cuotasVal : cantNum;
    const saldoOrigen = obtenerSaldoDisponibleParaOrigenMovimiento(state || {}, origen);
    const montoAValidar = esTcCarga ? cuotaMensualVal : cantNum;
    const saldoTotal = saldosAct.total || 0;

    if (origenCuenta !== 'tarjetaCredito' && cantNum > saldoTotal) {
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
      origenCuenta === 'tarjetaCredito' &&
      (state?.tarjetasCredito || []).length > 1 &&
      !String(tarjetaCreditoElegida || '').trim()
    ) {
      Alert.alert('Tarjeta', 'Selecciona con qué entidad o tarjeta hiciste la compra.');
      return;
    }
    if (
      abonoDeudaTarjeta &&
      (state?.tarjetasCredito || []).length > 1 &&
      !String(tarjetaAbonoElegida || '').trim()
    ) {
      Alert.alert('Tarjeta', 'Indica a qué tarjeta aplica el abono o el pago del recordatorio.');
      return;
    }

    const catObj = categorias.find((c) => c.nombre === categoria);
    if (catObj && catObj.limite) {
      const lim = parseFloat(catObj.limite);
      const ah = new Date();
      const m0 = ah.getMonth();
      const y0 = ah.getFullYear();
      const gastosCategoria = (state?.gastos || []).filter(
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
    const esTcCarga = origenCuenta === 'tarjetaCredito';
    const origenGuardado = esTcCarga ? 'tarjetaCredito' : origen;
    const fechaStr = `${fecha.getFullYear()}-${pad(fecha.getMonth() + 1)}-${pad(fecha.getDate())}T${pad(fecha.getHours())}:${pad(fecha.getMinutes())}:00`;
    const tcsGuard = state?.tarjetasCredito || [];
    const tarjetaIdGasto =
      esTcCarga && tcsGuard.length === 1
        ? String(tcsGuard[0].id)
        : esTcCarga && tcsGuard.length > 1
          ? String(tarjetaCreditoElegida)
          : undefined;
    const tarjetaIdAbono =
      abonoDeudaTarjeta && tcsGuard.length === 1
        ? String(tcsGuard[0].id)
        : abonoDeudaTarjeta && tcsGuard.length > 1
          ? String(tarjetaAbonoElegida || '').trim()
          : undefined;
    const tarjetaFilaGastoOAbono = esTcCarga ? tarjetaIdGasto : abonoDeudaTarjeta ? tarjetaIdAbono : undefined;

    const nuevo = {
      nombre: nombre.trim(),
      cantidad: cantNum,
      fecha: fechaStr,
      categoria,
      origen: origenGuardado,
      nota: nota.trim() || null,
      cuotas: esTcCarga ? cuotasVal : 1,
      cuotaMensual: esTcCarga ? cuotaMensualVal : cantNum,
      ...(abonoDeudaTarjeta ? { esAbonoDeudaTarjeta: true } : {}),
      ...(tarjetaFilaGastoOAbono ? { tarjetaCreditoId: String(tarjetaFilaGastoOAbono) } : {}),
    };

    const pagosPrev = state?.pagosProgramados || [];
    const habiaCierreProgramado =
      !!pagoProgramadoEnUso || pagosPrev.some((p) => pagoProgramadoCumplidoPorGasto(nuevo, p));
    const refPagoCorte = new Date();
    let corteMsg = null;
    if (pagoProgramadoEnUso) {
      const p0 = pagosPrev.find((x) => x && String(x.id) === String(pagoProgramadoEnUso));
      if (p0 && p0.esRecordatorioTarjeta && p0.tipoRecordatorioTarjeta === 'corte' && p0.tarjetaId) {
        const t = (state?.tarjetasCredito || []).find((x) => x && String(x.id) === String(p0.tarjetaId));
        const nom = (t && String(t.nombreEntidad || '').trim()) || 'Tu tarjeta';
        corteMsg = {
          detalle: `${nom} — ${formatearNumero(cantNum, 0)} ${moneda}`.trim(),
        };
      }
    } else if (abonoDeudaTarjeta) {
      const hit = abonoCoindiceCorteMensual(nuevo, state || {}, refPagoCorte);
      if (hit) {
        corteMsg = {
          detalle: `${hit.nombreEntidad} — ${formatearNumero(cantNum, 0)} ${moneda}`.trim(),
        };
      }
    }

    replaceState((s) => {
      let gastos = [...(s.gastos || []), nuevo];
      let pagos = [...(s.pagosProgramados || [])];
      const recordatoriosCum = [...(s.recordatoriosPagoRegistrado || [])];

      if (esTcCarga) {
        const tcs = s.tarjetasCredito || [];
        const tTarjeta =
          (tarjetaIdGasto && tcs.find((x) => x && String(x.id) === String(tarjetaIdGasto))) ||
          tcs.find((x) => (parseFloat(x.tasaEA) || 0) > 0) ||
          tcs[0];
        const tasaEaVal = tTarjeta ? parseFloat(tTarjeta.tasaEA) || 0 : 0;
        const fechasC = fechasCortesGastoConFallback(fechaStr, cuotasVal, s, tarjetaIdGasto);
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
            cuenta: 'tarjetaCredito',
            notaUsuario: nota.trim(),
            tasaEA: tasaEaVal,
            tarjetaCreditoId: tarjetaIdGasto,
          });
        });
      }

      if (pagoProgramadoEnUso) {
        const pRem = pagos.find((p) => p && String(p.id) === String(pagoProgramadoEnUso));
        pagos = pagos.filter((p) => p && String(p.id) !== String(pagoProgramadoEnUso));
        if (pRem && pRem.esRecordatorioTarjeta) {
          const k = claveRecordatorioPagoCumplido(pRem);
          if (k && !recordatoriosCum.includes(k)) recordatoriosCum.push(k);
        }
      } else {
        pagos = filtrarPagosProgramadosCumplidosPorGasto(nuevo, pagos);
      }

      const st = {
        ...s,
        gastos,
        pagosProgramados: pagos,
        recordatoriosPagoRegistrado: recordatoriosCum,
      };
      return {
        ...st,
        pagosProgramados: reemplazarPagosRecordatorioTarjetas(st.pagosProgramados, st, new Date()),
      };
    });

    if (corteMsg) {
      if (flashPagoRef.current) clearTimeout(flashPagoRef.current);
      flashPagoRef.current = null;
      setFlashPagoExito(false);
      setCortePagoMensaje(corteMsg);
    } else if (habiaCierreProgramado) {
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
        <View style={styles.pagosCardOuter}>
          <LinearGradient
            colors={['rgba(251, 146, 60, 0.22)', 'rgba(125, 211, 252, 0.12)', 'rgba(34, 211, 238, 0.08)']}
            start={{ x: 0, y: 0 }}
            end={{ x: 1, y: 1 }}
            style={styles.pagosCardGradient}
          >
            <View style={styles.pagosHeaderRow}>
              <View style={styles.pagosHeaderIconWrap}>
                <Ionicons name="calendar" size={22} color={colors.accentGold} />
              </View>
              <View style={{ flex: 1, minWidth: 0 }}>
                <Text style={styles.pagosHeaderTit}>Pagos programados</Text>
                <Text style={styles.pagosHeaderSub}>Toca un ítem y el formulario se rellenará</Text>
              </View>
            </View>
            {pagosPendientes.map((p, idx) => {
              const v = estiloPagoProgramadoFila(p, idx);
              return (
                <TouchableOpacity
                  key={p.id}
                  style={[
                    styles.pagoRow,
                    { marginTop: idx > 0 ? spacing.sm : 0, borderLeftColor: v.border, backgroundColor: v.rowBg },
                  ]}
                  onPress={() => aplicarPagoProgramado(p)}
                  activeOpacity={0.86}
                >
                  <View style={[styles.pagoRowIcon, { backgroundColor: v.iconBg }]}>
                    <Ionicons name={v.icon} size={20} color={v.accent} />
                  </View>
                  <View style={styles.pagoRowText}>
                    <Text style={[typography.body, styles.pagoConcepto]} numberOfLines={2}>
                      {p.concepto} — {formatearNumero(p.monto)} {moneda}
                    </Text>
                    <View style={styles.pagoCtaRow}>
                      <Text style={[styles.pagoCta, { color: v.accent }]}>Usar en formulario</Text>
                      <Ionicons
                        name="chevron-forward"
                        size={16}
                        color={v.accent}
                        style={styles.pagoCtaChevron}
                      />
                    </View>
                  </View>
                </TouchableOpacity>
              );
            })}
          </LinearGradient>
        </View>
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

        {abonoDeudaTarjeta && filasTarjeta.length > 1 ? (
          <>
            <FieldLabel>¿A qué tarjeta aplica este pago?</FieldLabel>
            <View style={styles.pickerWrap}>
              <Picker
                selectedValue={tarjetaAbonoElegida}
                onValueChange={setTarjetaAbonoElegida}
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

        {!abonoDeudaTarjeta && origenCuenta === 'tarjetaCredito' && (
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
              <Picker
                selectedValue={cuotasNum}
                onValueChange={(v) =>
                  setCuotas(typeof v === 'number' ? v : Math.max(1, parseInt(String(v), 10) || 1))
                }
                style={{ color: colors.text }}
              >
                {CUOTAS_OPTS.map((n) => (
                  <Picker.Item key={n} label={n === 1 ? '1 (contado)' : `${n} cuotas`} value={n} />
                ))}
              </Picker>
            </View>
            {cuotasNum > 1 && (
              <Text style={typography.small}>
                Cuota mensual aprox.: {formatearNumero(cantNum / cuotasNum)} {moneda}. Cada cuota (incl. la 1) se
                contabiliza en el mes de la fecha de corte definida en Saldo → Tarjeta.
              </Text>
            )}
            {origenCuenta === 'tarjetaCredito' && cuotasNum === 1 && (
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
    {cortePagoMensaje ? (
      <View
        style={[styles.flashPagoWrap, { paddingBottom: Math.max(insets.bottom, spacing.md) + spacing.xs }]}
        pointerEvents="box-none"
        accessibilityViewIsModal
      >
        <View style={styles.cortePagoBox} accessibilityLiveRegion="polite">
          <Text style={styles.cortePagoBadge}>Pagado</Text>
          <Text style={styles.flashPagoTitulo}>Corte del mes</Text>
          <Text style={styles.flashPagoSub}>
            {cortePagoMensaje.detalle}
            {'\n\n'}
            Queda registrado: cubriste el monto al corte. Puedes seguir con tranquilidad o revisar Saldo.
          </Text>
          <PrimaryButton
            title="Listo, gracias"
            onPress={() => setCortePagoMensaje(null)}
            style={{ marginTop: spacing.lg }}
          />
        </View>
      </View>
    ) : flashPagoExito ? (
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
  cortePagoBox: {
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
  cortePagoBadge: {
    alignSelf: 'center',
    overflow: 'hidden',
    backgroundColor: 'rgba(52, 211, 153, 0.2)',
    color: colors.mint,
    fontWeight: '800',
    fontSize: 13,
    paddingHorizontal: spacing.md,
    paddingVertical: 6,
    borderRadius: radii.pill || 999,
    marginBottom: spacing.sm,
    textAlign: 'center',
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
  pagosCardOuter: {
    marginBottom: spacing.md,
    borderRadius: radii.lg,
    overflow: 'hidden',
    ...shadows.card,
  },
  pagosCardGradient: {
    borderRadius: radii.lg,
    padding: spacing.lg,
    borderWidth: 1,
    borderColor: 'rgba(251, 146, 60, 0.35)',
  },
  pagosHeaderRow: {
    flexDirection: 'row',
    alignItems: 'center',
    marginBottom: spacing.md,
  },
  pagosHeaderIconWrap: {
    width: 44,
    height: 44,
    borderRadius: 22,
    backgroundColor: 'rgba(217, 180, 74, 0.22)',
    borderWidth: 1,
    borderColor: 'rgba(217, 180, 74, 0.45)',
    alignItems: 'center',
    justifyContent: 'center',
    marginRight: spacing.md,
  },
  pagosHeaderTit: {
    color: colors.text,
    fontSize: 12,
    fontWeight: '800',
    letterSpacing: 1.1,
    textTransform: 'uppercase',
  },
  pagosHeaderSub: { ...typography.small, color: colors.textSecondary, marginTop: 3, lineHeight: 18 },
  pagoRow: {
    flexDirection: 'row',
    alignItems: 'center',
    borderLeftWidth: 4,
    borderRadius: radii.md,
    padding: spacing.md,
    borderWidth: 1,
    borderColor: 'rgba(199, 195, 227, 0.18)',
  },
  pagoRowIcon: {
    width: 40,
    height: 40,
    borderRadius: 20,
    alignItems: 'center',
    justifyContent: 'center',
    marginRight: spacing.md,
  },
  pagoRowText: { flex: 1, minWidth: 0 },
  pagoConcepto: { flexShrink: 1, minWidth: 0 },
  pagoCtaRow: { flexDirection: 'row', alignItems: 'center', marginTop: 8 },
  pagoCtaChevron: { marginLeft: 4 },
  pagoCta: { fontWeight: '700', fontSize: 13 },
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
