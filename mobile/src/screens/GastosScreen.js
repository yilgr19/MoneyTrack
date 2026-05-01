import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
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
import ReceiptScannerModal from '../components/ReceiptScannerModal';
import { useApp } from '../context/AppContext';
import {
  formatearNumero,
  calcularSaldosPorCuenta,
  normalizarOrigenCuenta,
  normalizarCategoria,
  montoGastoCuentaParaPresupuestoEnMes,
  fechaALocalISO,
  pagoDebeMostrarseParaPagar,
  obtenerCuentasOrigenGastoElegible,
  obtenerSaldoDisponibleParaOrigenMovimiento,
  totalSaldoLiquido,
  reemplazarPagosRecordatorioTarjetas,
  filtrarPagosProgramadosCumplidosPorGasto,
  pagoProgramadoCumplidoPorGasto,
  abonoCoindiceCorteMensual,
  claveRecordatorioPagoCumplido,
  fechaGastoRecomendadaTrasOCR,
  existeAbonoDeudaTarjetaEnMes,
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
  'Pasas el tope de la categoría. ¿Guardar igual?',
  'Superas el límite del mes en esta categoría. ¿Continuar?',
  'Este monto pasa el tope. ¿Registrar?',
  'Vas por encima del límite de categoría. ¿Sigo?',
  'Tope de categoría superado. ¿Guardar?',
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
  /** Android: un solo picker `datetime` puede lanzar `dismiss` de undefined en la lib nativa. */
  const [androidFechaOpen, setAndroidFechaOpen] = useState(false);
  const [androidHoraOpen, setAndroidHoraOpen] = useState(false);
  const [categoria, setCategoria] = useState('');
  const [origen, setOrigen] = useState('');
  const [cuotas, setCuotas] = useState(1);
  const [nota, setNota] = useState('');
  /** Si el usuario indica que la nota enumera todos los productos/cantidades del recibo (metadata en el gasto). */
  const [notaListadoTicketCompleto, setNotaListadoTicketCompleto] = useState(false);
  /** manual | ocr | hibrido — cómo se rellenó el formulario (OCR no guarda solo; el usuario confirma). */
  const [tipoEntrada, setTipoEntrada] = useState('manual');
  const ocrSnapshotRef = useRef(null);
  const [scannerVisible, setScannerVisible] = useState(false);
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
        return `Hay ~${tTxt} en total, pero ninguna caja sola alcanza el monto. Divide el pago o mueve saldo.`;
      }
      return `Hay ~${tTxt} en total, pero ninguna cuenta sola alcanza el monto. Usa TC, divide o mueve saldo.`;
    }
    if (abonoDeudaTarjeta) {
      return 'No alcanza en efectivo/bancos/apps. Revisa Saldo o baja el monto.';
    }
    return 'No alcanza en cajas líquidas. Revisa Saldo, ingresos o usa tarjeta/cuotas.';
  }, [state, cantNum, cuentasDisponibles.length, abonoDeudaTarjeta]);

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

  const aplicarFechaGasto = useCallback((d) => {
    if (!d) return;
    if (
      tipoEntrada === 'ocr' &&
      ocrSnapshotRef.current != null &&
      d.getTime() !== ocrSnapshotRef.current.fechaMs
    ) {
      setTipoEntrada('hibrido');
    }
    setFecha(d);
  }, [tipoEntrada]);

  const ahora = new Date();
  const pagosPendientes = (state?.pagosProgramados || []).filter(
    (p) => p.activo !== false && pagoDebeMostrarseParaPagar(p, ahora)
  );

  function aplicarPagoProgramado(p) {
    setTipoEntrada('manual');
    ocrSnapshotRef.current = null;
    setNotaListadoTicketCompleto(false);
    setNombre(p.concepto || '');
    setCantidad(String(p.monto ?? ''));
    setCategoria(p.categoria || (categorias[0]?.nombre ?? ''));
    setCuotas(1);
    setNota(p.nota || '');
    setPagoProgramadoEnUso(p.id);
    const esPagoDeuda =
      !!p.esRecordatorioTarjeta ||
      (p.concepto && /l[ií]mite pago|corte tc|pago corte/i.test(String(p.concepto)));
    setAbonoDeudaTarjeta(esPagoDeuda);
    if (esPagoDeuda) {
      setOrigen('');
      if (p.tarjetaId) setTarjetaAbonoElegida(String(p.tarjetaId));
    } else {
      setOrigen(normalizarOrigenCuenta(p.cuenta) || p.cuenta || '');
    }
  }

  function onNombreChange(t) {
    if (tipoEntrada === 'ocr' && ocrSnapshotRef.current != null && t !== ocrSnapshotRef.current.nombre) {
      setTipoEntrada('hibrido');
    }
    setNombre(t);
  }

  function onCantidadChange(t) {
    if (tipoEntrada === 'ocr' && ocrSnapshotRef.current != null && t !== ocrSnapshotRef.current.cantidad) {
      setTipoEntrada('hibrido');
    }
    setCantidad(t);
  }

  function onNotaChange(t) {
    if (tipoEntrada === 'ocr' && ocrSnapshotRef.current != null && t !== ocrSnapshotRef.current.nota) {
      setTipoEntrada('hibrido');
    }
    if (!String(t).trim()) setNotaListadoTicketCompleto(false);
    setNota(t);
  }

  function aplicarDatosDesdeRecibo({ monto, establecimiento, fecha: fechaParsed, textoCompleto, productos }) {
    const est = String(establecimiento || '').trim();
    const fechaValida =
      fechaParsed instanceof Date && !Number.isNaN(fechaParsed.getTime()) ? fechaParsed : null;
    const listaProd = Array.isArray(productos) ? productos.filter((s) => String(s || '').trim()) : [];
    const hayAlgoParseado =
      (monto != null && monto > 0) || est.length > 0 || fechaValida != null || listaProd.length > 0;
    if (!hayAlgoParseado) {
      const ocrVacio = !String(textoCompleto || '').trim();
      Alert.alert(
        'No se pudo leer',
        ocrVacio
          ? 'No se extrajo texto: el OCR en JavaScript no es fiable en móvil. Instala una build de desarrollo o release que incluya el módulo nativo expo-text-extractor (ML Kit/Vision); con Expo Go el respaldo OCR suele fallar. También puedes registrar el gasto a mano.'
          : 'Sí hay texto pero no se detectaron total claro, fecha ni comercio en el formato. Completa los campos manualmente.'
      );
      return;
    }
    /**
     * Fecha para el gasto: si el OCR da otro mes, usamos hoy para que Inicio/análisis del mes en curso
     * incluyan el registro (la fecha sigue siendo editable antes de guardar).
     */
    const fechaNueva = fechaGastoRecomendadaTrasOCR(fechaValida, new Date());
    let cantStr = '';
    if (monto != null && monto > 0) {
      const v = Math.round(Number(monto) * 100) / 100;
      cantStr = Number.isInteger(v) ? String(v) : v.toFixed(2);
    }
    /** Nombre = lugar / comercio detectado en el ticket */
    const nombreVal = est || 'Recibo';
    /** Nota = detalle: ítems enumerados cuando el OCR los reconoce */
    let notaVal = '';
    if (listaProd.length > 0) {
      notaVal = listaProd.map((line, i) => `${i + 1}. ${line}`).join('\n');
    } else {
      notaVal = est ? `Recibo: ${est}` : 'Recibo escaneado';
    }
    setCantidad(cantStr);
    setNombre(nombreVal);
    setNota(notaVal);
    setNotaListadoTicketCompleto(false);
    setFecha(fechaNueva);
    ocrSnapshotRef.current = {
      cantidad: cantStr,
      nombre: nombreVal,
      nota: notaVal,
      fechaMs: fechaNueva.getTime(),
    };
    setTipoEntrada('ocr');
    if (monto == null || monto <= 0) {
      const partes = [];
      if (est.length) partes.push(`lugar: ${est}`);
      if (fechaValida) {
        partes.push(
          `fecha: ${fechaValida.toLocaleDateString('es-CO', { day: 'numeric', month: 'short', year: 'numeric' })}`
        );
      }
      if (listaProd.length) partes.push(`${listaProd.length} productos en la nota`);
      Alert.alert(
        'Lectura incompleta',
        `No se detectó el total del recibo.${partes.length ? `\n\nSí pudimos leer: ${partes.join(' · ')}.` : ''}\n\nIndica el monto a mano o vuelve a capturar el ticket.`
      );
    }
  }

  function abrirEscanerRecibo() {
    if (Platform.OS === 'web') {
      Alert.alert('Escáner de recibo', 'Disponible en la app iOS / Android: instala MoneyTrack para usar la cámara OCR.');
      return;
    }
    setScannerVisible(true);
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
          g.categoria === categoria &&
          montoGastoCuentaParaPresupuestoEnMes(g, state, m0, y0) > 0
      );
      const gastadoMes = gastosCategoria.reduce(
        (s, g) => s + montoGastoCuentaParaPresupuestoEnMes(g, state, m0, y0),
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
    if (abonoDeudaTarjeta) {
      const tidAb =
        tcsGuard.length === 1
          ? String(tcsGuard[0].id)
          : String(tarjetaAbonoElegida || '').trim();
      if (!tidAb && tcsGuard.length > 1) {
        Alert.alert('Tarjeta', 'Elige a qué tarjeta aplica este pago.');
        return;
      }
      if (existeAbonoDeudaTarjetaEnMes(state || {}, tidAb || null, fechaStr)) {
        Alert.alert(
          'Pago a la tarjeta',
          'Ya registraste un pago a esa tarjeta en este mes calendario. Para no mezclar cierres, usa un solo movimiento al mes o edita el anterior.',
        );
        return;
      }
    }
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
      ...(notaListadoTicketCompleto && nota.trim() ? { notaListadoTicketCompleto: true } : {}),
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
    setNotaListadoTicketCompleto(false);
    setTipoEntrada('manual');
    ocrSnapshotRef.current = null;
    setCuotas(1);
    setPagoProgramadoEnUso(null);
    setAbonoDeudaTarjeta(false);
    setFecha(new Date());
  }

  return (
    <View style={styles.pantalla}>
    <ScreenWrap contentStyle={{ paddingTop: spacing.xs }}>
      <HeaderConCampana label="Movimientos" title="Registrar gasto" subtitle="Salidas de dinero" />

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
                <Text style={styles.pagosHeaderSub}>Toca para rellenar el formulario</Text>
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
        <Text style={styles.entradaHint} accessibilityLiveRegion="polite">
          {tipoEntrada === 'manual' && 'Entrada: manual.'}
          {tipoEntrada === 'ocr' && 'Desde recibo: revisa categoría y cuenta.'}
          {tipoEntrada === 'hibrido' && 'Mixto: recibo + tus cambios.'}
        </Text>

        <FieldLabel>Nombre</FieldLabel>
        <TextInput
          style={styles.input}
          value={nombre}
          onChangeText={onNombreChange}
          placeholder="Ej: Supermercado"
          placeholderTextColor={colors.textFaint}
        />

        <FieldLabel>Cantidad</FieldLabel>
        <View style={styles.cantidadRow}>
          <TextInput
            style={[styles.input, styles.cantidadInput]}
            value={cantidad}
            onChangeText={onCantidadChange}
            keyboardType="decimal-pad"
            placeholder="0.00"
            placeholderTextColor={colors.textFaint}
          />
          <TouchableOpacity
            style={styles.camFab}
            onPress={abrirEscanerRecibo}
            accessibilityLabel="Escanear recibo con la cámara"
            accessibilityRole="button"
          >
            <Ionicons name="camera" size={22} color="#0c0812" />
          </TouchableOpacity>
        </View>

        <FieldLabel>Fecha y hora</FieldLabel>
        <TouchableOpacity
          style={styles.input}
          onPress={() => {
            if (Platform.OS === 'android') setAndroidFechaOpen(true);
            else setShowPicker(true);
          }}
        >
          <Text style={{ color: colors.text, fontSize: 16 }}>{fecha.toLocaleString('es')}</Text>
        </TouchableOpacity>
        {Platform.OS === 'ios' && showPicker ? (
          <DateTimePicker
            value={fecha}
            mode="datetime"
            display="spinner"
            onChange={(ev, d) => {
              if (ev?.type === 'dismissed') setShowPicker(false);
              if (d) aplicarFechaGasto(d);
            }}
          />
        ) : null}
        {Platform.OS === 'android' && androidFechaOpen ? (
          <DateTimePicker
            value={fecha}
            mode="date"
            display="default"
            onChange={(ev, d) => {
              setAndroidFechaOpen(false);
              if (ev?.type === 'dismissed' || !d) return;
              aplicarFechaGasto(
                new Date(
                  d.getFullYear(),
                  d.getMonth(),
                  d.getDate(),
                  fecha.getHours(),
                  fecha.getMinutes(),
                  fecha.getSeconds()
                )
              );
              setAndroidHoraOpen(true);
            }}
          />
        ) : null}
        {Platform.OS === 'android' && androidHoraOpen ? (
          <DateTimePicker
            value={fecha}
            mode="time"
            display="default"
            onChange={(ev, d) => {
              setAndroidHoraOpen(false);
              if (ev?.type === 'dismissed' || !d) return;
              aplicarFechaGasto(
                new Date(
                  fecha.getFullYear(),
                  fecha.getMonth(),
                  fecha.getDate(),
                  d.getHours(),
                  d.getMinutes(),
                  0
                )
              );
            }}
          />
        ) : null}

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
            <Text style={styles.fieldLab}>Abono a tarjeta (sin nuevo cargo a TC)</Text>
            <Text style={typography.small}>Solo cajas con saldo; no eliges tarjeta como origen.</Text>
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
            {avisoCuentaContexto || 'Monto o cuenta: revisa Saldo o el valor ingresado.'}
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
                ~{formatearNumero(cantNum / cuotasNum)} {moneda}/mes · imputación según corte (Saldo → Tarjeta).
              </Text>
            )}
            {origenCuenta === 'tarjetaCredito' && cuotasNum === 1 && (
              <Text style={typography.small}>Un pago: mes según tu corte en Saldo → Tarjeta.</Text>
            )}
          </>
        )}

        <FieldLabel>Nota (opcional)</FieldLabel>
        <Text style={[typography.small, { marginBottom: spacing.xs, color: colors.textSecondary, lineHeight: 18 }]}>
          Detalle libre; el total oficial es «Cantidad».
        </Text>
        <TextInput
          style={[styles.input, styles.inputNota]}
          value={nota}
          onChangeText={onNotaChange}
          placeholder="Ej.: 1. Producto — cantidad o total de línea…"
          placeholderTextColor={colors.textFaint}
          multiline
          textAlignVertical="top"
        />
        <View style={[styles.switchRow, { marginTop: spacing.sm }]}>
          <View style={{ flex: 1, minWidth: 0, paddingRight: spacing.md }}>
            <Text style={styles.fieldLab}>Nota = ítems completos del recibo</Text>
            <Text style={typography.small}>Actívalo si la nota lista todo lo comprado.</Text>
          </View>
          <Switch
            value={notaListadoTicketCompleto}
            onValueChange={setNotaListadoTicketCompleto}
            trackColor={{ false: colors.stroke, true: colors.accentDeep }}
            thumbColor={notaListadoTicketCompleto ? colors.accentBright : colors.textFaint}
          />
        </View>

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
            Pago al corte registrado.
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
    <ReceiptScannerModal
      visible={scannerVisible}
      onClose={() => setScannerVisible(false)}
      onDatosParsed={aplicarDatosDesdeRecibo}
    />
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
  entradaHint: {
    ...typography.small,
    color: colors.textSecondary,
    marginTop: spacing.xs,
    marginBottom: spacing.sm,
    lineHeight: 18,
  },
  cantidadRow: {
    flexDirection: 'row',
    alignItems: 'stretch',
    gap: spacing.sm,
  },
  cantidadInput: {
    flex: 1,
    minWidth: 0,
  },
  camFab: {
    width: 52,
    borderRadius: radii.md,
    backgroundColor: colors.mint,
    alignItems: 'center',
    justifyContent: 'center',
    borderWidth: 1,
    borderColor: 'rgba(12, 8, 18, 0.15)',
    ...Platform.select({
      ios: {
        shadowColor: '#34d399',
        shadowOpacity: 0.35,
        shadowRadius: 8,
        shadowOffset: { width: 0, height: 2 },
      },
      android: { elevation: 4 },
    }),
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
  inputNota: {
    minHeight: 120,
    paddingTop: spacing.md,
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
});
