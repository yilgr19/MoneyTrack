import React, { useState, useEffect, useMemo, useRef, useCallback } from 'react';
import {
  View,
  Text,
  StyleSheet,
  TextInput,
  Alert,
  Modal,
  TouchableOpacity,
  Pressable,
  KeyboardAvoidingView,
  Platform,
  ScrollView,
  Animated,
  PanResponder,
} from 'react-native';
import DateTimePicker from '@react-native-community/datetimepicker';
import { Picker } from '@react-native-picker/picker';
import { Ionicons } from '@expo/vector-icons';
import ScreenWrap from '../components/ScreenWrap';
import { HeaderConCampana } from '../components/HeaderConCampana';
import UICard from '../components/UICard';
import { PrimaryButton, GhostButton } from '../components/Buttons';
import { useApp } from '../context/AppContext';
import {
  CUENTAS,
  formatearNumero,
  calcularSaldosPorCuenta,
  BANCO_OTRO_VALUE,
  getBancosOptionsForMoneda,
  getBankLabelByValue,
  totalSaldoBancosDetalle,
  PLATAFORMA_OTRO_VALUE,
  getPlataformasOptions,
  cuentaSaldoPlataforma,
  getPlataformaLabelByValue,
  totalPlataformasTresSaldos,
  generarIdTarjetaCredito,
  resumenAlertasTarjetasCredito,
  fechaALocalISO,
  proximaOcurrenciaMensual,
  parseFechaHoraLocal,
  reemplazarPagosRecordatorioTarjetas,
} from '../lib/finance';
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

const CUENTA_ICONS = {
  efectivo: 'wallet-outline',
  banco: 'business-outline',
  tarjetaCredito: 'card-outline',
};

/** Nequi, Daviplata y billeteras se editan solo en la tarjeta «Mis plataformas». */
const CUENTA_IDS_CARD_PLATAFORMAS = new Set(['nequi', 'daviplata', 'billeteras']);

function newBancoLineId() {
  return `b-${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
}

function emptyBancoDraftLine(bancoOptions) {
  const first = bancoOptions[0] || { value: BANCO_OTRO_VALUE, label: 'Otro' };
  return {
    id: newBancoLineId(),
    bancoKey: first.value,
    otroNombre: '',
    monto: '',
  };
}

function persistedToDraft(row, bancoOptions) {
  const id = row.id || newBancoLineId();
  const nombre = row.nombre || '';
  const saldoVal = row.saldo != null ? row.saldo : 0;
  const found = bancoOptions.find(
    (b) => b.value !== BANCO_OTRO_VALUE && (b.label === nombre || b.value === nombre)
  );
  if (found) {
    return { id, bancoKey: found.value, otroNombre: '', monto: String(saldoVal) };
  }
  return { id, bancoKey: BANCO_OTRO_VALUE, otroNombre: nombre, monto: String(saldoVal) };
}

function draftToPersisted(row, bancoOptions) {
  const opt = bancoOptions.find((b) => b.value === row.bancoKey);
  const nombre =
    row.bancoKey === BANCO_OTRO_VALUE
      ? (row.otroNombre.trim() || 'Otro banco')
      : (opt?.label || row.bancoKey);
  return { id: row.id, nombre, saldo: parseFloat(row.monto) || 0 };
}

function newPlataformaLineId() {
  return `p-${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
}

function emptyPlataformaDraftLine(opts) {
  const first = opts[0] || { value: 'nequi', label: 'Nequi' };
  return {
    id: newPlataformaLineId(),
    platformKey: first.value,
    otroNombre: '',
    monto: '',
  };
}

function persistedPlataformaToDraft(row, opts) {
  const id = row.id || newPlataformaLineId();
  const saldoVal = row.saldo != null ? row.saldo : 0;
  if (row.platformValue && opts.some((o) => o.value === row.platformValue)) {
    return { id, platformKey: row.platformValue, otroNombre: '', monto: String(saldoVal) };
  }
  const nombre = row.nombre || '';
  const found = opts.find(
    (o) => o.value !== PLATAFORMA_OTRO_VALUE && (o.label === nombre || o.value === nombre)
  );
  if (found) return { id, platformKey: found.value, otroNombre: '', monto: String(saldoVal) };
  return { id, platformKey: PLATAFORMA_OTRO_VALUE, otroNombre: nombre, monto: String(saldoVal) };
}

function draftPlataformaToPersisted(row, opts) {
  const opt = opts.find((o) => o.value === row.platformKey);
  const nombre =
    row.platformKey === PLATAFORMA_OTRO_VALUE
      ? (row.otroNombre.trim() || 'Plataforma')
      : (opt?.label || row.platformKey);
  return {
    id: row.id,
    platformValue: row.platformKey,
    nombre,
    saldo: parseFloat(row.monto) || 0,
  };
}

function clampDiaMes(v, fallback) {
  const n = parseInt(v, 10);
  if (Number.isNaN(n)) return fallback;
  return Math.min(28, Math.max(1, n));
}

function soloFechaGuardada(str) {
  const s = String(str || '').trim();
  if (!s) return '';
  return s.slice(0, 10);
}

function defaultISOFromLegacyDia(diaRaw, kind) {
  const dia = clampDiaMes(diaRaw, kind === 'corte' ? 15 : 5);
  const pat = new Date(2020, 0, dia, 12, 0, 0);
  const prox = proximaOcurrenciaMensual(pat, new Date());
  return prox ? fechaALocalISO(prox) : '';
}

function tcToDraft(t) {
  const fc = soloFechaGuardada(t.fechaHoraCorte) || defaultISOFromLegacyDia(t.diaCorte, 'corte');
  const fl = soloFechaGuardada(t.fechaHoraLimitePago) || defaultISOFromLegacyDia(t.diaLimitePago, 'pago');
  return {
    id: t.id || generarIdTarjetaCredito(),
    nombreEntidad: t.nombreEntidad || '',
    tasaEA: t.tasaEA != null && String(t.tasaEA).trim() !== '' ? String(t.tasaEA) : '',
    cupoTotal: t.cupoTotal != null && String(t.cupoTotal).trim() !== '' ? String(t.cupoTotal) : '',
    cupoUtilizado:
      t.cupoUtilizado != null && String(t.cupoUtilizado).trim() !== '' ? String(t.cupoUtilizado) : '',
    fechaHoraCorte: fc,
    fechaHoraLimitePago: fl,
  };
}

function draftToTc(row) {
  const fc = soloFechaGuardada(row.fechaHoraCorte);
  const fl = soloFechaGuardada(row.fechaHoraLimitePago);
  const dc = parseFechaHoraLocal(fc);
  const dl = parseFechaHoraLocal(fl);
  return {
    id: row.id,
    nombreEntidad: row.nombreEntidad.trim(),
    tasaEA: parseFloat(String(row.tasaEA).replace(',', '.')) || 0,
    cupoTotal: parseFloat(String(row.cupoTotal).replace(',', '.')) || 0,
    cupoUtilizado: parseFloat(String(row.cupoUtilizado).replace(',', '.')) || 0,
    fechaHoraCorte: fc,
    fechaHoraLimitePago: fl,
    diaCorte: dc ? Math.min(28, dc.getDate()) : 15,
    diaLimitePago: dl ? Math.min(28, dl.getDate()) : 5,
  };
}

function emptyTcDraft() {
  return {
    id: generarIdTarjetaCredito(),
    nombreEntidad: '',
    tasaEA: '',
    cupoTotal: '',
    cupoUtilizado: '',
    fechaHoraCorte: defaultISOFromLegacyDia(15, 'corte'),
    fechaHoraLimitePago: defaultISOFromLegacyDia(5, 'pago'),
  };
}

function EditCard({ icon, title, subtitle, onPress, hint }) {
  return (
    <TouchableOpacity style={styles.editCard} onPress={onPress} activeOpacity={0.75}>
      <View style={styles.editCardIcon}>
        <Ionicons name={icon} size={22} color={colors.accentBright} />
      </View>
      <View style={styles.editCardBody}>
        <Text style={styles.editCardTitle}>{title}</Text>
        <Text style={styles.editCardSub} numberOfLines={2}>
          {subtitle}
        </Text>
        {hint ? (
          <Text style={styles.editCardHint} numberOfLines={2}>
            {hint}
          </Text>
        ) : null}
      </View>
      <Ionicons name="chevron-forward" size={20} color={colors.textFaint} />
    </TouchableOpacity>
  );
}

export default function SaldoScreen() {
  const { state, replaceState } = useApp();
  const [moneda, setMoneda] = useState(state.moneda || '');
  const [saldos, setSaldos] = useState(() => ({ ...emptySaldosCuentas(), ...state.saldosCuentas }));
  const [bancosDetalle, setBancosDetalle] = useState(() => state.bancosDetalle || []);
  const [plataformasDetalle, setPlataformasDetalle] = useState(() => state.plataformasDetalle || []);
  const [limiteTc, setLimiteTc] = useState(String(state.limiteTarjetaCredito || 0));
  const [presupuesto, setPresupuesto] = useState(String(state.presupuestoMensual || 0));
  const [nota, setNota] = useState(state.saldoInicialNota || '');

  /** null | 'moneda' | 'tarjetasCredito' | 'presupuesto' | 'nota' | { type: 'cuenta', id: string } */
  const [sheet, setSheet] = useState(null);
  const [draftMoneda, setDraftMoneda] = useState('');
  const [draftMonto, setDraftMonto] = useState('');
  const [draftPresupuesto, setDraftPresupuesto] = useState('');
  const [draftNota, setDraftNota] = useState('');
  const [bancoModalLines, setBancoModalLines] = useState([]);
  const [plataformaModalLines, setPlataformaModalLines] = useState([]);
  const [tarjetasCredito, setTarjetasCredito] = useState(() => state.tarjetasCredito || []);
  const [tcModalLines, setTcModalLines] = useState([]);
  /** null | { idx: number, field: 'corte' | 'limite' } */
  const [tcPicker, setTcPicker] = useState(null);

  const sheetDragY = useRef(new Animated.Value(0)).current;
  /** Offset Y del ScrollView del modal; si es ~0, el arrastre hacia abajo cierra desde cualquier zona */
  const sheetScrollY = useRef(0);

  const closeSheet = useCallback(() => {
    setTcPicker(null);
    setSheet(null);
  }, []);

  useEffect(() => {
    if (sheet !== null) {
      sheetDragY.setValue(0);
      sheetScrollY.current = 0;
    }
  }, [sheet, sheetDragY]);

  const sheetPanResponder = useMemo(
    () =>
      PanResponder.create({
        onStartShouldSetPanResponder: () => false,
        onMoveShouldSetPanResponderCapture: (_, g) => {
          if (g.dy < 14 || g.dy < Math.abs(g.dx) * 0.55) return false;
          return sheetScrollY.current <= 2;
        },
        onMoveShouldSetPanResponder: (_, g) => {
          if (g.dy < 14 || g.dy < Math.abs(g.dx) * 0.55) return false;
          return sheetScrollY.current <= 2;
        },
        onPanResponderMove: (_, g) => {
          if (g.dy > 0) sheetDragY.setValue(g.dy);
        },
        onPanResponderRelease: (_, g) => {
          const dismiss = g.dy > 80 || (g.vy > 0 && g.vy > 0.5);
          if (dismiss) {
            Animated.timing(sheetDragY, {
              toValue: 900,
              duration: 220,
              useNativeDriver: true,
            }).start(() => {
              sheetDragY.setValue(0);
              closeSheet();
            });
          } else {
            Animated.spring(sheetDragY, {
              toValue: 0,
              friction: 8,
              tension: 80,
              useNativeDriver: true,
            }).start();
          }
        },
      }),
    [closeSheet, sheetDragY]
  );

  useEffect(() => {
    setMoneda(state.moneda || '');
    setSaldos({ ...emptySaldosCuentas(), ...state.saldosCuentas });
    setBancosDetalle(state.bancosDetalle || []);
    setPlataformasDetalle(state.plataformasDetalle || []);
    const tc = state.tarjetasCredito || [];
    setTarjetasCredito(tc);
    if (tc.length > 0) {
      setLimiteTc(String(tc.reduce((s, t) => s + (parseFloat(t.cupoTotal) || 0), 0)));
    } else {
      setLimiteTc(String(state.limiteTarjetaCredito || 0));
    }
    setPresupuesto(String(state.presupuestoMensual || 0));
    setNota(state.saldoInicialNota || '');
  }, [
    state.moneda,
    state.saldosCuentas,
    state.bancosDetalle,
    state.plataformasDetalle,
    state.tarjetasCredito,
    state.limiteTarjetaCredito,
    state.presupuestoMensual,
    state.saldoInicialNota,
  ]);

  /**
   * Incluye borrador de esta pantalla (saldos, tarjetas, tope) para previsualizar el cupo libre
   * (cupo total − deuda) coherente con el formulario.
   */
  const dataCalculoBorrador = useMemo(
    () => ({
      ...state,
      saldosCuentas: { ...emptySaldosCuentas(), ...state.saldosCuentas, ...saldos },
      tarjetasCredito: tarjetasCredito || [],
      limiteTarjetaCredito: parseFloat(limiteTc) || 0,
    }),
    [state, saldos, tarjetasCredito, limiteTc]
  );
  const saldosActuales = useMemo(
    () => calcularSaldosPorCuenta(dataCalculoBorrador),
    [dataCalculoBorrador]
  );

  const monedaLabel = useMemo(() => MONEDAS.find((m) => m.value === moneda)?.label || 'Sin definir', [moneda]);

  const totalBancoMostrado = useMemo(() => {
    if (bancosDetalle.length > 0) return totalSaldoBancosDetalle(bancosDetalle);
    return parseFloat(saldos.banco) || 0;
  }, [bancosDetalle, saldos.banco]);

  const bancoOptions = useMemo(() => getBancosOptionsForMoneda(moneda), [moneda]);
  const plataformaOptions = useMemo(() => getPlataformasOptions(), []);

  const totalPlataformasMostrado = useMemo(
    () => totalPlataformasTresSaldos(saldos),
    [saldos.nequi, saldos.daviplata, saldos.billeteras]
  );

  const tcTarjetaHint = useMemo(() => {
    const tc = tarjetasCredito;
    if (tc && tc.length > 0) {
      const r = resumenAlertasTarjetasCredito({
        ...state,
        tarjetasCredito: tc,
        limiteTarjetaCredito: parseFloat(limiteTc) || 0,
      });
      const proxPago = Math.min(...r.tarjetas.map((t) => t.diasPago));
      const proxCorte = Math.min(...r.tarjetas.map((t) => t.diasCorte));
      return `Pago más cercano en ${proxPago} d · corte en ${proxCorte} d · ${tc.length} tarjeta(s)`;
    }
    const lim = parseFloat(limiteTc) || 0;
    if (lim > 0) return `Cupo total (legacy): ${formatearNumero(lim)} · toca para fechas y cupos por banco`;
    return 'Toca para cupo, tasa E.A., corte y pago';
  }, [state, tarjetasCredito, limiteTc]);

  useEffect(() => {
    if (sheet?.type !== 'cuenta' || sheet.id !== 'banco') return;
    setBancoModalLines((prev) =>
      prev.map((line) => {
        if (bancoOptions.some((b) => b.value === line.bancoKey)) return line;
        const hint = getBankLabelByValue(line.bancoKey);
        return {
          ...line,
          bancoKey: BANCO_OTRO_VALUE,
          otroNombre: line.otroNombre.trim() || hint || '',
        };
      })
    );
  }, [moneda, sheet?.type, sheet?.id, bancoOptions]);

  useEffect(() => {
    if (sheet !== 'plataformas') return;
    setPlataformaModalLines((prev) =>
      prev.map((line) => {
        if (plataformaOptions.some((o) => o.value === line.platformKey)) return line;
        const hint = getPlataformaLabelByValue(line.platformKey);
        return {
          ...line,
          platformKey: PLATAFORMA_OTRO_VALUE,
          otroNombre: line.otroNombre.trim() || hint || '',
        };
      })
    );
  }, [sheet, plataformaOptions]);

  function openMoneda() {
    setDraftMoneda(moneda);
    setSheet('moneda');
  }

  function openCuenta(cuentaId) {
    if (cuentaId === 'tarjetaCredito') {
      openTarjetasCredito();
      return;
    }
    if (cuentaId === 'banco') {
      const opts = getBancosOptionsForMoneda(moneda);
      let base = bancosDetalle.length > 0 ? [...bancosDetalle] : [];
      if (!base.length && (parseFloat(saldos.banco) || 0) > 0) {
        base = [
          {
            id: newBancoLineId(),
            nombre: 'Cuenta bancaria',
            saldo: parseFloat(saldos.banco) || 0,
          },
        ];
      }
      const drafts =
        base.length > 0 ? base.map((r) => persistedToDraft(r, opts)) : [emptyBancoDraftLine(opts)];
      setBancoModalLines(drafts);
      setSheet({ type: 'cuenta', id: 'banco' });
      return;
    }
    setDraftMonto(String(saldos[cuentaId] ?? ''));
    setSheet({ type: 'cuenta', id: cuentaId });
  }

  function openTarjetasCredito() {
    let list = tarjetasCredito.length > 0 ? [...tarjetasCredito] : [];
    if (!list.length && (parseFloat(limiteTc) || 0) > 0) {
      list = [
        {
          id: generarIdTarjetaCredito(),
          nombreEntidad: '',
          tasaEA: 0,
          cupoTotal: parseFloat(limiteTc) || 0,
          cupoUtilizado: 0,
          fechaHoraCorte: defaultISOFromLegacyDia(15, 'corte'),
          fechaHoraLimitePago: defaultISOFromLegacyDia(5, 'pago'),
        },
      ];
    }
    setTcModalLines(list.length > 0 ? list.map(tcToDraft) : [emptyTcDraft()]);
    setSheet('tarjetasCredito');
  }

  function patchTcLine(index, patch) {
    setTcModalLines((prev) => prev.map((line, i) => (i === index ? { ...line, ...patch } : line)));
  }

  function addTcLine() {
    setTcModalLines((prev) => [...prev, emptyTcDraft()]);
  }

  function removeTcLine(index) {
    setTcModalLines((prev) => prev.filter((_, i) => i !== index));
  }

  function applyTarjetasCredito() {
    /** El cupo libre sale de cupo total − deuda (filas y movimientos); no hace falta un «saldo inicial» aparte. */
    const montoTarjetaCta = 0;
    if (tcModalLines.length === 0) {
      setTarjetasCredito([]);
      setLimiteTc('0');
      setSaldos((prev) => ({ ...prev, tarjetaCredito: '0' }));
      replaceState((s) => ({
        ...s,
        tarjetasCredito: [],
        limiteTarjetaCredito: 0,
        saldosCuentas: {
          ...emptySaldosCuentas(),
          ...(s.saldosCuentas && typeof s.saldosCuentas === 'object' ? s.saldosCuentas : {}),
          tarjetaCredito: montoTarjetaCta,
        },
        pagosProgramados: reemplazarPagosRecordatorioTarjetas(
          s.pagosProgramados || [],
          { ...s, tarjetasCredito: [] }
        ),
      }));
      closeSheet();
      return;
    }
    const cleaned = [];
    for (let i = 0; i < tcModalLines.length; i++) {
      const row = tcModalLines[i];
      if (!row.nombreEntidad.trim()) {
        Alert.alert(
          'Tarjeta de crédito',
          `En la fila ${i + 1}, escribe el nombre de la entidad (banco o emisor).`
        );
        return;
      }
      if (!String(row.fechaHoraCorte || '').trim() || !parseFechaHoraLocal(row.fechaHoraCorte)) {
        Alert.alert('Tarjeta de crédito', `En la fila ${i + 1}, elige una fecha de corte válida.`);
        return;
      }
      if (!String(row.fechaHoraLimitePago || '').trim() || !parseFechaHoraLocal(row.fechaHoraLimitePago)) {
        Alert.alert('Tarjeta de crédito', `En la fila ${i + 1}, elige una fecha límite de pago válida.`);
        return;
      }
      cleaned.push(draftToTc(row));
    }
    const sumCupos = cleaned.reduce((s, t) => s + t.cupoTotal, 0);
    const tarjetasClean = cleaned.map((t) => {
      const fc = soloFechaGuardada(t.fechaHoraCorte);
      const fl = soloFechaGuardada(t.fechaHoraLimitePago);
      const corteOk = parseFechaHoraLocal(fc);
      const limOk = parseFechaHoraLocal(fl);
      return {
        id: t.id || generarIdTarjetaCredito(),
        nombreEntidad: String(t.nombreEntidad || '').trim(),
        tasaEA: parseFloat(t.tasaEA) || 0,
        cupoTotal: parseFloat(t.cupoTotal) || 0,
        cupoUtilizado: parseFloat(t.cupoUtilizado) || 0,
        fechaHoraCorte: corteOk ? fc : '',
        fechaHoraLimitePago: limOk ? fl : '',
        diaCorte: corteOk ? Math.min(28, corteOk.getDate()) : 15,
        diaLimitePago: limOk ? Math.min(28, limOk.getDate()) : 5,
      };
    });
    setTarjetasCredito(tarjetasClean);
    setLimiteTc(String(sumCupos));
    setSaldos((prev) => ({ ...prev, tarjetaCredito: '0' }));
    replaceState((s) => ({
      ...s,
      tarjetasCredito: tarjetasClean,
      limiteTarjetaCredito: sumCupos,
      saldosCuentas: {
        ...emptySaldosCuentas(),
        ...(s.saldosCuentas && typeof s.saldosCuentas === 'object' ? s.saldosCuentas : {}),
        tarjetaCredito: montoTarjetaCta,
      },
      pagosProgramados: reemplazarPagosRecordatorioTarjetas(
        s.pagosProgramados || [],
        { ...s, tarjetasCredito: tarjetasClean }
      ),
    }));
    closeSheet();
  }

  function patchBancoLine(index, patch) {
    setBancoModalLines((prev) =>
      prev.map((line, i) => (i === index ? { ...line, ...patch } : line))
    );
  }

  function addBancoLine() {
    const opts = getBancosOptionsForMoneda(moneda);
    setBancoModalLines((prev) => [...prev, emptyBancoDraftLine(opts)]);
  }

  function removeBancoLine(index) {
    setBancoModalLines((prev) => (prev.length <= 1 ? prev : prev.filter((_, i) => i !== index)));
  }

  function openPlataformas() {
    const opts = getPlataformasOptions();
    let base = plataformasDetalle.length > 0 ? [...plataformasDetalle] : [];
    if (!base.length) {
      const n = parseFloat(saldos.nequi) || 0;
      const d = parseFloat(saldos.daviplata) || 0;
      const b = parseFloat(saldos.billeteras) || 0;
      if (n > 0) {
        base.push({ id: newPlataformaLineId(), platformValue: 'nequi', nombre: 'Nequi', saldo: n });
      }
      if (d > 0) {
        base.push({
          id: newPlataformaLineId(),
          platformValue: 'daviplata',
          nombre: 'Daviplata',
          saldo: d,
        });
      }
      if (b > 0) {
        base.push({
          id: newPlataformaLineId(),
          platformValue: PLATAFORMA_OTRO_VALUE,
          nombre: 'Otras plataformas',
          saldo: b,
        });
      }
    }
    const drafts =
      base.length > 0 ? base.map((r) => persistedPlataformaToDraft(r, opts)) : [emptyPlataformaDraftLine(opts)];
    setPlataformaModalLines(drafts);
    setSheet('plataformas');
  }

  function patchPlataformaLine(index, patch) {
    setPlataformaModalLines((prev) =>
      prev.map((line, i) => (i === index ? { ...line, ...patch } : line))
    );
  }

  function addPlataformaLine() {
    const opts = getPlataformasOptions();
    setPlataformaModalLines((prev) => [...prev, emptyPlataformaDraftLine(opts)]);
  }

  function removePlataformaLine(index) {
    setPlataformaModalLines((prev) =>
      prev.length <= 1 ? prev : prev.filter((_, i) => i !== index)
    );
  }

  function applyPlataformas() {
    const opts = getPlataformasOptions();
    for (let i = 0; i < plataformaModalLines.length; i++) {
      const row = plataformaModalLines[i];
      if (row.platformKey === PLATAFORMA_OTRO_VALUE && !row.otroNombre.trim()) {
        Alert.alert(
          'Plataforma',
          'En «Otro», escribe el nombre de la plataforma o elige una de la lista.'
        );
        return;
      }
    }
    const detalle = plataformaModalLines.map((row) => draftPlataformaToPersisted(row, opts));
    const sums = { nequi: 0, daviplata: 0, billeteras: 0 };
    detalle.forEach((r) => {
      const cu = cuentaSaldoPlataforma(r.platformValue || PLATAFORMA_OTRO_VALUE);
      sums[cu] += parseFloat(r.saldo) || 0;
    });
    setPlataformasDetalle(detalle);
    setSaldos((prev) => ({
      ...prev,
      nequi: String(sums.nequi),
      daviplata: String(sums.daviplata),
      billeteras: String(sums.billeteras),
    }));
    closeSheet();
  }

  function openPresupuesto() {
    setDraftPresupuesto(presupuesto);
    setSheet('presupuesto');
  }

  function openNota() {
    setDraftNota(nota);
    setSheet('nota');
  }

  function applyMoneda() {
    setMoneda(draftMoneda);
    closeSheet();
  }

  function applyCuenta() {
    if (sheet?.type !== 'cuenta') return;
    const id = sheet.id;
    if (id === 'banco') {
      const opts = getBancosOptionsForMoneda(moneda);
      for (let i = 0; i < bancoModalLines.length; i++) {
        const row = bancoModalLines[i];
        if (row.bancoKey === BANCO_OTRO_VALUE && !row.otroNombre.trim()) {
          Alert.alert('Banco', 'En la fila «Otro», escribe el nombre del banco o elige uno de la lista.');
          return;
        }
      }
      const detalle = bancoModalLines.map((row) => draftToPersisted(row, opts));
      const total = totalSaldoBancosDetalle(detalle);
      setBancosDetalle(detalle);
      setSaldos((prev) => ({ ...prev, banco: String(total) }));
      closeSheet();
      return;
    }
    setSaldos((prev) => ({ ...prev, [id]: draftMonto }));
    closeSheet();
  }

  function applyPresupuesto() {
    setPresupuesto(draftPresupuesto);
    closeSheet();
  }

  function applyNota() {
    setNota(draftNota);
    closeSheet();
  }

  function guardar() {
    if (!moneda) {
      Alert.alert('Moneda', 'Selecciona un tipo de moneda (toca la tarjeta «Moneda base»).');
      return;
    }
    const sc = { ...emptySaldosCuentas() };
    CUENTAS.forEach((c) => {
      sc[c.id] = parseFloat(saldos[c.id]) || 0;
    });
    if (bancosDetalle.length > 0) {
      sc.banco = totalSaldoBancosDetalle(bancosDetalle);
    }
    if (plataformasDetalle.length > 0) {
      const sums = { nequi: 0, daviplata: 0, billeteras: 0 };
      plataformasDetalle.forEach((r) => {
        const cu = cuentaSaldoPlataforma(r.platformValue || PLATAFORMA_OTRO_VALUE);
        sums[cu] += parseFloat(r.saldo) || 0;
      });
      sc.nequi = sums.nequi;
      sc.daviplata = sums.daviplata;
      sc.billeteras = sums.billeteras;
    }
    const bancosClean = bancosDetalle.map((r) => ({
      id: r.id || newBancoLineId(),
      nombre: String(r.nombre || '').trim() || 'Banco',
      saldo: parseFloat(r.saldo) || 0,
    }));
    const plataformasClean = plataformasDetalle.map((r) => ({
      id: r.id || newPlataformaLineId(),
      platformValue: r.platformValue || PLATAFORMA_OTRO_VALUE,
      nombre: String(r.nombre || '').trim() || 'Plataforma',
      saldo: parseFloat(r.saldo) || 0,
    }));
    const tarjetasClean = (tarjetasCredito || []).map((t) => {
      const fc = soloFechaGuardada(t.fechaHoraCorte);
      const fl = soloFechaGuardada(t.fechaHoraLimitePago);
      const corteOk = parseFechaHoraLocal(fc);
      const limOk = parseFechaHoraLocal(fl);
      return {
        id: t.id || generarIdTarjetaCredito(),
        nombreEntidad: String(t.nombreEntidad || '').trim(),
        tasaEA: parseFloat(t.tasaEA) || 0,
        cupoTotal: parseFloat(t.cupoTotal) || 0,
        cupoUtilizado: parseFloat(t.cupoUtilizado) || 0,
        fechaHoraCorte: corteOk ? fc : '',
        fechaHoraLimitePago: limOk ? fl : '',
        diaCorte: corteOk ? Math.min(28, corteOk.getDate()) : 15,
        diaLimitePago: limOk ? Math.min(28, limOk.getDate()) : 5,
      };
    });
    const limitePersist =
      tarjetasClean.length > 0
        ? tarjetasClean.reduce((s, t) => s + t.cupoTotal, 0)
        : parseFloat(limiteTc) || 0;
    replaceState((s) => ({
      ...s,
      moneda,
      saldosCuentas: sc,
      bancosDetalle: bancosClean,
      plataformasDetalle: plataformasClean,
      tarjetasCredito: tarjetasClean,
      limiteTarjetaCredito: limitePersist,
      presupuestoMensual: parseFloat(presupuesto) || 0,
      saldoInicialNota: nota.trim(),
      pagosProgramados: reemplazarPagosRecordatorioTarjetas(
        s.pagosProgramados || [],
        { ...s, tarjetasCredito: tarjetasClean }
      ),
    }));
    Alert.alert('Guardado', 'Saldo inicial actualizado.');
  }

  const sheetTitle =
    sheet === 'moneda'
      ? 'Moneda base'
      : sheet === 'tarjetasCredito'
        ? 'Tarjetas de crédito'
        : sheet === 'plataformas'
          ? 'Mis plataformas'
          : sheet === 'presupuesto'
            ? 'Presupuesto mensual'
            : sheet === 'nota'
              ? 'Nota'
              : sheet?.type === 'cuenta'
                ? sheet.id === 'banco'
                  ? 'Cuentas en banco'
                  : (CUENTAS.find((c) => c.id === sheet.id)?.nombre ?? 'Cuenta')
                : '';

  const modalScrollTall =
    sheet === 'tarjetasCredito' ||
    (sheet?.type === 'cuenta' && sheet.id === 'banco') ||
    sheet === 'plataformas';

  return (
    <ScreenWrap contentStyle={{ paddingTop: spacing.xs }}>
      <HeaderConCampana
        label="Cuentas"
        title="Saldo inicial"
        subtitle="Toca una tarjeta para editar"
      />

      <View style={styles.cardStack}>
        <EditCard
          icon="cash-outline"
          title="Moneda base"
          subtitle={moneda ? monedaLabel : 'Toca para elegir moneda'}
          hint={!moneda ? 'Requerido para guardar' : null}
          onPress={openMoneda}
        />

        {CUENTAS.filter((c) => !CUENTA_IDS_CARD_PLATAFORMAS.has(c.id)).map((c) => {
          if (c.id === 'banco') {
            const sub =
              totalBancoMostrado > 0 || bancosDetalle.length > 0
                ? `${formatearNumero(totalBancoMostrado)} ${moneda || '—'}`
                : 'Toca para ingresar saldo inicial';
            const hintB =
              bancosDetalle.length > 0
                ? bancosDetalle.map((r) => r.nombre).join(' · ')
                : null;
            return (
              <EditCard
                key={c.id}
                icon={CUENTA_ICONS.banco}
                title={c.nombre}
                subtitle={sub}
                hint={hintB}
                onPress={() => openCuenta('banco')}
              />
            );
          }
          const raw = saldos[c.id];
          const num = parseFloat(raw) || 0;
          const subTarjeta =
            `${formatearNumero(saldosActuales.tarjetaCredito || 0)} ${moneda || '—'}`;
          const sub =
            c.id === 'tarjetaCredito'
              ? subTarjeta
              : raw !== '' && raw !== undefined && String(raw).trim() !== ''
                ? `${formatearNumero(num)} ${moneda || '—'}`
                : 'Toca para ingresar saldo inicial';
          const hint = c.id === 'tarjetaCredito' ? tcTarjetaHint : null;
          return (
            <EditCard
              key={c.id}
              icon={CUENTA_ICONS[c.id] || 'ellipse-outline'}
              title={c.id === 'tarjetaCredito' ? `${c.nombre} · cupo libre` : c.nombre}
              subtitle={sub}
              hint={hint}
              onPress={() => openCuenta(c.id)}
            />
          );
        })}

        <EditCard
          icon="apps-outline"
          title="Mis plataformas"
          subtitle={
            totalPlataformasMostrado > 0 || plataformasDetalle.length > 0
              ? `${formatearNumero(totalPlataformasMostrado)} ${moneda || '—'}`
              : 'Nequi, Daviplata y más · toca para editar'
          }
          hint={
            plataformasDetalle.length > 0
              ? plataformasDetalle.map((r) => r.nombre).join(' · ')
              : null
          }
          onPress={openPlataformas}
        />

        <EditCard
          icon="pie-chart-outline"
          title="Presupuesto mensual"
          subtitle={
            presupuesto && String(presupuesto).trim() !== '' && parseFloat(presupuesto) > 0
              ? `${formatearNumero(parseFloat(presupuesto))} ${moneda || ''}`.trim()
              : 'Opcional · toca para definir'
          }
          onPress={openPresupuesto}
        />

        <EditCard
          icon="document-text-outline"
          title="Nota"
          subtitle={nota.trim() ? nota.trim() : 'Opcional · comentario o referencia'}
          onPress={openNota}
        />
      </View>

      <PrimaryButton title="Guardar saldo inicial" onPress={guardar} style={{ marginBottom: spacing.md }} />

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

      <Modal visible={sheet !== null} animationType="slide" transparent onRequestClose={closeSheet}>
        <View style={styles.modalOverlay}>
          <Pressable style={StyleSheet.absoluteFill} onPress={closeSheet} accessibilityLabel="Cerrar" />
          <KeyboardAvoidingView
            behavior={Platform.OS === 'ios' ? 'padding' : undefined}
            style={styles.modalAvoid}
          >
            <Animated.View
              style={[styles.modalSheet, { transform: [{ translateY: sheetDragY }] }]}
              {...sheetPanResponder.panHandlers}
              accessibilityLabel="Arrastra hacia abajo para cerrar"
            >
              <View style={styles.modalHeader}>
                <View style={styles.modalHandle} />
                <Text style={styles.modalTitle}>{sheetTitle}</Text>
              </View>

              <ScrollView
                keyboardShouldPersistTaps="handled"
                showsVerticalScrollIndicator={false}
                style={modalScrollTall ? styles.modalScrollTall : styles.modalScroll}
                onScroll={(e) => {
                  sheetScrollY.current = e.nativeEvent.contentOffset.y;
                }}
                scrollEventThrottle={16}
              >
                {sheet === 'moneda' && (
                  <>
                    <Text style={styles.modalLab}>Moneda</Text>
                    <View style={styles.pickerWrap}>
                      <Picker
                        selectedValue={draftMoneda}
                        onValueChange={setDraftMoneda}
                        style={{ color: colors.text }}
                      >
                        {MONEDAS.map((m) => (
                          <Picker.Item key={m.value || 'x'} label={m.label} value={m.value} />
                        ))}
                      </Picker>
                    </View>
                  </>
                )}

                {sheet?.type === 'cuenta' && sheet.id === 'banco' && (
                  <>
                    <Text style={styles.modalHint}>
                      {moneda
                        ? `Bancos sugeridos para ${moneda}. `
                        : 'Elige una moneda base arriba para ver bancos por país. '}
                      Usa + para añadir otra cuenta. El total «Banco» es la suma de todas.
                    </Text>
                    {bancoModalLines.map((line, idx) => (
                      <View key={line.id} style={styles.bancoBlock}>
                        <View style={styles.bancoBlockHead}>
                          <Text style={styles.bancoBlockTitle}>Cuenta {idx + 1}</Text>
                          {bancoModalLines.length > 1 ? (
                            <TouchableOpacity
                              onPress={() => removeBancoLine(idx)}
                              hitSlop={10}
                              accessibilityLabel="Quitar banco"
                            >
                              <Ionicons name="trash-outline" size={22} color={colors.danger} />
                            </TouchableOpacity>
                          ) : null}
                        </View>
                        <Text style={styles.modalLab}>Banco</Text>
                        <View style={styles.pickerWrap}>
                          <Picker
                            selectedValue={
                              bancoOptions.some((b) => b.value === line.bancoKey)
                                ? line.bancoKey
                                : BANCO_OTRO_VALUE
                            }
                            onValueChange={(v) => patchBancoLine(idx, { bancoKey: v })}
                            style={{ color: colors.text }}
                          >
                            {bancoOptions.map((b) => (
                              <Picker.Item key={b.value} label={b.label} value={b.value} />
                            ))}
                          </Picker>
                        </View>
                        {line.bancoKey === BANCO_OTRO_VALUE && (
                          <>
                            <Text style={[styles.modalLab, { marginTop: spacing.sm }]}>Nombre del banco</Text>
                            <TextInput
                              style={styles.input}
                              value={line.otroNombre}
                              onChangeText={(t) => patchBancoLine(idx, { otroNombre: t })}
                              placeholder="Ej. Banco XYZ"
                              placeholderTextColor={colors.textFaint}
                            />
                          </>
                        )}
                        <Text style={[styles.modalLab, { marginTop: spacing.sm }]}>Saldo en esta cuenta</Text>
                        <TextInput
                          style={styles.input}
                          keyboardType="decimal-pad"
                          value={line.monto}
                          onChangeText={(t) => patchBancoLine(idx, { monto: t })}
                          placeholder="0,00"
                          placeholderTextColor={colors.textFaint}
                        />
                      </View>
                    ))}
                    <TouchableOpacity style={styles.addBancoBtn} onPress={addBancoLine} activeOpacity={0.75}>
                      <View style={styles.addBancoCircle}>
                        <Ionicons name="add" size={26} color={colors.accentBright} />
                      </View>
                      <Text style={styles.addBancoTxt}>Agregar otro banco</Text>
                    </TouchableOpacity>
                  </>
                )}

                {sheet === 'plataformas' && (
                  <>
                    <Text style={styles.modalHint}>
                      Plataformas digitales en Colombia. Nequi y Daviplata van a sus propios saldos; el resto se
                      suma en «Otras plataformas» para tus gastos.
                    </Text>
                    {plataformaModalLines.map((line, idx) => (
                      <View key={line.id} style={styles.bancoBlock}>
                        <View style={styles.bancoBlockHead}>
                          <Text style={styles.bancoBlockTitle}>Plataforma {idx + 1}</Text>
                          {plataformaModalLines.length > 1 ? (
                            <TouchableOpacity
                              onPress={() => removePlataformaLine(idx)}
                              hitSlop={10}
                              accessibilityLabel="Quitar plataforma"
                            >
                              <Ionicons name="trash-outline" size={22} color={colors.danger} />
                            </TouchableOpacity>
                          ) : null}
                        </View>
                        <Text style={styles.modalLab}>App o billetera</Text>
                        <View style={styles.pickerWrap}>
                          <Picker
                            selectedValue={
                              plataformaOptions.some((o) => o.value === line.platformKey)
                                ? line.platformKey
                                : PLATAFORMA_OTRO_VALUE
                            }
                            onValueChange={(v) => patchPlataformaLine(idx, { platformKey: v })}
                            style={{ color: colors.text }}
                          >
                            {plataformaOptions.map((p) => (
                              <Picker.Item key={p.value} label={p.label} value={p.value} />
                            ))}
                          </Picker>
                        </View>
                        {line.platformKey === PLATAFORMA_OTRO_VALUE && (
                          <>
                            <Text style={[styles.modalLab, { marginTop: spacing.sm }]}>Nombre</Text>
                            <TextInput
                              style={styles.input}
                              value={line.otroNombre}
                              onChangeText={(t) => patchPlataformaLine(idx, { otroNombre: t })}
                              placeholder="Ej. otra app"
                              placeholderTextColor={colors.textFaint}
                            />
                          </>
                        )}
                        <Text style={[styles.modalLab, { marginTop: spacing.sm }]}>Saldo</Text>
                        <TextInput
                          style={styles.input}
                          keyboardType="decimal-pad"
                          value={line.monto}
                          onChangeText={(t) => patchPlataformaLine(idx, { monto: t })}
                          placeholder="0,00"
                          placeholderTextColor={colors.textFaint}
                        />
                      </View>
                    ))}
                    <TouchableOpacity
                      style={styles.addBancoBtn}
                      onPress={addPlataformaLine}
                      activeOpacity={0.75}
                    >
                      <View style={styles.addBancoCircle}>
                        <Ionicons name="add" size={26} color={colors.accentBright} />
                      </View>
                      <Text style={styles.addBancoTxt}>Agregar otra plataforma</Text>
                    </TouchableOpacity>
                  </>
                )}

                {sheet === 'tarjetasCredito' && (
                  <>
                    <Text style={styles.modalHint}>
                      El cupo libre se calcula a partir de cupo total, cupo utilizado (deuda) y tus movimientos en
                      Gastos. Los cupos por entidad se suman como límite total. Indica la fecha de corte y la fecha
                      límite de pago: cada mes se usará el mismo día del mes (si el mes tiene menos días, se ajusta al
                      último día). Al aplicar, se actualizan los recordatorios en «Pagos programados».
                    </Text>
                    {tcModalLines.length === 0 ? (
                      <Text style={[styles.modalHint, { marginBottom: spacing.sm }]}>
                        No hay filas: pulsa «Agregar otra tarjeta» o Aplicar para quitar toda la configuración de
                        tarjetas.
                      </Text>
                    ) : null}
                    {tcModalLines.map((line, idx) => (
                      <View key={line.id} style={styles.bancoBlock}>
                        <View style={styles.bancoBlockHead}>
                          <Text style={styles.bancoBlockTitle}>Tarjeta {idx + 1}</Text>
                          {tcModalLines.length > 1 ? (
                            <TouchableOpacity
                              onPress={() => removeTcLine(idx)}
                              hitSlop={10}
                              accessibilityLabel="Quitar tarjeta"
                            >
                              <Ionicons name="trash-outline" size={22} color={colors.danger} />
                            </TouchableOpacity>
                          ) : null}
                        </View>
                        <Text style={styles.modalLab}>Nombre de la entidad</Text>
                        <TextInput
                          style={styles.input}
                          value={line.nombreEntidad}
                          onChangeText={(t) => patchTcLine(idx, { nombreEntidad: t })}
                          placeholder="Ej. Banco X · Visa"
                          placeholderTextColor={colors.textFaint}
                        />
                        <Text style={[styles.modalLab, { marginTop: spacing.sm }]}>Tasa interés E.A. (%)</Text>
                        <TextInput
                          style={styles.input}
                          keyboardType="decimal-pad"
                          value={line.tasaEA}
                          onChangeText={(t) => patchTcLine(idx, { tasaEA: t })}
                          placeholder="Ej. 35,4"
                          placeholderTextColor={colors.textFaint}
                        />
                        <Text style={[styles.modalLab, { marginTop: spacing.sm }]}>Cupo total</Text>
                        <TextInput
                          style={styles.input}
                          keyboardType="decimal-pad"
                          value={line.cupoTotal}
                          onChangeText={(t) => patchTcLine(idx, { cupoTotal: t })}
                          placeholder="0"
                          placeholderTextColor={colors.textFaint}
                        />
                        <Text style={[styles.modalLab, { marginTop: spacing.sm }]}>Cupo utilizado (deuda)</Text>
                        <TextInput
                          style={styles.input}
                          keyboardType="decimal-pad"
                          value={line.cupoUtilizado}
                          onChangeText={(t) => patchTcLine(idx, { cupoUtilizado: t })}
                          placeholder="0"
                          placeholderTextColor={colors.textFaint}
                        />
                        <Text style={[styles.modalLab, { marginTop: spacing.sm }]}>Fecha de corte</Text>
                        <TouchableOpacity
                          style={styles.input}
                          onPress={() => setTcPicker({ idx, field: 'corte' })}
                          activeOpacity={0.75}
                        >
                          <Text style={{ color: colors.text, fontSize: 16 }}>
                            {parseFechaHoraLocal(line.fechaHoraCorte)
                              ? parseFechaHoraLocal(line.fechaHoraCorte).toLocaleDateString('es', {
                                  dateStyle: 'short',
                                })
                              : 'Elegir fecha'}
                          </Text>
                        </TouchableOpacity>
                        <Text style={[styles.modalLab, { marginTop: spacing.sm }]}>Fecha límite de pago</Text>
                        <TouchableOpacity
                          style={styles.input}
                          onPress={() => setTcPicker({ idx, field: 'limite' })}
                          activeOpacity={0.75}
                        >
                          <Text style={{ color: colors.text, fontSize: 16 }}>
                            {parseFechaHoraLocal(line.fechaHoraLimitePago)
                              ? parseFechaHoraLocal(line.fechaHoraLimitePago).toLocaleDateString('es', {
                                  dateStyle: 'short',
                                })
                              : 'Elegir fecha'}
                          </Text>
                        </TouchableOpacity>
                      </View>
                    ))}
                    {tcPicker != null && tcModalLines[tcPicker.idx] ? (
                      <DateTimePicker
                        value={
                          parseFechaHoraLocal(
                            tcPicker.field === 'corte'
                              ? tcModalLines[tcPicker.idx].fechaHoraCorte
                              : tcModalLines[tcPicker.idx].fechaHoraLimitePago
                          ) || new Date()
                        }
                        mode="date"
                        display={Platform.OS === 'ios' ? 'spinner' : 'default'}
                        onChange={(ev, d) => {
                          const pick = tcPicker;
                          if (Platform.OS !== 'ios') setTcPicker(null);
                          if (ev.type === 'dismissed') setTcPicker(null);
                          if (d && pick != null && tcModalLines[pick.idx]) {
                            const key =
                              pick.field === 'corte' ? 'fechaHoraCorte' : 'fechaHoraLimitePago';
                            patchTcLine(pick.idx, { [key]: fechaALocalISO(d) });
                          }
                        }}
                      />
                    ) : null}
                    <TouchableOpacity style={styles.addBancoBtn} onPress={addTcLine} activeOpacity={0.75}>
                      <View style={styles.addBancoCircle}>
                        <Ionicons name="add" size={26} color={colors.accentBright} />
                      </View>
                      <Text style={styles.addBancoTxt}>Agregar otra tarjeta</Text>
                    </TouchableOpacity>
                  </>
                )}

                {sheet?.type === 'cuenta' && sheet.id !== 'banco' && (
                  <>
                    <Text style={styles.modalLab}>Saldo inicial</Text>
                    <TextInput
                      style={styles.input}
                      keyboardType="decimal-pad"
                      value={draftMonto}
                      onChangeText={setDraftMonto}
                      placeholder="0,00"
                      placeholderTextColor={colors.textFaint}
                    />
                  </>
                )}

                {sheet === 'presupuesto' && (
                  <>
                    <Text style={styles.modalLab}>Monto mensual</Text>
                    <TextInput
                      style={styles.input}
                      keyboardType="decimal-pad"
                      value={draftPresupuesto}
                      onChangeText={setDraftPresupuesto}
                      placeholder="0 — opcional"
                      placeholderTextColor={colors.textFaint}
                    />
                  </>
                )}

                {sheet === 'nota' && (
                  <>
                    <Text style={styles.modalLab}>Texto</Text>
                    <TextInput
                      style={[styles.input, styles.inputMultiline]}
                      value={draftNota}
                      onChangeText={setDraftNota}
                      placeholder="Ej. saldo al 1 de enero"
                      placeholderTextColor={colors.textFaint}
                      multiline
                    />
                  </>
                )}
              </ScrollView>

              <PrimaryButton
                title="Aplicar"
                onPress={() => {
                  if (sheet === 'moneda') applyMoneda();
                  else if (sheet === 'tarjetasCredito') applyTarjetasCredito();
                  else if (sheet === 'plataformas') applyPlataformas();
                  else if (sheet?.type === 'cuenta') applyCuenta();
                  else if (sheet === 'presupuesto') applyPresupuesto();
                  else if (sheet === 'nota') applyNota();
                }}
                style={{ marginTop: spacing.md }}
              />
              <GhostButton title="Cancelar" onPress={closeSheet} style={{ marginTop: spacing.sm }} />
            </Animated.View>
          </KeyboardAvoidingView>
        </View>
      </Modal>
    </ScreenWrap>
  );
}

const styles = StyleSheet.create({
  cardStack: {
    marginBottom: spacing.md,
  },
  editCard: {
    flexDirection: 'row',
    alignItems: 'center',
    backgroundColor: colors.surface,
    borderRadius: radii.lg,
    padding: spacing.md,
    marginBottom: spacing.sm,
    borderWidth: 1,
    borderColor: colors.stroke,
  },
  editCardIcon: {
    width: 44,
    height: 44,
    borderRadius: radii.md,
    backgroundColor: colors.surfaceHighlight,
    alignItems: 'center',
    justifyContent: 'center',
    marginRight: spacing.md,
    borderWidth: 1,
    borderColor: colors.stroke,
  },
  editCardBody: {
    flex: 1,
    minWidth: 0,
  },
  editCardTitle: {
    color: colors.text,
    fontSize: 16,
    fontWeight: '700',
    letterSpacing: -0.2,
  },
  editCardSub: {
    color: colors.textMuted,
    fontSize: 13,
    marginTop: 4,
    lineHeight: 18,
  },
  editCardHint: {
    color: colors.textFaint,
    fontSize: 11,
    marginTop: 4,
  },
  modalOverlay: {
    flex: 1,
    backgroundColor: 'rgba(0,0,0,0.55)',
    justifyContent: 'flex-end',
  },
  modalAvoid: {
    width: '100%',
    maxHeight: '92%',
  },
  modalSheet: {
    backgroundColor: colors.surfaceSolid,
    borderTopLeftRadius: radii.xl,
    borderTopRightRadius: radii.xl,
    paddingHorizontal: spacing.lg,
    paddingTop: spacing.sm,
    paddingBottom: spacing.xl,
    borderTopWidth: 1,
    borderColor: colors.stroke,
  },
  modalHeader: {
    marginBottom: 0,
  },
  modalHandle: {
    alignSelf: 'center',
    width: 40,
    height: 4,
    borderRadius: 2,
    backgroundColor: colors.strokeStrong,
    marginBottom: spacing.md,
  },
  modalTitle: {
    ...typography.title,
    marginBottom: spacing.md,
  },
  modalScroll: {
    maxHeight: 360,
  },
  modalScrollTall: {
    maxHeight: 520,
  },
  modalLab: {
    ...typography.label,
    marginBottom: spacing.xs,
    color: colors.textMuted,
    letterSpacing: 0.8,
  },
  modalHint: {
    ...typography.small,
    marginBottom: spacing.md,
    color: colors.textFaint,
    lineHeight: 20,
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
  inputMultiline: {
    minHeight: 100,
    textAlignVertical: 'top',
  },
  pickerWrap: {
    borderWidth: 1,
    borderColor: colors.stroke,
    borderRadius: radii.md,
    overflow: 'hidden',
    backgroundColor: 'rgba(0,0,0,0.12)',
  },
  bancoBlock: {
    marginBottom: spacing.lg,
    paddingBottom: spacing.md,
    borderBottomWidth: 1,
    borderBottomColor: colors.stroke,
  },
  bancoBlockHead: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    marginBottom: spacing.sm,
  },
  bancoBlockTitle: {
    color: colors.accent,
    fontSize: 12,
    fontWeight: '700',
    letterSpacing: 1,
    textTransform: 'uppercase',
  },
  addBancoBtn: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'center',
    gap: spacing.sm,
    paddingVertical: spacing.md,
    marginBottom: spacing.sm,
    borderWidth: 1,
    borderColor: colors.strokeStrong,
    borderRadius: radii.md,
    borderStyle: 'dashed',
    backgroundColor: colors.surfaceHighlight,
  },
  addBancoCircle: {
    width: 40,
    height: 40,
    borderRadius: 20,
    backgroundColor: 'rgba(139, 92, 246, 0.2)',
    alignItems: 'center',
    justifyContent: 'center',
  },
  addBancoTxt: {
    color: colors.accentBright,
    fontWeight: '700',
    fontSize: 15,
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
