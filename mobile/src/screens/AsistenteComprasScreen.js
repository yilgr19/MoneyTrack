import React, { useEffect, useMemo, useState, useCallback } from 'react';
import {
  View,
  Text,
  StyleSheet,
  TextInput,
  TouchableOpacity,
  ScrollView,
  Alert,
  Platform,
} from 'react-native';
import { Ionicons } from '@expo/vector-icons';
import { Picker } from '@react-native-picker/picker';
import { rootNavigationRef } from '../navigation/rootNavigationRef';
import ScreenWrap from '../components/ScreenWrap';
import UICard from '../components/UICard';
import { PrimaryButton, GhostButton } from '../components/Buttons';
import { useApp } from '../context/AppContext';
import {
  normalizarCategoria,
  formatearNumero,
  obtenerCuentasOrigenGastoElegible,
  calcularSaldosPorCuenta,
} from '../lib/finance';
import {
  generarIdIntencionCompra,
  costoPorSesion,
  mensajeCostoPorUso,
  datosTermometroCategoria,
  estimarListaSuperDesdeHistorial,
  elegirCategoriaSuperPorDefecto,
  PRECIO_REF_CINE_DEFAULT,
  generarIdListaSuperLinea,
  ordenarLineasListaSuper,
  URGENCIA_LISTA_SUPER,
} from '../lib/asistenteComprasLogic';
import { colors, spacing, radii, typography } from '../theme';

const TABS = [
  { id: 'intencion', label: 'Intención' },
  { id: 'deseos', label: 'Deseos' },
  { id: 'super', label: 'Súper / básicos' },
];

const LISTA_SUPER_BASE = [
  'Leche',
  'Huevos',
  'Pan',
  'Papel higiénico',
  'Detergente',
  'Fruta',
  'Verdura',
  'Aceite',
];

function pad(n) {
  return String(n).padStart(2, '0');
}

function formatCountdown(ms) {
  if (ms <= 0) return '00:00:00';
  const sTotal = Math.floor(ms / 1000);
  const h = Math.floor(sTotal / 3600);
  const m = Math.floor((sTotal % 3600) / 60);
  const s = sTotal % 60;
  return `${pad(h)}:${pad(m)}:${pad(s)}`;
}

function puedeRegistrarCompraPorRegla48h(intencion, ahora) {
  if (!intencion || intencion.estado !== 'pendiente') return false;
  if (!intencion.aplicabaCooldown) return true;
  const hasta = intencion.cooldownHasta;
  if (hasta == null) return true;
  return ahora >= hasta;
}

function normNombre(n) {
  return String(n || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

function etiquetaUrgencia(u) {
  if (u === 'urgente') return 'Urgente · comprar cuanto antes';
  if (u === 'puede_esperar') return 'Puede esperar';
  return 'Prioridad normal';
}

function colorUrgenciaEtiqueta(u) {
  if (u === 'urgente') return colors.danger;
  if (u === 'puede_esperar') return colors.textMuted;
  return colors.chartBlue;
}

export default function AsistenteComprasScreen({ route }) {
  const { state, replaceState } = useApp();
  const moneda = state?.moneda || '';
  const [tab, setTab] = useState('intencion');
  const [ahora, setAhora] = useState(Date.now());

  const [nombreDraft, setNombreDraft] = useState('');
  const [precioDraft, setPrecioDraft] = useState('');
  const [categoriaDraft, setCategoriaDraft] = useState('');
  const [editId, setEditId] = useState(null);

  const [listaSuperCat, setListaSuperCat] = useState('');
  const [articuloNuevo, setArticuloNuevo] = useState('');

  const umbral =
    state?.asistenteUmbral48h != null && !Number.isNaN(parseFloat(state.asistenteUmbral48h))
      ? Math.max(0, parseFloat(state.asistenteUmbral48h))
      : 50;
  const [umbralTexto, setUmbralTexto] = useState(String(umbral));

  const categorias = useMemo(
    () => (state?.categorias || []).map(normalizarCategoria),
    [state?.categorias]
  );

  const intenciones = state?.intencionesCompra || [];
  const editando = useMemo(
    () => intenciones.find((i) => i && i.id === editId) || null,
    [intenciones, editId]
  );

  useEffect(() => {
    setUmbralTexto(String(umbral));
  }, [umbral]);

  useEffect(() => {
    const t = setInterval(() => setAhora(Date.now()), 1000);
    return () => clearInterval(t);
  }, []);

  useEffect(() => {
    if (!state) return;
    const def = elegirCategoriaSuperPorDefecto(state);
    if (def && !listaSuperCat) setListaSuperCat(def);
  }, [state, listaSuperCat]);

  const tabRoute = route?.params?.tab;
  useEffect(() => {
    if (tabRoute === 'super' || tabRoute === 'intencion' || tabRoute === 'deseos') {
      setTab(tabRoute);
    }
  }, [tabRoute]);

  const lineasSuper = state?.listaSuperCompraItems || [];

  const lineaPorNombre = useCallback(
    (nombre) => {
      const k = normNombre(nombre);
      return lineasSuper.find((l) => normNombre(l.nombre) === k) || null;
    },
    [lineasSuper]
  );

  function toggleArticuloEnLista(nombre) {
    const k = normNombre(nombre);
    const exist = lineasSuper.find((l) => normNombre(l.nombre) === k);
    replaceState((s) => {
      const prev = s.listaSuperCompraItems || [];
      if (exist) {
        return { ...s, listaSuperCompraItems: prev.filter((l) => l.id !== exist.id) };
      }
      return {
        ...s,
        listaSuperCompraItems: [
          ...prev,
          { id: generarIdListaSuperLinea(), nombre: String(nombre).trim(), urgencia: 'normal' },
        ],
      };
    });
  }

  function setUrgenciaLinea(idLinea, urgencia) {
    replaceState((s) => ({
      ...s,
      listaSuperCompraItems: (s.listaSuperCompraItems || []).map((l) =>
        l.id === idLinea ? { ...l, urgencia } : l
      ),
    }));
  }

  /** Quita la línea porque ya compraste ese artículo */
  function marcarCompradoSuper(idLinea) {
    replaceState((s) => ({
      ...s,
      listaSuperCompraItems: (s.listaSuperCompraItems || []).filter((l) => l.id !== idLinea),
    }));
  }

  function guardarUmbral() {
    const n = Math.max(0, parseFloat(String(umbralTexto).replace(',', '.')) || 0);
    replaceState((s) => ({ ...s, asistenteUmbral48h: n }));
    Alert.alert('Listo', `Umbral 48 h: ${formatearNumero(n)} ${moneda}.`);
  }

  function crearIntencion() {
    const nombre = nombreDraft.trim();
    const precio = parseFloat(String(precioDraft).replace(',', '.')) || 0;
    const cat = String(categoriaDraft || '').trim();
    if (!nombre || precio <= 0 || !cat) {
      Alert.alert('Datos', 'Nombre, precio estimado y categoría son obligatorios.');
      return;
    }
    const precioNum = precio;
    const aplicabaCooldown = precioNum >= umbral;
    const cooldownHasta = aplicabaCooldown ? Date.now() + 48 * 3600000 : null;
    const id = generarIdIntencionCompra();
    const nueva = {
      id,
      nombre,
      precioEstimado: precioNum,
      nombreCategoria: cat,
      vecesPorSemana: 3,
      minutosPorSesion: 120,
      añosUso: 3,
      creadoEn: Date.now(),
      aplicabaCooldown,
      cooldownHasta,
      estado: 'pendiente',
    };
    replaceState((s) => ({
      ...s,
      intencionesCompra: [nueva, ...((s.intencionesCompra || []).filter((x) => x.estado === 'pendiente'))],
    }));
    setNombreDraft('');
    setPrecioDraft('');
    setCategoriaDraft('');
    setEditId(id);
    setTab('intencion');
  }

  function actualizarIntencion(id, patch) {
    replaceState((s) => ({
      ...s,
      intencionesCompra: (s.intencionesCompra || []).map((i) =>
        i.id === id ? { ...i, ...patch } : i
      ),
    }));
  }

  function registrarGastoDesdeIntencion(intencion, origenValor) {
    if (!puedeRegistrarCompraPorRegla48h(intencion, Date.now())) {
      Alert.alert('Regla 48 h', 'Aún debes esperar antes de registrar esta compra.');
      return;
    }
    const precio = intencion.precioEstimado;
    const opts = obtenerCuentasOrigenGastoElegible(state || {}, precio, precio, {});
    if (opts.length === 0) {
      Alert.alert('Saldo', 'No hay ninguna cuenta con saldo suficiente para este monto.');
      return;
    }
    if (!origenValor && opts.length > 1) {
      Alert.alert(
        'Cuenta',
        '¿Desde qué caja pagas?',
        [
          ...opts.map((o) => ({
            text: o.label.slice(0, 60),
            onPress: () => registrarGastoDesdeIntencion(intencion, o.value),
          })),
          { text: 'Cancelar', style: 'cancel' },
        ],
        { cancelable: true }
      );
      return;
    }
    const origen = origenValor || opts[0].value;
    const disponible = opts.find((o) => o.value === origen);
    if (!disponible || precio > (disponible.saldo || 0)) {
      Alert.alert('Saldo', 'No hay suficiente saldo en la cuenta elegida. Revisa en Saldo.');
      return;
    }
    const d = new Date();
    const fechaStr = `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(
      d.getMinutes()
    )}:00`;
    const nuevo = {
      nombre: String(intencion.nombre || '').trim(),
      cantidad: precio,
      fecha: fechaStr,
      categoria: String(intencion.nombreCategoria || '').trim(),
      origen,
      nota: 'Registrado desde Asistente de compras',
      cuotas: 1,
      cuotaMensual: precio,
    };
    replaceState((s) => ({
      ...s,
      gastos: [...(s.gastos || []), nuevo],
      intencionesCompra: (s.intencionesCompra || []).filter((x) => x.id !== intencion.id),
    }));
    Alert.alert('Listo', 'Compra registrada en tu historial.');
    if (editId === intencion.id) setEditId(null);
  }

  function yaNoLoQuiero(intencion) {
    const m = intencion.precioEstimado;
    const primera = (state?.metas || [])[0];
    const nombreMeta = primera ? String(primera.nombre || '').trim() : '';
    function quitar() {
      replaceState((s) => ({
        ...s,
        intencionesCompra: (s.intencionesCompra || []).filter((x) => x.id !== intencion.id),
      }));
      if (editId === intencion.id) setEditId(null);
    }
    const msg = primera
      ? `Acabas de ahorrar ${formatearNumero(m)} ${moneda}. Si tienes ese efectivo disponible, puedes aportar a tu meta «${nombreMeta}».`
      : `Acabas de ahorrar ${formatearNumero(m)} ${moneda}. Ese gasto imaginario ya no pesa en tu decisión; puedes reasignarlo cuando quieras.`;
    Alert.alert('¡Felicidades!', msg, [
      ...(primera
        ? [
            {
              text: `Aportar a «${nombreMeta.slice(0, 22)}${nombreMeta.length > 22 ? '…' : ''}»`,
              onPress: () => {
                const saldos = calcularSaldosPorCuenta(state || {});
                if ((saldos.efectivo || 0) < m) {
                  Alert.alert(
                    'Saldo insuficiente en efectivo',
                    'Cuando puedas, aporta manualmente desde Metas.'
                  );
                  quitar();
                  return;
                }
                const d = new Date();
                const fechaStr = `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
                replaceState((s) => ({
                  ...s,
                  intencionesCompra: (s.intencionesCompra || []).filter((x) => x.id !== intencion.id),
                  contribucionesMetas: [
                    ...(s.contribucionesMetas || []),
                    { metaId: primera.id, cantidad: m, fecha: fechaStr, origen: 'efectivo' },
                  ],
                }));
                if (editId === intencion.id) setEditId(null);
                Alert.alert('Listo', `Aporte de ${formatearNumero(m)} desde efectivo hacia «${nombreMeta}».`);
              },
            },
          ]
        : []),
      {
        text: primera ? 'Solo cerrar' : 'Gracias',
        style: 'cancel',
        onPress: quitar,
      },
    ]);
  }

  const analisis = useMemo(() => {
    if (!editando || !state)
      return {
        costoSes: 0,
        msgValor: '',
        termo: null,
        puede: true,
        bloqueo48: false,
        restanteMs: 0,
      };
    const ah = new Date();
    const mes = ah.getMonth();
    const año = ah.getFullYear();
    const precio = editando.precioEstimado;
    const cs = costoPorSesion(precio, editando.vecesPorSemana, editando.añosUso);
    const msgValor = mensajeCostoPorUso({
      nombreProducto: editando.nombre,
      precio,
      costoSesion: cs,
    });
    const termo = datosTermometroCategoria(state, editando.nombreCategoria, mes, año, precio);
    const puede = puedeRegistrarCompraPorRegla48h(editando, ahora);
    const bloqueo48 = editando.aplicabaCooldown && !puede;
    const restanteMs =
      editando.aplicabaCooldown && editando.cooldownHasta != null
        ? Math.max(0, editando.cooldownHasta - ahora)
        : 0;
    return { costoSes: cs, msgValor, termo, puede, bloqueo48, restanteMs };
  }, [editando, state, ahora]);

  const estimacionSuper = useMemo(() => {
    const cat = String(listaSuperCat || '').trim();
    const n = (state?.listaSuperCompraItems || []).length;
    if (!cat || n === 0) return null;
    return estimarListaSuperDesdeHistorial(state || {}, cat, n, 1);
  }, [state, listaSuperCat]);

  const articulosLista = useMemo(() => {
    const extra = state?.listaSuperArticulosExtra || [];
    const merged = [...LISTA_SUPER_BASE, ...extra];
    const seen = new Set();
    return merged.filter((x) => {
      const k = String(x).trim().toLowerCase();
      if (!k || seen.has(k)) return false;
      seen.add(k);
      return true;
    });
  }, [state?.listaSuperArticulosExtra]);

  function agregarArticuloExtra() {
    const t = articuloNuevo.trim();
    if (!t) return;
    if (lineaPorNombre(t)) {
      Alert.alert('Lista', 'Ese artículo ya está en tu lista.');
      return;
    }
    replaceState((s) => ({
      ...s,
      listaSuperArticulosExtra: [...new Set([...(s.listaSuperArticulosExtra || []), t])],
      listaSuperCompraItems: [
        ...(s.listaSuperCompraItems || []),
        { id: generarIdListaSuperLinea(), nombre: t, urgencia: 'normal' },
      ],
    }));
    setArticuloNuevo('');
  }

  function guardarPreferidaSuper() {
    replaceState((s) => ({ ...s, listaSuperCategoriaPreferida: String(listaSuperCat || '').trim() }));
    Alert.alert('Listo', 'Categoría preferida guardada para futuras listas.');
  }

  return (
    <ScreenWrap includeTopInset={false} scrollEnabled={false} contentStyle={{ paddingTop: spacing.xs }}>
      <Text style={typography.label}>Compras</Text>
      <Text style={typography.hero}>Asistente de compras</Text>
      <Text style={[typography.subtitle, { marginBottom: spacing.md }]}>
        Intenciones (sin gasto hasta que confirmes), valor y lista de súper guiada por tus registros.
      </Text>

      <View style={styles.tabRow}>
        {TABS.map((x) => (
          <TouchableOpacity
            key={x.id}
            style={[styles.tabBtn, tab === x.id && styles.tabBtnActive]}
            onPress={() => setTab(x.id)}
          >
            <Text style={[styles.tabLbl, tab === x.id && styles.tabLblActive]}>{x.label}</Text>
          </TouchableOpacity>
        ))}
      </View>

      <ScrollView
        style={{ flex: 1 }}
        showsVerticalScrollIndicator={false}
        keyboardShouldPersistTaps="handled"
        keyboardDismissMode="on-drag"
        automaticallyAdjustKeyboardInsets={Platform.OS !== 'web'}
      >
        {tab === 'intencion' && (
          <>
            <UICard style={{ marginBottom: spacing.md }}>
              <Text style={typography.label}>Intención de compra</Text>
              <Text style={[typography.small, { color: colors.textMuted, marginBottom: spacing.md, lineHeight: 20 }]}>
                Aún no es un gasto: solo dejas claro qué te tienta y cuánto cuesta aproximado.
              </Text>
              <Text style={styles.lab}>Nombre</Text>
              <TextInput
                style={styles.input}
                value={nombreDraft}
                onChangeText={setNombreDraft}
                placeholder="Ej. PlayStation 5"
                placeholderTextColor={colors.textFaint}
              />
              <Text style={styles.lab}>Precio estimado</Text>
              <TextInput
                style={styles.input}
                value={precioDraft}
                onChangeText={setPrecioDraft}
                keyboardType="decimal-pad"
                placeholder={moneda ? `Ej. 500 ${moneda}` : 'Ej. 500'}
                placeholderTextColor={colors.textFaint}
              />
              <Text style={styles.lab}>Categoría</Text>
              {categorias.length === 0 ? (
                <Text style={typography.small}>Crea categorías en Más → Categorías.</Text>
              ) : (
                <View style={styles.pickerWrap}>
                  <Picker
                    selectedValue={categoriaDraft}
                    onValueChange={setCategoriaDraft}
                    style={{ color: colors.text }}
                    dropdownIconColor={colors.text}
                  >
                    <Picker.Item label="— Elegir —" value="" />
                    {categorias.map((c) => (
                      <Picker.Item key={c.nombre} label={`${c.icono || '📋'} ${c.nombre}`} value={c.nombre} />
                    ))}
                  </Picker>
                </View>
              )}
              <PrimaryButton title="Ver análisis de valor" onPress={crearIntencion} style={{ marginTop: spacing.md }} />
            </UICard>

            <UICard style={{ marginBottom: spacing.md }}>
              <Text style={typography.label}>Regla de las 48 horas</Text>
              <Text style={[typography.small, { color: colors.textSecondary, lineHeight: 20, marginBottom: spacing.sm }]}>
                Si el precio iguala o supera el umbral, bloqueamos la confirmación durante 48 h. Valor actual:{' '}
                {formatearNumero(umbral)} {moneda}.
              </Text>
              <Text style={styles.lab}>Umbral ({moneda})</Text>
              <TextInput
                style={styles.input}
                value={umbralTexto}
                onChangeText={setUmbralTexto}
                keyboardType="decimal-pad"
                placeholder="50"
                placeholderTextColor={colors.textFaint}
              />
              <GhostButton title="Guardar umbral" onPress={guardarUmbral} style={{ marginTop: spacing.sm }} />
            </UICard>

            {intenciones.filter((x) => x.estado === 'pendiente').length > 0 ? (
              <UICard style={{ marginBottom: spacing.md }}>
                <Text style={typography.label}>Abrir intenciones recientes</Text>
                {intenciones
                  .filter((x) => x.estado === 'pendiente')
                  .map((i) => (
                    <TouchableOpacity
                      key={i.id}
                      style={styles.pickRow}
                      onPress={() => {
                        setEditId(i.id);
                      }}
                      activeOpacity={0.72}
                    >
                      <Text style={styles.pickTitle}>{i.nombre}</Text>
                      <Text style={styles.pickSub}>
                        {formatearNumero(i.precioEstimado)} · {i.nombreCategoria}
                      </Text>
                      <Ionicons name="analytics-outline" size={20} color={colors.chartBlue} />
                    </TouchableOpacity>
                  ))}
              </UICard>
            ) : null}

            {editando ? (
              <UICard accent style={{ marginBottom: spacing.md }}>
                <Text style={typography.title}>Análisis de valor</Text>
                <Text style={[typography.small, { color: colors.textMuted, marginBottom: spacing.md }]}>
                  {editando.nombre} · {formatearNumero(editando.precioEstimado)} {moneda}
                </Text>

                <Text style={typography.label}>A · Costo por uso</Text>
                <Text style={[typography.small, { marginBottom: spacing.sm }]}>
                  ¿Cuántas veces a la semana lo usarías y durante cuánto tiempo por sesión? Ajusta para ver el costo por uso.
                </Text>
                <Text style={styles.lab}>Veces por semana</Text>
                <TextInput
                  style={styles.input}
                  keyboardType="decimal-pad"
                  value={String(editando.vecesPorSemana)}
                  onChangeText={(t) => actualizarIntencion(editando.id, { vecesPorSemana: parseFloat(t) || 1 })}
                  placeholderTextColor={colors.textFaint}
                />
                <Text style={styles.lab}>Minutos por sesión</Text>
                <TextInput
                  style={styles.input}
                  keyboardType="number-pad"
                  value={String(editando.minutosPorSesion)}
                  onChangeText={(t) =>
                    actualizarIntencion(editando.id, { minutosPorSesion: parseFloat(t) || 30 })
                  }
                  placeholderTextColor={colors.textFaint}
                />
                <Text style={styles.lab}>Años de uso estimados</Text>
                <TextInput
                  style={styles.input}
                  keyboardType="decimal-pad"
                  value={String(editando.añosUso)}
                  onChangeText={(t) => actualizarIntencion(editando.id, { añosUso: parseFloat(t) || 1 })}
                  placeholderTextColor={colors.textFaint}
                />

                <View style={[styles.msgBox, { marginTop: spacing.md }]}>
                  <Ionicons name="bulb-outline" size={20} color={colors.accentGold} style={{ marginRight: spacing.sm }} />
                  <View style={{ flex: 1 }}>
                    <Text style={[typography.body, { lineHeight: 22 }]}>{analisis.msgValor}</Text>
                    <Text style={[typography.small, { color: colors.textMuted, marginTop: spacing.sm, lineHeight: 18 }]}>
                      Comparación referencia (~{formatearNumero(PRECIO_REF_CINE_DEFAULT)} {moneda}/sesión cine); sesiones imaginadas de ~
                      {Math.round(editando.minutosPorSesion)} min, {editando.vecesPorSemana}× semana durante{' '}
                      {editando.añosUso} año(s).
                    </Text>
                  </View>
                </View>

                {analisis.termo && analisis.termo.hayLimite ? (
                  <View style={{ marginTop: spacing.lg }}>
                    <Text style={typography.label}>B · Tu categoría este mes</Text>
                    <Text style={[typography.small, { color: colors.textMuted, marginBottom: spacing.sm }]}>
                      {analisis.termo.etiquetaLimite} · Gastado{' '}
                      {formatearNumero(analisis.termo.gastado)} + propuesta{' '}
                      {formatearNumero(analisis.termo.propuesto)}
                    </Text>
                    <View style={styles.thermoTrack}>
                      <View style={styles.thermoRow}>
                        <View style={[styles.thermoSegGastado, { width: `${Math.min(100, analisis.termo.barraPct)}%` }]} />
                        <View
                          style={[
                            styles.thermoSegProp,
                            {
                              width: `${Math.min(
                                100 - analisis.termo.barraPct,
                                Math.max(0, analisis.termo.sombraHastaPct - analisis.termo.barraPct)
                              )}%`,
                            },
                          ]}
                        />
                      </View>
                    </View>
                    {analisis.termo.alertaTexto ? (
                      <Text style={[typography.small, { marginTop: spacing.sm, color: colors.warning, lineHeight: 20 }]}>
                        {analisis.termo.alertaTexto}
                      </Text>
                    ) : null}
                  </View>
                ) : (
                  <Text style={[typography.small, { marginTop: spacing.md, color: colors.textMuted }]}>
                    Define límite en la categoría o un presupuesto mensual para ver la barra comparativa.
                  </Text>
                )}

                <View style={{ marginTop: spacing.lg }}>
                  <Text style={typography.label}>C · Cápsula del tiempo (48 h)</Text>
                  {editando.aplicabaCooldown ? (
                    <>
                      <Text style={[typography.small, { lineHeight: 20 }]}>
                        El monto está por encima del umbral. Estamos enfriando el impulso: si en dos días aún lo
                        quieres, podrás confirmar.
                      </Text>
                      <View style={styles.cronoBox}>
                        <Text style={styles.cronoText}>Faltan {formatCountdown(analisis.restanteMs)}</Text>
                        <Text style={[typography.small, { color: colors.textMuted, marginTop: spacing.xs }]}>
                          para poder confirmar desde aquí esta compra
                        </Text>
                      </View>
                    </>
                  ) : (
                    <Text style={[typography.small, { lineHeight: 20 }]}>
                      Por debajo del umbral: no aplicamos cuenta regresiva. Puedes registrar la compra cuando quieras.
                    </Text>
                  )}
                </View>

                <PrimaryButton
                  title={
                    analisis.bloqueo48 ? 'Registrar compra (esperando 48 h)' : 'Registrar compra (añadir al historial)'
                  }
                  disabled={analisis.bloqueo48}
                  onPress={() => registrarGastoDesdeIntencion(editando, null)}
                  style={{ marginTop: spacing.lg }}
                />
                <TouchableOpacity onPress={() => setEditId(null)} style={{ marginTop: spacing.md }}>
                  <Text style={{ color: colors.textMuted, textAlign: 'center', fontWeight: '600' }}>Cerrar análisis</Text>
                </TouchableOpacity>
              </UICard>
            ) : null}
          </>
        )}

        {tab === 'deseos' && (
          <UICard style={{ marginBottom: spacing.md }}>
            <Text style={typography.label}>Lista de deseos · validados por tiempo</Text>
            <Text style={[typography.small, { color: colors.textMuted, marginBottom: spacing.md, lineHeight: 20 }]}>
              Cuando la regla de 48 h termina (o no aplicaba), decide si compraste o libraste ese gasto.
            </Text>
            {intenciones.filter((i) => i.estado === 'pendiente').length === 0 ? (
              <Text style={typography.body}>No tienes compras pendientes. Crea una intención en la primera pestaña.</Text>
            ) : (
              intenciones
                .filter((i) => i.estado === 'pendiente')
                .map((i) => {
                  const puedeBtn = puedeRegistrarCompraPorRegla48h(i, ahora);
                  const esperaMs =
                    i.aplicabaCooldown && i.cooldownHasta != null ? Math.max(0, i.cooldownHasta - ahora) : 0;
                  return (
                    <View key={i.id} style={styles.tarjDeseo}>
                      <Text style={styles.tarjTit}>{i.nombre}</Text>
                      <Text style={styles.tarjSub}>
                        {formatearNumero(i.precioEstimado)} {moneda} · {i.nombreCategoria}
                      </Text>
                      {!puedeBtn && i.aplicabaCooldown ? (
                        <Text style={[typography.small, { color: colors.warning, marginTop: spacing.sm }]}>
                          ⏳ {formatCountdown(esperaMs)}
                        </Text>
                      ) : null}
                      <View style={{ flexDirection: 'row', marginTop: spacing.md }}>
                        <View style={{ flex: 1, marginRight: spacing.xs }}>
                          <PrimaryButton
                            disabled={!puedeBtn}
                            title="¡Lo compré!"
                            onPress={() => {
                              if (puedeBtn) registrarGastoDesdeIntencion(i, null);
                            }}
                          />
                        </View>
                        <View style={{ flex: 1, marginLeft: spacing.xs }}>
                          <GhostButton title="Ya no lo quiero" onPress={() => yaNoLoQuiero(i)} />
                        </View>
                      </View>
                    </View>
                  );
                })
            )}
          </UICard>
        )}

        {tab === 'super' && (
          <UICard style={{ marginBottom: spacing.md }}>
            <Text style={typography.label}>Checklist rápido · cosas que faltan</Text>
            <Text style={[typography.small, { color: colors.textMuted, marginBottom: spacing.md, lineHeight: 20 }]}>
              Marca lo que llevarás al súper o al mercado; estimamos según tus gastos pasados en la categoría elegida.
            </Text>

            <Text style={styles.lab}>Categoría de referencia (mercado / despensa)</Text>
            {categorias.length === 0 ? (
              <Text style={typography.small}>Crea al menos una categoría tipo despensa.</Text>
            ) : (
              <View style={styles.pickerWrap}>
                <Picker
                  selectedValue={listaSuperCat}
                  onValueChange={setListaSuperCat}
                  style={{ color: colors.text }}
                  dropdownIconColor={colors.text}
                >
                  {categorias.map((c) => (
                    <Picker.Item key={c.nombre} label={`${c.icono || '📋'} ${c.nombre}`} value={c.nombre} />
                  ))}
                </Picker>
              </View>
            )}
            <GhostButton title="Guardar como categoría preferida" onPress={guardarPreferidaSuper} style={{ marginTop: spacing.sm }} />

            {lineasSuper.length > 0 ? (
              <View style={{ marginTop: spacing.md, marginBottom: spacing.sm }}>
                <Text style={[typography.small, { color: colors.textMuted, marginBottom: spacing.sm }]}>
                  Pendiente por comprar ({lineasSuper.length})
                </Text>
                {ordenarLineasListaSuper(lineasSuper).map((ln) => (
                  <View key={ln.id} style={styles.superPendRow}>
                    <View style={{ flex: 1, minWidth: 0 }}>
                      <Text style={{ color: colors.text, fontWeight: '600', fontSize: 15 }}>{ln.nombre}</Text>
                      <Text style={[typography.small, { color: colorUrgenciaEtiqueta(ln.urgencia), marginTop: 2 }]}>
                        {etiquetaUrgencia(ln.urgencia)}
                      </Text>
                    </View>
                    <TouchableOpacity
                      onPress={() => marcarCompradoSuper(ln.id)}
                      style={styles.superHechoBtn}
                      hitSlop={{ top: 8, bottom: 8, left: 8, right: 8 }}
                    >
                      <Ionicons name="checkmark-circle" size={26} color={colors.mint} />
                    </TouchableOpacity>
                  </View>
                ))}
              </View>
            ) : null}

            <Text style={[styles.lab, { marginTop: spacing.lg }]}>Añadir a la lista y urgencia</Text>
            <Text style={[typography.small, { color: colors.textMuted, marginBottom: spacing.md, lineHeight: 20 }]}>
              Marca lo que falta. Para cada ítem elige: urgente, normal o puede esperar.
            </Text>
            <View style={styles.articulosListaBlock}>
              {articulosLista.map((nombre) => {
                const line = lineasSuper.find((l) => normNombre(l.nombre) === normNombre(nombre));
                const activo = !!line;
                return (
                  <View key={nombre} style={styles.superArtCard}>
                    <TouchableOpacity
                      style={styles.superArtRow}
                      onPress={() => toggleArticuloEnLista(nombre)}
                      activeOpacity={0.72}
                    >
                      <Ionicons
                        name={activo ? 'checkbox' : 'square-outline'}
                        size={22}
                        color={activo ? colors.mint : colors.textFaint}
                      />
                      <Text style={[styles.superArtNombre, activo && styles.superArtNombreOn]}>{nombre}</Text>
                    </TouchableOpacity>
                    {activo && line ? (
                      <View style={styles.urgenciaRow}>
                        {URGENCIA_LISTA_SUPER.map((u) => (
                          <TouchableOpacity
                            key={u.id}
                            style={[
                              styles.urgenciaChip,
                              line.urgencia === u.id && styles.urgenciaChipOn,
                              line.urgencia === u.id && u.id === 'urgente' && styles.urgenciaChipUrge,
                              line.urgencia === u.id && u.id === 'puede_esperar' && styles.urgenciaChipEspera,
                            ]}
                            onPress={() => setUrgenciaLinea(line.id, u.id)}
                          >
                            <Text
                              style={[
                                styles.urgenciaChipTxt,
                                line.urgencia === u.id && styles.urgenciaChipTxtOn,
                              ]}
                            >
                              {u.label}
                            </Text>
                          </TouchableOpacity>
                        ))}
                      </View>
                    ) : null}
                  </View>
                );
              })}
            </View>

            <View style={{ flexDirection: 'row', alignItems: 'center', marginTop: spacing.md }}>
              <TextInput
                style={[styles.input, { flex: 1, marginBottom: 0, marginRight: spacing.sm }]}
                placeholder="Otros (ej. sal)"
                placeholderTextColor={colors.textFaint}
                value={articuloNuevo}
                onChangeText={setArticuloNuevo}
              />
              <PrimaryButton title="Añadir" onPress={agregarArticuloExtra} />
            </View>

            {estimacionSuper ? (
              <View style={[styles.msgBox, { marginTop: spacing.md }]}>
                <Ionicons name="cart-outline" size={22} color={colors.chartBlue} style={{ marginRight: spacing.sm }} />
                <Text style={[typography.body, { flex: 1, lineHeight: 22 }]}>{estimacionSuper.mensaje}</Text>
              </View>
            ) : (
              <Text style={[typography.small, { marginTop: spacing.md, color: colors.textMuted }]}>
                Marca artículos arriba para ver una estimación orientativa.
              </Text>
            )}
          </UICard>
        )}
      </ScrollView>
    </ScreenWrap>
  );
}

const styles = StyleSheet.create({
  tabRow: { flexDirection: 'row', marginBottom: spacing.md },
  tabBtn: {
    flex: 1,
    marginHorizontal: spacing.xs,
    paddingVertical: spacing.sm,
    borderRadius: radii.md,
    borderWidth: 1,
    borderColor: colors.stroke,
    alignItems: 'center',
    backgroundColor: colors.surfaceSolid,
  },
  tabBtnActive: { borderColor: colors.chartBlue, backgroundColor: 'rgba(167, 216, 222, 0.08)' },
  tabLbl: { ...typography.small, fontWeight: '600', color: colors.textMuted },
  tabLblActive: { color: colors.text },
  lab: { ...typography.label, marginTop: spacing.md, marginBottom: spacing.xs },
  input: {
    borderWidth: 1,
    borderColor: colors.stroke,
    borderRadius: radii.md,
    paddingHorizontal: spacing.md,
    paddingVertical: Platform.OS === 'ios' ? spacing.md : spacing.sm,
    color: colors.text,
    marginBottom: 0,
  },
  pickerWrap: {
    borderWidth: 1,
    borderColor: colors.stroke,
    borderRadius: radii.md,
    overflow: 'hidden',
    backgroundColor: 'rgba(0,0,0,0.15)',
    marginBottom: spacing.xs,
  },
  pickRow: {
    flexDirection: 'row',
    alignItems: 'center',
    paddingVertical: spacing.sm,
    borderBottomWidth: 1,
    borderBottomColor: colors.stroke,
  },
  pickTitle: { flex: 1, color: colors.text, fontWeight: '700' },
  pickSub: { flex: 2, color: colors.textMuted, fontSize: 13 },
  msgBox: {
    flexDirection: 'row',
    alignItems: 'flex-start',
    backgroundColor: 'rgba(217, 180, 74, 0.08)',
    borderRadius: radii.md,
    padding: spacing.md,
    borderWidth: 1,
    borderColor: colors.stroke,
  },
  thermoTrack: {
    height: 14,
    borderRadius: 7,
    backgroundColor: colors.barTrack,
    overflow: 'hidden',
    width: '100%',
  },
  thermoRow: { flexDirection: 'row', height: 14, width: '100%' },
  thermoSegGastado: { height: 14, backgroundColor: colors.chartBlue },
  thermoSegProp: { height: 14, backgroundColor: 'rgba(232, 121, 249, 0.55)' },
  cronoBox: {
    marginTop: spacing.md,
    padding: spacing.md,
    borderRadius: radii.md,
    borderWidth: 1,
    borderColor: colors.stroke,
    backgroundColor: colors.surfaceHighlight,
  },
  cronoText: { fontSize: 22, fontWeight: '800', color: colors.accentBright, letterSpacing: 1 },
  tarjDeseo: {
    marginBottom: spacing.lg,
    paddingBottom: spacing.lg,
    borderBottomWidth: 1,
    borderBottomColor: colors.stroke,
  },
  tarjTit: { fontSize: 17, fontWeight: '700', color: colors.text },
  tarjSub: { fontSize: 13, color: colors.textMuted, marginTop: 4 },
  articulosListaBlock: { marginTop: spacing.xs },
  superArtCard: {
    marginBottom: spacing.sm,
    borderRadius: radii.md,
    borderWidth: 1,
    borderColor: colors.stroke,
    overflow: 'hidden',
    backgroundColor: 'rgba(0,0,0,0.12)',
  },
  superArtRow: {
    flexDirection: 'row',
    alignItems: 'center',
    paddingHorizontal: spacing.md,
    paddingVertical: spacing.sm,
  },
  superArtNombre: { marginLeft: spacing.sm, flex: 1, color: colors.textSecondary, fontSize: 15 },
  superArtNombreOn: { color: colors.text, fontWeight: '600' },
  urgenciaRow: {
    flexDirection: 'row',
    flexWrap: 'wrap',
    paddingHorizontal: spacing.sm,
    paddingBottom: spacing.sm,
    borderTopWidth: 1,
    borderTopColor: colors.stroke,
  },
  urgenciaChip: {
    paddingVertical: spacing.xs,
    paddingHorizontal: spacing.sm,
    marginRight: spacing.xs,
    marginTop: spacing.xs,
    borderRadius: radii.md,
    borderWidth: 1,
    borderColor: colors.stroke,
    backgroundColor: colors.surfaceSolid,
  },
  urgenciaChipOn: { borderWidth: 1 },
  urgenciaChipUrge: {
    borderColor: 'rgba(199, 120, 136, 0.6)',
    backgroundColor: 'rgba(199, 120, 136, 0.12)',
  },
  urgenciaChipEspera: {
    borderColor: 'rgba(155, 148, 184, 0.5)',
    backgroundColor: 'rgba(155, 148, 184, 0.1)',
  },
  urgenciaChipTxt: { fontSize: 12, fontWeight: '600', color: colors.textMuted },
  urgenciaChipTxtOn: { color: colors.accentBright },
  superPendRow: {
    flexDirection: 'row',
    alignItems: 'center',
    paddingVertical: spacing.sm,
    borderBottomWidth: 1,
    borderBottomColor: colors.stroke,
  },
  superHechoBtn: { padding: spacing.xs },
});
