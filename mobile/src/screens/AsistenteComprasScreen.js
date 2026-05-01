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
import { LinearGradient } from 'expo-linear-gradient';
import { Picker } from '@react-native-picker/picker';
import { rootNavigationRef } from '../navigation/rootNavigationRef';
import ScreenWrap from '../components/ScreenWrap';
import UICard from '../components/UICard';
import { PrimaryButton, GhostButton } from '../components/Buttons';
import { useApp } from '../context/AppContext';
import { normalizarCategoria, formatearNumero } from '../lib/finance';
import {
  generarIdIntencionCompra,
  costoPorSesion,
  construirAnalisisMensajeIntencion,
  datosTermometroCategoria,
  estimarListaSuperDesdeHistorial,
  elegirCategoriaSuperPorDefecto,
  generarIdListaSuperLinea,
  ordenarLineasListaSuper,
  URGENCIA_LISTA_SUPER,
  puedeRegistrarCompraPorRegla48h,
  formatCountdownMs,
} from '../lib/asistenteComprasLogic';
import { registrarGastoDesdeIntencionConUi, yaNoLoQuieroIntencionConUi } from '../lib/intencionesCompraAcciones';
import { colors, spacing, radii, typography, shadows } from '../theme';

const TABS = [
  {
    id: 'intencion',
    label: 'Intención',
    icon: 'flash-outline',
    grad: ['rgba(167, 139, 250, 0.55)', 'rgba(91, 33, 182, 0.5)', 'rgba(30, 22, 40, 0.85)'],
    border: 'rgba(196, 181, 253, 0.45)',
    fg: '#f5f3ff',
    fgMuted: 'rgba(196, 181, 253, 0.75)',
  },
  {
    id: 'deseos',
    label: 'Deseos',
    icon: 'heart-outline',
    grad: ['rgba(251, 113, 133, 0.45)', 'rgba(192, 38, 211, 0.42)', 'rgba(30, 22, 40, 0.88)'],
    border: 'rgba(251, 182, 196, 0.42)',
    fg: '#fdf2f8',
    fgMuted: 'rgba(251, 182, 196, 0.72)',
  },
  {
    id: 'super',
    label: 'Súper / básicos',
    icon: 'basket-outline',
    grad: ['rgba(45, 212, 191, 0.5)', 'rgba(167, 216, 222, 0.38)', 'rgba(18, 14, 28, 0.9)'],
    border: 'rgba(125, 211, 192, 0.5)',
    fg: colors.text,
    fgMuted: 'rgba(125, 211, 192, 0.7)',
  },
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
    registrarGastoDesdeIntencionConUi({
      state,
      intencion,
      origenValor,
      replaceState,
      onRemoved: () => {
        if (editId === intencion.id) setEditId(null);
      },
    });
  }

  function yaNoLoQuiero(intencion) {
    yaNoLoQuieroIntencionConUi({
      state,
      intencion,
      moneda,
      replaceState,
      onRemoved: () => {
        if (editId === intencion.id) setEditId(null);
      },
    });
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
        historialConfiable: false,
        ticketsHistorial: 0,
      };
    const ah = new Date();
    const mes = ah.getMonth();
    const año = ah.getFullYear();
    const precio = editando.precioEstimado;
    const cs = costoPorSesion(precio, editando.vecesPorSemana, editando.añosUso);
    const { msgValor, historialConfiable, ticketsHistorial } = construirAnalisisMensajeIntencion(
      state,
      editando,
      cs
    );
    const termo = datosTermometroCategoria(state, editando.nombreCategoria, mes, año, precio);
    const puede = puedeRegistrarCompraPorRegla48h(editando, ahora);
    const bloqueo48 = editando.aplicabaCooldown && !puede;
    const restanteMs =
      editando.aplicabaCooldown && editando.cooldownHasta != null
        ? Math.max(0, editando.cooldownHasta - ahora)
        : 0;
    return { costoSes: cs, msgValor, termo, puede, bloqueo48, restanteMs, historialConfiable, ticketsHistorial };
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

      <View style={styles.tabSegmentOuter}>
        {TABS.map((x) => {
          const active = tab === x.id;
          return (
            <TouchableOpacity
              key={x.id}
              style={styles.tabSegTouch}
              onPress={() => setTab(x.id)}
              activeOpacity={0.88}
              accessibilityRole="tab"
              accessibilityState={{ selected: active }}
            >
              {active ? (
                <LinearGradient
                  colors={x.grad}
                  start={{ x: 0, y: 0 }}
                  end={{ x: 1, y: 1 }}
                  style={[styles.tabSegActive, { borderColor: x.border }]}
                >
                  <Ionicons name={x.icon} size={22} color={x.fg} style={styles.tabSegIcon} />
                  <Text style={[styles.tabSegLblOn, { color: x.fg }]} numberOfLines={2}>
                    {x.label}
                  </Text>
                </LinearGradient>
              ) : (
                <View style={styles.tabSegIdle}>
                  <Ionicons name={x.icon} size={20} color={x.fgMuted} style={styles.tabSegIcon} />
                  <Text style={[styles.tabSegLblOff, { color: colors.textFaint }]} numberOfLines={2}>
                    {x.label}
                  </Text>
                </View>
              )}
            </TouchableOpacity>
          );
        })}
      </View>

      <ScrollView
        style={styles.scrollMain}
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
                      {analisis.historialConfiable
                        ? `El primer párrafo usa tus ${analisis.ticketsHistorial} gasto(s) registrado(s) en «${editando.nombreCategoria}» (últimos meses). `
                        : analisis.ticketsHistorial > 0
                          ? `Aún faltan registros en «${editando.nombreCategoria}» para contrastar con tu historial con confianza. `
                          : `Sin bastantes gastos previos en «${editando.nombreCategoria}», el análisis se basa en el precio, el costo por uso y las sesiones que estimaste abajo. `}
                      Uso estimado: ~{Math.round(editando.minutosPorSesion)} min por sesión, {editando.vecesPorSemana}× semana,{' '}
                      {editando.añosUso} año(s) de vida útil.
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
                        <Text style={styles.cronoText}>Faltan {formatCountdownMs(analisis.restanteMs)}</Text>
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
                          ⏳ {formatCountdownMs(esperaMs)}
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
  scrollMain: { flex: 1, backgroundColor: 'transparent' },
  tabSegmentOuter: {
    flexDirection: 'row',
    marginBottom: spacing.lg,
    padding: 5,
    borderRadius: radii.xl,
    backgroundColor: 'rgba(0,0,0,0.22)',
    borderWidth: 1,
    borderColor: 'rgba(255,255,255,0.07)',
    gap: 6,
  },
  tabSegTouch: { flex: 1, minWidth: 0 },
  tabSegActive: {
    borderRadius: radii.lg,
    borderWidth: 1.5,
    paddingVertical: spacing.md,
    paddingHorizontal: spacing.xs,
    alignItems: 'center',
    justifyContent: 'center',
    minHeight: 76,
    overflow: 'hidden',
    ...shadows.soft,
  },
  tabSegIdle: {
    borderRadius: radii.lg,
    paddingVertical: spacing.md,
    paddingHorizontal: spacing.xs,
    alignItems: 'center',
    justifyContent: 'center',
    minHeight: 76,
    backgroundColor: 'rgba(255,255,255,0.03)',
    borderWidth: 1,
    borderColor: 'rgba(255,255,255,0.04)',
  },
  tabSegIcon: { marginBottom: 6 },
  tabSegLblOn: {
    fontSize: 11,
    fontWeight: '800',
    textAlign: 'center',
    letterSpacing: 0.2,
    lineHeight: 14,
  },
  tabSegLblOff: {
    fontSize: 11,
    fontWeight: '600',
    textAlign: 'center',
    letterSpacing: 0.15,
    lineHeight: 14,
  },
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
