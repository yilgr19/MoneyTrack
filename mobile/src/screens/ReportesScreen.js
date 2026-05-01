import React, { useCallback, useMemo } from 'react';
import { View, Text, StyleSheet, SectionList, TouchableOpacity, Alert } from 'react-native';
import { Ionicons } from '@expo/vector-icons';
import ScreenWrap from '../components/ScreenWrap';
import { useApp } from '../context/AppContext';
import {
  formatearNumero,
  obtenerCuentasDestinoIngreso,
  normalizarOrigenCuenta,
  reemplazarPagosRecordatorioTarjetas,
} from '../lib/finance';
import { withAvisoGastoMovimiento } from '../lib/notificacionesApp';
import { rootNavigationRef } from '../navigation/rootNavigationRef';
import { colors, spacing, radii, typography } from '../theme';

function tsDeFecha(f) {
  if (f == null) return 0;
  const s = String(f).trim();
  if (!s) return 0;
  const d = new Date(s.includes('T') ? s : `${s.slice(0, 10)}T12:00:00`);
  return Number.isNaN(d.getTime()) ? 0 : d.getTime();
}

function formatearFechaMov(f) {
  if (f == null) return '—';
  const s = String(f).trim();
  if (!s) return '—';
  const d = new Date(s.includes('T') ? s : `${s.slice(0, 10)}T12:00:00`);
  if (Number.isNaN(d.getTime())) return s;
  return d.toLocaleDateString('es', { day: 'numeric', month: 'short', year: 'numeric' });
}

function cortarLabelCuenta(label) {
  if (!label) return '';
  const i = label.indexOf(' (');
  return i > 0 ? label.slice(0, i).trim() : label;
}

const NOMBRES_MES = [
  'Enero',
  'Febrero',
  'Marzo',
  'Abril',
  'Mayo',
  'Junio',
  'Julio',
  'Agosto',
  'Septiembre',
  'Octubre',
  'Noviembre',
  'Diciembre',
];

/** Clave YYYY-MM para ordenar meses; null si no hay fecha útil */
function claveMesDesdeTs(ts) {
  if (ts == null || ts <= 0) return null;
  const d = new Date(ts);
  if (Number.isNaN(d.getTime())) return null;
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
}

/** Todas las claves YYYY-MM entre min y max (inclusive), orden descendente (más reciente primero). */
function mesesCalendarioEntreDesc(minKey, maxKey) {
  const [ya, ma] = minKey.split('-').map((x) => parseInt(x, 10, 10));
  const [yb, mb] = maxKey.split('-').map((x) => parseInt(x, 10, 10));
  if (!Number.isFinite(ya) || !Number.isFinite(ma) || !Number.isFinite(yb) || !Number.isFinite(mb)) return [];
  const out = [];
  let y = ya;
  let m = ma;
  for (;;) {
    out.push(`${y}-${String(m).padStart(2, '0')}`);
    if (y === yb && m === mb) break;
    m += 1;
    if (m > 12) {
      m = 1;
      y += 1;
    }
  }
  return out.sort((a, b) => b.localeCompare(a));
}

function filaPlaceholderMesVacio(keyMes) {
  return {
    id: `mes-sin-movs-${keyMes}`,
    kind: 'vacioMes',
    ts: 0,
    titulo: 'Sin movimientos este mes',
    sub: '',
    monto: 0,
    detalle: null,
  };
}

export default function ReportesScreen() {
  const { state, replaceState } = useApp();
  const moneda = (state.moneda && String(state.moneda).trim()) || '';

  const navegarEditarGasto = useCallback((gastoId) => {
    const id = String(gastoId || '').trim();
    if (!id) return;
    if (rootNavigationRef.isReady()) {
      rootNavigationRef.navigate('Gastos', { editarGastoId: id });
    }
  }, []);

  const confirmarEliminarGasto = useCallback(
    (gastoId, titulo) => {
      const id = String(gastoId || '').trim();
      if (!id) return;
      const t = String(titulo || 'Gasto').trim().slice(0, 72);
      Alert.alert('¿Quitar este gasto?', `"${t || 'Este registro'}" saldrá del historial.`, [
        { text: 'No', style: 'cancel' },
        {
          text: 'Quitar',
          style: 'destructive',
          onPress: () =>
            replaceState((s) => {
              const g = (s.gastos || []).find((x) => x && String(x.id) === id);
              const nombre = g ? String(g.nombre || '').trim() || t : t;
              const mon = (s.moneda && String(s.moneda).trim()) || '';
              const montoLine = g
                ? `${formatearNumero(Math.abs(parseFloat(g.cantidad) || 0))} ${mon}`.trim()
                : '';
              const gastos = (s.gastos || []).filter((x) => !x || String(x.id) !== id);
              let st = { ...s, gastos };
              st = withAvisoGastoMovimiento(st, 'eliminado', { nombre, montoLine });
              return {
                ...st,
                pagosProgramados: reemplazarPagosRecordatorioTarjetas(st.pagosProgramados, st, new Date()),
              };
            }),
        },
      ]);
    },
    [replaceState]
  );

  const { labelCuenta } = useMemo(() => {
    const cuentas = obtenerCuentasDestinoIngreso(state || {});
    const m = new Map();
    cuentas.forEach((c) => m.set(c.value, cortarLabelCuenta(c.label)));
    const rowTc = cuentas.find((x) => x.value === 'tarjetaCredito');
    const fn = (origen) => {
      if (origen == null || origen === '') return '—';
      const o = String(origen);
      if (m.has(o)) return m.get(o);
      if (normalizarOrigenCuenta(o) === 'tarjetaCredito') {
        return rowTc ? cortarLabelCuenta(rowTc.label) : 'Tarjeta de crédito';
      }
      return o;
    };
    return { labelCuenta: fn };
  }, [state]);

  const seccionesMovs = useMemo(() => {
    const out = [];
    const metas = state.metas || [];
    (state.ingresos || []).forEach((i, idx) => {
      const t = tsDeFecha(i.fecha);
      const tit = (i.nota && String(i.nota).trim()) || 'Ingreso';
      out.push({
        id: `ing-${idx}`,
        kind: 'ingreso',
        ts: t,
        titulo: tit,
        sub: [formatearFechaMov(i.fecha), labelCuenta(i.origen)].filter(Boolean).join(' · '),
        monto: Math.abs(parseFloat(i.cantidad) || 0),
        detalle: null,
      });
    });
    (state.gastos || []).forEach((g, idx) => {
      const t = tsDeFecha(g.fecha);
      const orig = labelCuenta(g.origen);
      const esTc = normalizarOrigenCuenta(g.origen) === 'tarjetaCredito';
      const nCuotas = esTc && (g.cuotas || 1) > 1 ? g.cuotas : null;
      const partes = [g.categoria, orig, nCuotas ? `${nCuotas} cuotas` : null].filter(Boolean);
      const gid = g?.id != null && String(g.id).trim() ? String(g.id) : null;
      out.push({
        id: gid ? `gas-${gid}` : `gas-idx-${idx}`,
        kind: 'gasto',
        gastoId: gid,
        esTransferenciaBolsillo: !!g.esTransferenciaBolsillo,
        ts: t,
        titulo: (g.nombre && String(g.nombre).trim()) || 'Gasto',
        sub: [formatearFechaMov(g.fecha), partes.length ? partes.join(' · ') : null].filter(Boolean).join(' · '),
        monto: Math.abs(parseFloat(g.cantidad) || 0),
        detalle: esTc ? 'Tarjeta' : null,
      });
    });
    (state.contribucionesMetas || []).forEach((c, idx) => {
      const t = tsDeFecha(c.fecha);
      const nomMeta = metas.find((m) => m.id === c.metaId)?.nombre || 'Meta';
      out.push({
        id: `meta-${idx}`,
        kind: 'meta',
        ts: t,
        titulo: `Aporte: ${nomMeta}`,
        sub: [formatearFechaMov(c.fecha), labelCuenta(c.origen)].filter(Boolean).join(' · '),
        monto: Math.abs(parseFloat(c.cantidad) || 0),
        detalle: null,
      });
    });
    out.sort((a, b) => {
      if (b.ts !== a.ts) return b.ts - a.ts;
      return String(b.id).localeCompare(String(a.id));
    });

    const porMes = new Map();
    const sinFecha = [];
    for (const row of out) {
      const k = claveMesDesdeTs(row.ts);
      if (k == null) {
        sinFecha.push(row);
        continue;
      }
      if (!porMes.has(k)) porMes.set(k, []);
      porMes.get(k).push(row);
    }

    const clavesConDatos = Array.from(porMes.keys());
    const clavesOrdenadas =
      clavesConDatos.length > 0
        ? (() => {
            const asc = [...clavesConDatos].sort((a, b) => a.localeCompare(b));
            const minK = asc[0];
            const maxK = asc[asc.length - 1];
            return mesesCalendarioEntreDesc(minK, maxK);
          })()
        : [];
    const secciones = clavesOrdenadas.map((k) => {
      const [ys, ms] = k.split('-');
      const y = parseInt(ys, 10);
      const m = parseInt(ms, 10) - 1;
      const titulo = Number.isFinite(y) && m >= 0 && m <= 11 ? `${NOMBRES_MES[m]} ${y}` : k;
      const dataMes = porMes.get(k);
      const data = dataMes && dataMes.length > 0 ? dataMes : [filaPlaceholderMesVacio(k)];
      return { title: titulo, keyMes: k, data };
    });
    if (sinFecha.length > 0) {
      secciones.push({ title: 'Sin fecha clara', keyMes: '_sin', data: sinFecha });
    }
    return secciones;
  }, [state, labelCuenta]);

  const prefijoMonto = (kind) => {
    if (kind === 'ingreso') return '+';
    if (kind === 'vacioMes') return '';
    return '−';
  };

  const colorMonto = (kind) => {
    if (kind === 'ingreso') return colors.mint;
    if (kind === 'gasto') return colors.danger;
    if (kind === 'vacioMes') return colors.textMuted;
    return colors.accentGold;
  };

  const iconoFila = (kind) => {
    if (kind === 'ingreso') return 'trending-up';
    if (kind === 'gasto') return 'trending-down';
    if (kind === 'vacioMes') return 'calendar-outline';
    return 'flag';
  };

  return (
    <ScreenWrap includeTopInset={false} scrollEnabled={false} contentStyle={{ paddingTop: spacing.xs }}>
      <Text style={typography.label}>Historial</Text>
      <Text style={typography.hero}>Movimientos</Text>
      <Text style={[typography.subtitle, { marginBottom: spacing.lg }]}>
        Ingresos, gastos (incl. tarjeta) y aportes a metas, por mes calendario entre tu primer y último movimiento
        (más reciente arriba; los meses sin registros muestran un aviso)
      </Text>
      {seccionesMovs.length === 0 ? (
        <View style={styles.vacio}>
          <Text style={typography.small}>Aún no hay movimientos registrados.</Text>
        </View>
      ) : (
        <SectionList
          sections={seccionesMovs}
          keyExtractor={(it) => it.id}
          style={styles.listaMovs}
          contentContainerStyle={{ paddingBottom: spacing.xl }}
          stickySectionHeadersEnabled
          renderSectionHeader={({ section }) => (
            <View style={styles.sectionHeader}>
              <Text style={styles.sectionHeaderText}>{section.title}</Text>
            </View>
          )}
          renderItem={({ item }) => {
            const puedeGestionarGasto =
              item.kind === 'gasto' && item.gastoId && !item.esTransferenciaBolsillo;
            const esVacioMes = item.kind === 'vacioMes';
            return (
              <View style={[styles.row, esVacioMes && styles.rowVacioMes]}>
                <View style={styles.iconCircle}>
                  <Ionicons name={iconoFila(item.kind)} size={22} color={colorMonto(item.kind)} />
                </View>
                <View style={styles.rowText}>
                  <View style={styles.rowTop}>
                    <Text style={styles.badge} numberOfLines={1}>
                      {item.kind === 'ingreso'
                        ? 'Ingreso'
                        : item.kind === 'gasto'
                          ? 'Gasto'
                          : item.kind === 'vacioMes'
                            ? 'Mes'
                            : 'Meta'}
                    </Text>
                    {item.detalle ? (
                      <Text style={styles.badgeTarj} numberOfLines={1}>
                        {item.detalle}
                      </Text>
                    ) : null}
                  </View>
                  <Text style={styles.tit} numberOfLines={2}>
                    {item.titulo}
                  </Text>
                  <Text style={styles.sub} numberOfLines={2}>
                    {item.sub}
                  </Text>
                </View>
                <View style={styles.rowTail}>
                  {puedeGestionarGasto ? (
                    <View style={styles.rowMiniActs}>
                      <TouchableOpacity
                        onPress={() => navegarEditarGasto(item.gastoId)}
                        hitSlop={{ top: 10, bottom: 10, left: 8, right: 8 }}
                        accessibilityLabel="Editar gasto"
                      >
                        <Ionicons name="create-outline" size={17} color={colors.textMuted} />
                      </TouchableOpacity>
                      <TouchableOpacity
                        onPress={() => confirmarEliminarGasto(item.gastoId, item.titulo)}
                        hitSlop={{ top: 10, bottom: 10, left: 8, right: 8 }}
                        accessibilityLabel="Quitar gasto"
                      >
                        <Ionicons name="close-circle-outline" size={17} color={colors.textMuted} />
                      </TouchableOpacity>
                    </View>
                  ) : null}
                  {esVacioMes ? null : (
                    <Text style={[typography.monoAmount, styles.monto, { color: colorMonto(item.kind) }]}>
                      {prefijoMonto(item.kind)} {formatearNumero(item.monto)} {moneda}
                    </Text>
                  )}
                </View>
              </View>
            );
          }}
        />
      )}
    </ScreenWrap>
  );
}

const styles = StyleSheet.create({
  listaMovs: { flex: 1, minHeight: 0, backgroundColor: 'transparent' },
  sectionHeader: {
    backgroundColor: colors.bg,
    paddingTop: spacing.sm,
    paddingBottom: spacing.xs,
    marginBottom: 2,
    borderBottomWidth: 1,
    borderBottomColor: colors.stroke,
  },
  sectionHeaderText: {
    ...typography.label,
    color: colors.textSecondary,
    fontSize: 13,
    letterSpacing: 0.8,
  },
  vacio: {
    padding: spacing.lg,
    backgroundColor: colors.surface,
    borderRadius: radii.lg,
    borderWidth: 1,
    borderColor: colors.stroke,
  },
  rowVacioMes: { opacity: 0.9 },
  row: {
    flexDirection: 'row',
    alignItems: 'flex-start',
    backgroundColor: colors.surface,
    borderRadius: radii.lg,
    padding: spacing.md,
    marginBottom: spacing.sm,
    borderWidth: 1,
    borderColor: colors.stroke,
  },
  iconCircle: {
    width: 44,
    height: 44,
    borderRadius: radii.md,
    backgroundColor: colors.surfaceHighlight,
    alignItems: 'center',
    justifyContent: 'center',
    marginRight: spacing.md,
    borderWidth: 1,
    borderColor: colors.stroke,
    flexShrink: 0,
  },
  rowText: { flex: 1, minWidth: 0 },
  rowTop: { flexDirection: 'row', alignItems: 'center', flexWrap: 'wrap', marginBottom: 4 },
  badge: {
    fontSize: 11,
    fontWeight: '700',
    color: colors.textMuted,
    textTransform: 'uppercase',
    letterSpacing: 0.6,
    marginRight: 6,
  },
  badgeTarj: {
    fontSize: 11,
    fontWeight: '700',
    color: colors.chartBlue,
    textTransform: 'uppercase',
    letterSpacing: 0.5,
    marginLeft: 2,
  },
  tit: { color: colors.text, fontSize: 16, fontWeight: '700', letterSpacing: -0.2 },
  sub: { color: colors.textMuted, fontSize: 13, marginTop: 4, lineHeight: 18 },
  rowTail: {
    marginLeft: spacing.sm,
    flexShrink: 0,
    alignItems: 'flex-end',
    justifyContent: 'center',
    gap: 4,
  },
  rowMiniActs: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 4,
    opacity: 0.92,
  },
  monto: { textAlign: 'right' },
});
