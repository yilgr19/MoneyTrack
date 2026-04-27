import React, { useMemo } from 'react';
import { View, Text, StyleSheet, FlatList } from 'react-native';
import { Ionicons } from '@expo/vector-icons';
import ScreenWrap from '../components/ScreenWrap';
import { useApp } from '../context/AppContext';
import { formatearNumero, obtenerCuentasDestinoIngreso, normalizarOrigenCuenta } from '../lib/finance';
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

export default function ReportesScreen() {
  const { state } = useApp();
  const moneda = (state.moneda && String(state.moneda).trim()) || '';

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

  const filas = useMemo(() => {
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
      out.push({
        id: `gas-${idx}`,
        kind: 'gasto',
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
    out.sort((a, b) => b.ts - a.ts);
    return out;
  }, [state, labelCuenta]);

  const prefijoMonto = (kind) => (kind === 'ingreso' ? '+' : '−');

  const colorMonto = (kind) => {
    if (kind === 'ingreso') return colors.mint;
    if (kind === 'gasto') return colors.danger;
    return colors.accentGold;
  };

  const iconoFila = (kind) => {
    if (kind === 'ingreso') return 'trending-up';
    if (kind === 'gasto') return 'trending-down';
    return 'flag';
  };

  return (
    <ScreenWrap includeTopInset={false} contentStyle={{ paddingTop: spacing.xs }}>
      <Text style={typography.label}>Historial</Text>
      <Text style={typography.hero}>Movimientos</Text>
      <Text style={[typography.subtitle, { marginBottom: spacing.lg }]}>
        Ingresos, gastos (incl. tarjeta) y aportes a metas, por fecha
      </Text>
      {filas.length === 0 ? (
        <View style={styles.vacio}>
          <Text style={typography.small}>Aún no hay movimientos registrados.</Text>
        </View>
      ) : (
        <FlatList
          data={filas}
          keyExtractor={(it) => it.id}
          contentContainerStyle={{ paddingBottom: spacing.xl }}
          renderItem={({ item }) => (
            <View style={styles.row}>
              <View style={styles.iconCircle}>
                <Ionicons name={iconoFila(item.kind)} size={22} color={colorMonto(item.kind)} />
              </View>
              <View style={styles.rowText}>
                <View style={styles.rowTop}>
                  <Text style={styles.badge} numberOfLines={1}>
                    {item.kind === 'ingreso' ? 'Ingreso' : item.kind === 'gasto' ? 'Gasto' : 'Meta'}
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
              <Text style={[typography.monoAmount, styles.monto, { color: colorMonto(item.kind) }]}>
                {prefijoMonto(item.kind)} {formatearNumero(item.monto)} {moneda}
              </Text>
            </View>
          )}
        />
      )}
    </ScreenWrap>
  );
}

const styles = StyleSheet.create({
  vacio: {
    padding: spacing.lg,
    backgroundColor: colors.surface,
    borderRadius: radii.lg,
    borderWidth: 1,
    borderColor: colors.stroke,
  },
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
  monto: { marginLeft: spacing.sm, flexShrink: 0, textAlign: 'right' },
});
