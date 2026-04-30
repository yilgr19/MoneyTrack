import React, { useMemo, useState, useCallback } from 'react';
import { View, Text, ScrollView, StyleSheet, TouchableOpacity } from 'react-native';
import { LinearGradient } from 'expo-linear-gradient';
import { Ionicons } from '@expo/vector-icons';
import ScreenWrap from '../components/ScreenWrap';
import DonutChart from '../components/charts/DonutChart';
import UICard from '../components/UICard';
import { useApp } from '../context/AppContext';
import { formatearNumero, parseFechaHoraLocal } from '../lib/finance';
import { construirDatosInformeMensual, opcionesMesesInforme } from '../lib/informeMensual';
import { colors, spacing, radii, typography, TAB_BAR_SCROLL_PADDING } from '../theme';

function fmtFechaCorta(f) {
  if (f == null) return '—';
  const s = String(f).trim();
  if (!s) return '—';
  const d = parseFechaHoraLocal(s) || new Date(s.includes('T') ? s : `${s.slice(0, 10)}T12:00:00`);
  if (Number.isNaN(d.getTime())) return s.slice(0, 10);
  return d.toLocaleDateString('es', { day: 'numeric', month: 'short' });
}

function FilaExtracto({ label, value, valueColor, bold }) {
  return (
    <View style={styles.filaExt}>
      <Text style={[styles.filaExtLab, bold && styles.filaExtLabBold]} numberOfLines={2}>
        {label}
      </Text>
      <Text
        style={[styles.filaExtVal, bold && styles.filaExtValBold, valueColor ? { color: valueColor } : null]}
        numberOfLines={1}
      >
        {value}
      </Text>
    </View>
  );
}

function BarraHorizontal({ label, monto, max, color }) {
  const pct = max > 0 ? Math.min(100, (monto / max) * 100) : 0;
  return (
    <View style={styles.barWrap}>
      <View style={styles.barHead}>
        <Text style={styles.barLab} numberOfLines={1}>
          {label}
        </Text>
        <Text style={styles.barVal}>{formatearNumero(monto)}</Text>
      </View>
      <View style={styles.barTrack}>
        <View style={[styles.barFill, { width: `${pct}%`, backgroundColor: color }]} />
      </View>
    </View>
  );
}

export default function InformesMensualesScreen() {
  const { state } = useApp();
  const moneda = (state?.moneda && String(state.moneda).trim()) || '';
  const suf = moneda ? ` ${moneda}` : '';

  const mesesOpt = useMemo(() => opcionesMesesInforme(36), []);
  const [idxMes, setIdxMes] = useState(0);
  const ym = mesesOpt[idxMes]?.value || '';

  const datos = useMemo(() => construirDatosInformeMensual(state || {}, ym), [state, ym]);

  const mesAnterior = useCallback(() => {
    setIdxMes((i) => Math.min(i + 1, mesesOpt.length - 1));
  }, [mesesOpt.length]);

  const mesSiguiente = useCallback(() => {
    setIdxMes((i) => Math.max(i - 1, 0));
  }, []);

  const segmentosCategorias = useMemo(() => {
    if (!datos) return [];
    const rows = datos.categoriaRows;
    if (rows.length === 0) return [];
    const top = rows.slice(0, 5);
    const rest = rows.slice(5).reduce((s, r) => s + r.monto, 0);
    const out = top.map((r) => ({
      value: r.monto,
      color: r.color || colors.textMuted,
      label: `${r.icono ? `${r.icono} ` : ''}${r.nombre}`.trim(),
    }));
    if (rest > 0.01) {
      out.push({ value: rest, color: '#94a3b8', label: 'Otros' });
    }
    return out;
  }, [datos]);

  const segmentosFlujo = useMemo(() => {
    if (!datos) return [];
    const ing = datos.totalIngresos;
    const gas = datos.totalGastosPres;
    if (ing <= 0 && gas <= 0) return [];
    return [
      { value: Math.max(0, ing), color: colors.mint, label: 'Ingresos' },
      { value: Math.max(0, gas), color: colors.danger, label: 'Gastos del mes' },
    ];
  }, [datos]);

  const maxCuenta = useMemo(() => {
    if (!datos?.cuentaRows?.length) return 0;
    return Math.max(...datos.cuentaRows.map((r) => r.monto));
  }, [datos]);

  const coloresCuenta = {
    efectivo: '#facc15',
    banco: '#60a5fa',
    tarjetaCredito: '#c084fc',
    nequi: '#22d3ee',
    daviplata: '#fb923c',
    billeteras: '#a78bfa',
  };

  const centroDonutCat = useMemo(() => {
    if (!segmentosCategorias.length) return { line1: undefined, line2: undefined };
    return {
      line1: `${formatearNumero(datos.totalGastosPres)}${suf}`.trim(),
      line2: 'Total gastos (criterio mes)',
    };
  }, [segmentosCategorias, datos, suf]);

  const centroDonutFlujo = useMemo(() => {
    if (!segmentosFlujo.length || !datos) return { line1: undefined, line2: undefined };
    const ing = datos.totalIngresos;
    const gas = datos.totalGastosPres;
    const t = ing + gas;
    if (t <= 0) return { line1: undefined, line2: undefined };
    const pi = Math.round((ing / t) * 100);
    const pg = Math.round((gas / t) * 100);
    return {
      line1: `${pi}%  ·  ${pg}%`,
      line2: 'Ingresos · gastos',
    };
  }, [segmentosFlujo, datos]);

  const refDoc = ym.replace('-', '');

  return (
    <ScreenWrap
      contentStyle={{
        paddingTop: spacing.xs,
        paddingBottom: TAB_BAR_SCROLL_PADDING + spacing.xl,
      }}
    >
      <View style={styles.selectorMes}>
        <TouchableOpacity
          onPress={mesAnterior}
          disabled={idxMes >= mesesOpt.length - 1}
          style={[styles.selBtn, idxMes >= mesesOpt.length - 1 && styles.selBtnOff]}
          hitSlop={10}
        >
          <Ionicons name="chevron-back" size={22} color={colors.accentBright} />
        </TouchableOpacity>
        <View style={styles.selMid}>
          <Text style={typography.label}>Periodo del informe</Text>
          <Text style={styles.selTit}>{mesesOpt[idxMes]?.label || ym}</Text>
        </View>
        <TouchableOpacity
          onPress={mesSiguiente}
          disabled={idxMes <= 0}
          style={[styles.selBtn, idxMes <= 0 && styles.selBtnOff]}
          hitSlop={10}
        >
          <Ionicons name="chevron-forward" size={22} color={colors.accentBright} />
        </TouchableOpacity>
      </View>

      <LinearGradient
        colors={['rgba(45, 52, 72, 0.95)', 'rgba(22, 24, 34, 0.98)']}
        start={{ x: 0, y: 0 }}
        end={{ x: 1, y: 1 }}
        style={styles.extractoHeader}
      >
        <View style={styles.extBadge}>
          <Ionicons name="document-text-outline" size={18} color={colors.accentGold} />
          <Text style={styles.extBadgeTxt}>Informe mensual</Text>
        </View>
        <Text style={styles.extTit}>Resumen financiero del periodo</Text>
        <Text style={styles.extSub}>
          Ref. {refDoc} · Generado en la app · Cifras según tus registros
        </Text>
        <View style={styles.extDiv} />
        {datos ? (
          <>
            <FilaExtracto
              label="Total ingresos del mes"
              value={`+${formatearNumero(datos.totalIngresos)}${suf}`}
              valueColor={colors.mint}
            />
            <FilaExtracto
              label="Total gastos (imputados al mes)"
              value={`−${formatearNumero(datos.totalGastosPres)}${suf}`}
              valueColor={colors.danger}
            />
            <FilaExtracto
              label="Resultado (ingresos − gastos)"
              value={`${datos.resultado >= 0 ? '+' : '−'}${formatearNumero(Math.abs(datos.resultado))}${suf}`}
              valueColor={datos.resultado >= 0 ? colors.mint : colors.danger}
              bold
            />
            <FilaExtracto
              label="Movimientos registrados"
              value={`${datos.numIngresos} ingresos · ${datos.numGastosMes} gastos`}
            />
          </>
        ) : null}
      </LinearGradient>

      {datos && datos.tope > 0 ? (
        <UICard>
          <Text style={typography.label}>Presupuesto mensual</Text>
          <Text style={[typography.title, { marginTop: spacing.xs, fontSize: 18 }]}>Tope y uso</Text>
          <FilaExtracto label="Tope configurado" value={`${formatearNumero(datos.tope)}${suf}`} />
          <FilaExtracto
            label="Gasto contado para presupuesto"
            value={`${formatearNumero(datos.totalGastosPres)}${suf}`}
          />
          <FilaExtracto
            label={datos.disponiblePres >= 0 ? 'Disponible bajo el tope' : 'Exceso sobre el tope'}
            value={`${datos.disponiblePres >= 0 ? '' : '−'}${formatearNumero(Math.abs(datos.disponiblePres))}${suf}`}
            valueColor={datos.disponiblePres >= 0 ? colors.mint : colors.danger}
            bold
          />
        </UICard>
      ) : null}

      {datos && datos.totalAportesMetas > 0 ? (
        <UICard>
          <Text style={typography.label}>Metas</Text>
          <Text style={[typography.title, { marginTop: spacing.xs, fontSize: 18 }]}>Aportes del mes</Text>
          <FilaExtracto
            label="Total aportes"
            value={`${formatearNumero(datos.totalAportesMetas)}${suf}`}
            valueColor={colors.accentGold}
          />
          {datos.aportesLines.slice(0, 6).map((a, i) => (
            <FilaExtracto key={`${a.nombre}-${i}`} label={a.nombre} value={`${formatearNumero(a.cant)}${suf}`} />
          ))}
        </UICard>
      ) : null}

      <UICard>
        <Text style={typography.label}>Gráficos</Text>
        <Text style={[typography.title, { marginTop: spacing.xs, fontSize: 18 }]}>Distribución</Text>
        {segmentosCategorias.length > 0 ? (
          <DonutChart
            segments={segmentosCategorias}
            title="Gastos por categoría"
            size={168}
            centerLine1={centroDonutCat.line1}
            centerLine2={centroDonutCat.line2}
            emptyHint="Sin gastos en el mes"
          />
        ) : (
          <Text style={[typography.small, { marginTop: spacing.sm }]}>Sin gastos imputados a este mes.</Text>
        )}
        {segmentosFlujo.length > 0 ? (
          <DonutChart
            segments={segmentosFlujo}
            title="Ingresos vs gastos"
            size={168}
            centerLine1={centroDonutFlujo.line1}
            centerLine2={centroDonutFlujo.line2}
            emptyHint="Sin movimientos"
          />
        ) : null}
      </UICard>

      {datos && datos.cuentaRows.length > 0 ? (
        <UICard>
          <Text style={typography.label}>Por cuenta de origen</Text>
          <Text style={[typography.title, { marginTop: spacing.xs, fontSize: 18 }]}>Gasto imputado al mes</Text>
          <Text style={[typography.small, { marginBottom: spacing.md }]}>
            Suma de cargos según la cuenta desde la que salió el dinero.
          </Text>
          {datos.cuentaRows.map((r) => (
            <BarraHorizontal
              key={r.id}
              label={r.label}
              monto={r.monto}
              max={maxCuenta}
              color={coloresCuenta[r.id] || colors.chartBlue}
            />
          ))}
        </UICard>
      ) : null}

      {datos && datos.categoriaRows.length > 0 ? (
        <UICard>
          <Text style={typography.label}>Detalle</Text>
          <Text style={[typography.title, { marginTop: spacing.xs, fontSize: 18 }]}>Categorías</Text>
          <ScrollView horizontal showsHorizontalScrollIndicator={false} style={{ marginTop: spacing.sm }}>
            <View>
              <View style={styles.tabHead}>
                <Text style={[styles.tabCel, styles.tabCelCat]}>Categoría</Text>
                <Text style={[styles.tabCel, styles.tabCelNum]}>Monto</Text>
                <Text style={[styles.tabCel, styles.tabCelPct]}>%</Text>
              </View>
              {datos.categoriaRows.map((r) => {
                const pct =
                  datos.totalGastosPres > 0 ? Math.round((r.monto / datos.totalGastosPres) * 1000) / 10 : 0;
                return (
                  <View key={r.nombre} style={styles.tabRow}>
                    <Text style={[styles.tabCel, styles.tabCelCat]} numberOfLines={1}>
                      {r.icono ? `${r.icono} ` : ''}
                      {r.nombre}
                    </Text>
                    <Text style={[styles.tabCel, styles.tabCelNum]}>{formatearNumero(r.monto)}</Text>
                    <Text style={[styles.tabCel, styles.tabCelPct]}>{pct}%</Text>
                  </View>
                );
              })}
            </View>
          </ScrollView>
        </UICard>
      ) : null}

      {datos && datos.topGastos.length > 0 ? (
        <UICard>
          <Text style={typography.label}>Movimientos</Text>
          <Text style={[typography.title, { marginTop: spacing.xs, fontSize: 18 }]}>Mayores gastos del mes</Text>
          <Text style={[typography.small, { marginBottom: spacing.sm }]}>
            Monto imputado a este mes (incluye cuotas de tarjeta cuando aplica).
          </Text>
          {datos.topGastos.map(({ g, m }, i) => (
            <View key={`${g.nombre}-${i}-${m}`} style={styles.movRow}>
              <View style={{ flex: 1, minWidth: 0 }}>
                <Text style={styles.movTit} numberOfLines={1}>
                  {g.nombre || 'Gasto'}
                </Text>
                <Text style={styles.movSub} numberOfLines={1}>
                  {fmtFechaCorta(g.fecha)}
                  {g.categoria ? ` · ${g.categoria}` : ''}
                </Text>
              </View>
              <Text style={styles.movMonto}>
                −{formatearNumero(m)}
                {suf}
              </Text>
            </View>
          ))}
        </UICard>
      ) : null}

      <Text style={[typography.small, { textAlign: 'center', marginTop: spacing.md, opacity: 0.85 }]}>
        Los totales usan las mismas reglas que Inicio y presupuesto (cuotas de tarjeta, fechas de corte).
      </Text>
    </ScreenWrap>
  );
}

const styles = StyleSheet.create({
  selectorMes: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    marginBottom: spacing.md,
    backgroundColor: colors.surface,
    borderRadius: radii.lg,
    paddingVertical: spacing.sm,
    paddingHorizontal: spacing.sm,
    borderWidth: 1,
    borderColor: colors.stroke,
  },
  selBtn: { padding: spacing.sm },
  selBtnOff: { opacity: 0.35 },
  selMid: { flex: 1, alignItems: 'center' },
  selTit: { ...typography.title, fontSize: 17, marginTop: 4 },
  extractoHeader: {
    borderRadius: radii.xl,
    padding: spacing.lg,
    marginBottom: spacing.md,
    borderWidth: 1,
    borderColor: colors.stroke,
  },
  extBadge: {
    flexDirection: 'row',
    alignItems: 'center',
    alignSelf: 'flex-start',
    backgroundColor: 'rgba(217, 180, 74, 0.12)',
    paddingHorizontal: 10,
    paddingVertical: 5,
    borderRadius: radii.pill,
    gap: 6,
  },
  extBadgeTxt: { color: colors.accentGold, fontSize: 11, fontWeight: '800', letterSpacing: 0.8 },
  extTit: {
    color: colors.text,
    fontSize: 20,
    fontWeight: '800',
    marginTop: spacing.md,
    letterSpacing: -0.3,
  },
  extSub: { color: colors.textMuted, fontSize: 13, marginTop: 6, lineHeight: 18 },
  extDiv: {
    height: 1,
    backgroundColor: 'rgba(199, 195, 227, 0.2)',
    marginVertical: spacing.md,
  },
  filaExt: {
    flexDirection: 'row',
    justifyContent: 'space-between',
    alignItems: 'flex-start',
    marginTop: spacing.sm,
    gap: spacing.md,
  },
  filaExtLab: { flex: 1, color: colors.textSecondary, fontSize: 14, lineHeight: 20 },
  filaExtLabBold: { fontWeight: '800', color: colors.text },
  filaExtVal: { fontSize: 15, fontWeight: '700', color: colors.text, fontVariant: ['tabular-nums'] },
  filaExtValBold: { fontSize: 17 },
  barWrap: { marginBottom: spacing.md },
  barHead: { flexDirection: 'row', justifyContent: 'space-between', alignItems: 'center' },
  barLab: { flex: 1, color: colors.textSecondary, fontSize: 13, marginRight: spacing.sm },
  barVal: { fontSize: 14, fontWeight: '700', color: colors.text, fontVariant: ['tabular-nums'] },
  barTrack: {
    height: 8,
    backgroundColor: 'rgba(125, 193, 145, 0.12)',
    borderRadius: 4,
    marginTop: 6,
    overflow: 'hidden',
  },
  barFill: { height: '100%', borderRadius: 4 },
  tabHead: { flexDirection: 'row', borderBottomWidth: 1, borderBottomColor: colors.stroke, paddingBottom: 8 },
  tabRow: { flexDirection: 'row', paddingVertical: 10, borderBottomWidth: 1, borderBottomColor: colors.stroke },
  tabCel: { fontSize: 13, color: colors.textSecondary },
  tabCelCat: { width: 200, color: colors.text, fontWeight: '600' },
  tabCelNum: { width: 100, textAlign: 'right', fontVariant: ['tabular-nums'] },
  tabCelPct: { width: 48, textAlign: 'right', fontWeight: '700' },
  movRow: {
    flexDirection: 'row',
    alignItems: 'center',
    paddingVertical: spacing.sm,
    borderBottomWidth: 1,
    borderBottomColor: colors.stroke,
    gap: spacing.sm,
  },
  movTit: { fontSize: 15, fontWeight: '700', color: colors.text },
  movSub: { fontSize: 12, color: colors.textMuted, marginTop: 2 },
  movMonto: { fontSize: 15, fontWeight: '800', color: colors.danger, fontVariant: ['tabular-nums'] },
});
