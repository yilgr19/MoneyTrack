import React, { useMemo } from 'react';
import { View, Text } from 'react-native';
import ScreenWrap from '../components/ScreenWrap';
import UICard from '../components/UICard';
import { useApp } from '../context/AppContext';
import { formatearNumero, obtenerMesAño, montoGastoAfectaSaldo } from '../lib/finance';
import { colors, spacing, typography, layoutStyles } from '../theme';

const NOMBRES_MES = [
  'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
  'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre',
];

export default function ReportesScreen() {
  const { state } = useApp();
  const moneda = state.moneda || '';

  const meses = useMemo(() => {
    const ingresos = state.ingresos || [];
    const gastos = state.gastos || [];
    const contrib = state.contribucionesMetas || [];
    const set = new Set();
    const add = (f) => {
      if (!f) return;
      const { mes, año } = obtenerMesAño(f);
      set.add(JSON.stringify({ mes, año }));
    };
    ingresos.forEach((i) => add(i.fecha));
    gastos.forEach((g) => add(g.fecha));
    contrib.forEach((c) => add(c.fecha));
    return Array.from(set)
      .map((s) => JSON.parse(s))
      .sort((a, b) => (b.año !== a.año ? b.año - a.año : b.mes - a.mes));
  }, [state.ingresos, state.gastos, state.contribucionesMetas]);

  return (
    <ScreenWrap includeTopInset={false} contentStyle={{ paddingTop: spacing.xs }}>
      <Text style={typography.label}>Histórico</Text>
      <Text style={typography.hero}>Reportes</Text>
      <Text style={[typography.subtitle, { marginBottom: spacing.lg }]}>Totales por mes</Text>

      {meses.length === 0 ? (
        <UICard style={{ marginBottom: 0 }}>
          <Text style={typography.small}>Aún no hay datos por mes.</Text>
        </UICard>
      ) : (
        meses.map(({ mes, año }) => {
          const ingMes = (state.ingresos || [])
            .filter((i) => {
              const k = obtenerMesAño(i.fecha);
              return k.mes === mes && k.año === año;
            })
            .reduce((s, i) => s + i.cantidad, 0);
          const gasMes = (state.gastos || [])
            .filter((g) => {
              const k = obtenerMesAño(g.fecha);
              return k.mes === mes && k.año === año;
            })
            .reduce((s, g) => s + montoGastoAfectaSaldo(g), 0);
          const apMes = (state.contribucionesMetas || [])
            .filter((c) => {
              const k = obtenerMesAño(c.fecha);
              return k.mes === mes && k.año === año;
            })
            .reduce((s, c) => s + c.cantidad, 0);
          return (
            <UICard key={`${año}-${mes}`}>
              <Text style={typography.label}>{NOMBRES_MES[mes]} {año}</Text>
              <View style={layoutStyles.statRow}>
                <Text style={[typography.small, layoutStyles.statLabel]}>Ingresos</Text>
                <Text style={[typography.monoAmount, layoutStyles.statValue, { color: colors.mint }]}>
                  {formatearNumero(ingMes)} {moneda}
                </Text>
              </View>
              <View style={layoutStyles.statRow}>
                <Text style={[typography.small, layoutStyles.statLabel]}>Gastos</Text>
                <Text style={[typography.monoAmount, layoutStyles.statValue, { color: colors.danger }]}>
                  {formatearNumero(gasMes)} {moneda}
                </Text>
              </View>
              <View style={layoutStyles.statRow}>
                <Text style={[typography.small, layoutStyles.statLabel]}>Metas</Text>
                <Text style={[typography.monoAmount, layoutStyles.statValue]}>
                  {formatearNumero(apMes)} {moneda}
                </Text>
              </View>
            </UICard>
          );
        })
      )}
    </ScreenWrap>
  );
}
