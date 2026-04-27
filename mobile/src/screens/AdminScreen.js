import React from 'react';
import { View, Text, StyleSheet, TouchableOpacity, Alert } from 'react-native';
import ScreenWrap from '../components/ScreenWrap';
import UICard from '../components/UICard';
import { useApp } from '../context/AppContext';
import { formatearNumero, calcularSaldosPorCuenta, montoGastoAfectaSaldo } from '../lib/finance';
import { colors, spacing, radii, typography } from '../theme';

export default function AdminScreen() {
  const { state, resetPartial, resetFull } = useApp();
  const moneda = state.moneda || '';
  const saldos = calcularSaldosPorCuenta(state);
  const totalGastos = (state.gastos || []).reduce((s, g) => s + montoGastoAfectaSaldo(g), 0);

  return (
    <ScreenWrap includeTopInset={false} contentStyle={{ paddingTop: spacing.xs }}>
      <Text style={typography.label}>Sistema</Text>
      <Text style={typography.hero}>Administrar</Text>
      <Text style={[typography.subtitle, { marginBottom: spacing.lg }]}>Resumen y reseteos</Text>

      <UICard>
        <Text style={typography.label}>Estado</Text>
        <View style={styles.statBlock}>
          <Text style={typography.small}>Saldo total calculado</Text>
          <Text
            style={styles.bigNum}
            adjustsFontSizeToFit
            minimumFontScale={0.65}
            numberOfLines={2}
            maxFontSizeMultiplier={1.2}
          >
            {formatearNumero(saldos.total)} {moneda}
          </Text>
        </View>
        <View style={styles.statBlock}>
          <Text style={typography.small}>Gastos acumulados</Text>
          <Text style={typography.monoAmount}>
            {formatearNumero(totalGastos)} {moneda}
          </Text>
        </View>
        <Text style={[typography.small, { marginTop: spacing.md, lineHeight: 20 }]}>
          Excel (importar/exportar) está en la versión web. Aquí solo resets locales.
        </Text>
      </UICard>

      <TouchableOpacity
        style={[styles.btn, styles.btnWarn]}
        activeOpacity={0.88}
        onPress={() => {
          Alert.alert(
            'Resetear saldo y movimientos',
            'Se pondrán en cero saldos iniciales, ingresos, gastos, metas y presupuesto. ¿Continuar?',
            [
              { text: 'Cancelar', style: 'cancel' },
              { text: 'Resetear', style: 'destructive', onPress: () => resetPartial() },
            ]
          );
        }}
      >
        <Text style={styles.btnText}>Resetear saldo y gastos</Text>
      </TouchableOpacity>

      <TouchableOpacity
        style={[styles.btn, styles.btnDanger]}
        activeOpacity={0.88}
        onPress={() => {
          Alert.alert(
            'Resetear todo',
            'Se borrará también moneda, categorías y pagos programados. ¿Continuar?',
            [
              { text: 'Cancelar', style: 'cancel' },
              { text: 'Borrar todo', style: 'destructive', onPress: () => resetFull() },
            ]
          );
        }}
      >
        <Text style={styles.btnText}>Resetear todo el proyecto</Text>
      </TouchableOpacity>
    </ScreenWrap>
  );
}

const styles = StyleSheet.create({
  statBlock: { marginTop: spacing.md },
  bigNum: {
    fontSize: 24,
    fontWeight: '700',
    color: colors.mint,
    marginTop: 4,
    letterSpacing: -0.5,
    maxWidth: '100%',
  },
  btn: {
    paddingVertical: spacing.md,
    paddingHorizontal: spacing.lg,
    borderRadius: radii.md,
    alignItems: 'center',
    marginBottom: spacing.md,
    borderWidth: 1,
  },
  btnWarn: {
    backgroundColor: 'rgba(180, 83, 9, 0.35)',
    borderColor: 'rgba(251, 191, 36, 0.4)',
  },
  btnDanger: {
    backgroundColor: 'rgba(153, 27, 27, 0.4)',
    borderColor: 'rgba(248, 113, 113, 0.35)',
    marginBottom: 0,
  },
  btnText: { color: '#fff', fontWeight: '700', fontSize: 15 },
});
