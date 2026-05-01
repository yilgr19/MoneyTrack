import React, { useCallback, useState } from 'react';
import {
  View,
  Text,
  StyleSheet,
  TouchableOpacity,
  Alert,
  ActivityIndicator,
  Platform,
} from 'react-native';
import ScreenWrap from '../components/ScreenWrap';
import UICard from '../components/UICard';
import { useApp } from '../context/AppContext';
import { formatearNumero, calcularSaldosPorCuenta, montoGastoAfectaSaldo } from '../lib/finance';
import { colors, spacing, radii, typography } from '../theme';

export default function AdminScreen() {
  const { state, resetPartial, resetFull, exportarDatosRespaldo, importarDatosRespaldo } = useApp();
  const moneda = state.moneda || '';
  const saldos = calcularSaldosPorCuenta(state);
  const totalGastos = (state.gastos || []).reduce((s, g) => s + montoGastoAfectaSaldo(g), 0);
  const [exportCargando, setExportCargando] = useState(false);
  const [importCargando, setImportCargando] = useState(false);

  const exportarDatos = useCallback(async () => {
    if (Platform.OS === 'web') {
      Alert.alert('Exportar', 'Usa la app instalada en el teléfono para generar el archivo de respaldo.');
      return;
    }
    setExportCargando(true);
    try {
      const r = await exportarDatosRespaldo();
      Alert.alert(r.ok ? 'Exportar datos' : 'No se pudo exportar', r.mensaje || (r.ok ? 'Listo.' : 'Intenta de nuevo.'));
    } finally {
      setExportCargando(false);
    }
  }, [exportarDatosRespaldo]);

  const importarDatos = useCallback(() => {
    if (Platform.OS === 'web') {
      Alert.alert(
        'Importar',
        'Usa la app instalada para elegir el archivo de respaldo (.csv reciente o .json si lo exportaste antes).'
      );
      return;
    }
    Alert.alert(
      'Importar respaldo',
      'Se reemplazará toda la información local por la del archivo (gastos, ingresos, cuentas, tarjetas, metas, etc.). ¿Continuar?',
      [
        { text: 'Cancelar', style: 'cancel' },
        {
          text: 'Elegir archivo',
          onPress: async () => {
            setImportCargando(true);
            try {
              const r = await importarDatosRespaldo();
              if (r.cancelado) return;
              Alert.alert(r.ok ? 'Importación lista' : 'Error', r.mensaje || '');
            } finally {
              setImportCargando(false);
            }
          },
        },
      ]
    );
  }, [importarDatosRespaldo]);

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
          Excel masivo sigue en la versión web; aquí puedes respaldo completo en archivo propio de la app.
        </Text>
      </UICard>

      <UICard style={{ marginBottom: spacing.md }}>
        <Text style={typography.label}>Respaldo de datos (APK / instalación nueva)</Text>
        <Text style={[typography.small, { marginTop: spacing.sm, lineHeight: 20, color: colors.textSecondary }]}>
          <Text style={{ fontWeight: '700' }}>Exportar</Text> genera un archivo <Text style={{ fontWeight: '700' }}>.csv</Text>{' '}
          con toda tu información en un formato que la app entiende al importar. Guárdalo donde quieras (Descargas, Drive,
          correo…). <Text style={{ fontWeight: '700' }}>Importar</Text> restaura gastos, ingresos, cuentas, tarjetas,
          metas y el resto; también acepta respaldos <Text style={{ fontWeight: '700' }}>.json</Text> antiguos.
        </Text>
        <TouchableOpacity
          style={[styles.btn, styles.btnBackup, exportCargando && { opacity: 0.75 }]}
          activeOpacity={0.88}
          disabled={exportCargando}
          onPress={exportarDatos}
        >
          {exportCargando ? (
            <ActivityIndicator color="#fff" />
          ) : (
            <Text style={styles.btnText}>Exportar datos</Text>
          )}
        </TouchableOpacity>
        <TouchableOpacity
          style={[styles.btn, styles.btnBackupMuted, importCargando && { opacity: 0.75 }]}
          activeOpacity={0.88}
          disabled={importCargando}
          onPress={importarDatos}
        >
          {importCargando ? (
            <ActivityIndicator color={colors.mint} />
          ) : (
            <Text style={styles.btnTextImport}>Importar datos</Text>
          )}
        </TouchableOpacity>
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
  btnBackup: {
    marginTop: spacing.md,
    backgroundColor: 'rgba(34, 197, 94, 0.35)',
    borderColor: 'rgba(134, 239, 172, 0.45)',
  },
  btnBackupMuted: {
    marginTop: spacing.sm,
    backgroundColor: 'rgba(32, 26, 44, 0.75)',
    borderColor: 'rgba(134, 239, 172, 0.25)',
    marginBottom: 0,
  },
  btnText: { color: '#fff', fontWeight: '700', fontSize: 15 },
  btnTextImport: { color: colors.mint, fontWeight: '700', fontSize: 15 },
});
