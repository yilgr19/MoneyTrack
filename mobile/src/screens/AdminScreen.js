import React, { useCallback, useState } from 'react';
import {
  View,
  Text,
  StyleSheet,
  TouchableOpacity,
  Alert,
  ActivityIndicator,
  Linking,
  Platform,
} from 'react-native';
import ScreenWrap from '../components/ScreenWrap';
import UICard from '../components/UICard';
import { useApp } from '../context/AppContext';
import { formatearNumero, calcularSaldosPorCuenta, montoGastoAfectaSaldo } from '../lib/finance';
import {
  programarNotificacionLocalDePrueba,
  contarNotificacionesLocalesProgramadas,
  diagnosticarNotificacionesLocales,
  sincronizarNotificacionesLocalesPagosProgramados,
  sincronizarNotificacionesLocalesTarjetasCredito,
} from '../lib/notificacionesLocalesPagosProgramados';
import { notificacionesSistemaDisponibles } from '../lib/notificacionesLocalesEntorno';
import { colors, spacing, radii, typography } from '../theme';

export default function AdminScreen() {
  const { state, resetPartial, resetFull, exportarDatosRespaldo, importarDatosRespaldo } = useApp();
  const moneda = state.moneda || '';
  const saldos = calcularSaldosPorCuenta(state);
  const totalGastos = (state.gastos || []).reduce((s, g) => s + montoGastoAfectaSaldo(g), 0);
  const [pruebaNotifCargando, setPruebaNotifCargando] = useState(false);
  const [diagNotif, setDiagNotif] = useState('');
  const [diagNotifCargando, setDiagNotifCargando] = useState(false);
  const [exportCargando, setExportCargando] = useState(false);
  const [importCargando, setImportCargando] = useState(false);

  const probarNotificacionLocal = useCallback(async () => {
    setPruebaNotifCargando(true);
    try {
      const r = await programarNotificacionLocalDePrueba();
      const n = await contarNotificacionesLocalesProgramadas();
      Alert.alert(r.ok ? 'Prueba programada' : 'No se pudo programar', `${r.mensaje}${n >= 0 ? `\n\nPendientes en el sistema ahora: ${n}.` : ''}`);
    } finally {
      setPruebaNotifCargando(false);
    }
  }, []);

  const verPendientesNotif = useCallback(async () => {
    const n = await contarNotificacionesLocalesProgramadas();
    Alert.alert(
      'Notificaciones locales pendientes',
      n < 0
        ? 'No se pudo leer la lista (revisa que no sea web).'
        : `Hay ${n} aviso(s) programado(s) (pagos, tarjetas, pruebas…). Se disparan en sus fechas/intervalos.`
    );
  }, []);

  const actualizarDiagNotif = useCallback(async () => {
    setDiagNotifCargando(true);
    try {
      if (notificacionesSistemaDisponibles()) {
        try {
          await sincronizarNotificacionesLocalesPagosProgramados(state);
          await sincronizarNotificacionesLocalesTarjetasCredito(state);
        } catch (e) {
          if (typeof __DEV__ !== 'undefined' && __DEV__) console.warn('[MoneyTrack] reprogramar alarmas:', e);
        }
      }
      const d = await diagnosticarNotificacionesLocales(state);
      setDiagNotif(d.texto);
    } finally {
      setDiagNotifCargando(false);
    }
  }, [state]);

  const abrirAjustesApp = useCallback(() => {
    Linking.openSettings().catch(() => {
      Alert.alert('Ajustes', 'Abre manualmente los ajustes del teléfono y busca MoneyTrack.');
    });
  }, []);

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
      Alert.alert('Importar', 'Usa la app instalada para elegir un archivo .moneytrack.json.');
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
          Exporta todo lo guardado en el teléfono (meses o años de movimientos, categorías, tarjetas, pagos
          programados, metas, lista súper, etc.) en un archivo{' '}
          <Text style={{ fontWeight: '700' }}>.moneytrack.json</Text>. Al reinstalar o cambiar de móvil, usa
          Importar y elige ese archivo. El formato solo lo entiende MoneyTrack.
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

      <UICard style={{ marginBottom: spacing.md }}>
        <Text style={typography.label}>Notificaciones del sistema</Text>
        <Text style={[typography.small, { marginTop: spacing.sm, lineHeight: 20, color: colors.textSecondary }]}>
          Prueba rápida (~20 s): necesitas APK o dev build (no Expo Go). Cierra la app o ponla en segundo plano antes
          de que suene. Si nunca ves avisos, pulsa «Diagnóstico» tras instalar de nuevo esta versión (canales Android
          nuevos).
        </Text>
        <TouchableOpacity
          style={[styles.btn, styles.btnInfo, pruebaNotifCargando && { opacity: 0.75 }]}
          activeOpacity={0.88}
          disabled={pruebaNotifCargando}
          onPress={probarNotificacionLocal}
        >
          {pruebaNotifCargando ? (
            <ActivityIndicator color="#fff" />
          ) : (
            <Text style={styles.btnText}>Probar notificación en ~20 segundos</Text>
          )}
        </TouchableOpacity>
        <TouchableOpacity style={[styles.btn, styles.btnInfoMuted]} activeOpacity={0.88} onPress={verPendientesNotif}>
          <Text style={styles.btnTextMuted}>Ver cuántas hay programadas</Text>
        </TouchableOpacity>
        <TouchableOpacity
          style={[styles.btn, styles.btnInfoMuted, diagNotifCargando && { opacity: 0.75 }]}
          activeOpacity={0.88}
          disabled={diagNotifCargando}
          onPress={actualizarDiagNotif}
        >
          {diagNotifCargando ? (
            <ActivityIndicator color={colors.textSecondary} />
          ) : (
            <Text style={styles.btnTextMuted}>Diagnóstico y reprogramar alarmas</Text>
          )}
        </TouchableOpacity>
        {Platform.OS !== 'web' ? (
          <TouchableOpacity style={[styles.btn, styles.btnInfoMuted]} activeOpacity={0.88} onPress={abrirAjustesApp}>
            <Text style={styles.btnTextMuted}>Abrir ajustes de la app</Text>
          </TouchableOpacity>
        ) : null}
        {diagNotif ? (
          <Text
            style={[
              typography.small,
              { marginTop: spacing.sm, lineHeight: 20, color: colors.textMuted, fontFamily: 'monospace' },
            ]}
            selectable
          >
            {diagNotif}
          </Text>
        ) : null}
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
  btnInfo: {
    marginTop: spacing.md,
    backgroundColor: 'rgba(75, 36, 108, 0.85)',
    borderColor: 'rgba(199, 195, 227, 0.35)',
  },
  btnInfoMuted: {
    marginTop: spacing.sm,
    backgroundColor: 'rgba(32, 26, 44, 0.6)',
    borderColor: colors.stroke,
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
  btnTextMuted: { color: colors.textSecondary, fontWeight: '700', fontSize: 15 },
  btnTextImport: { color: colors.mint, fontWeight: '700', fontSize: 15 },
});
