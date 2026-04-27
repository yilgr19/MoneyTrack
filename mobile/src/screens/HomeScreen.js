import React, { useMemo } from 'react';
import { View, Text, StyleSheet, TextInput, TouchableOpacity, Alert, useWindowDimensions } from 'react-native';
import { LinearGradient } from 'expo-linear-gradient';
import ScreenWrap from '../components/ScreenWrap';
import UICard from '../components/UICard';
import { PrimaryButton } from '../components/Buttons';
import { useApp } from '../context/AppContext';
import {
  formatearNumero,
  CUENTAS,
  calcularSaldosPorCuenta,
  obtenerMesAño,
  montoGastoAfectaSaldo,
  verificarAlertaTarjetaCredito,
  normalizarCategoria,
} from '../lib/finance';
import { colors, spacing, radii, typography, layoutStyles } from '../theme';

export default function HomeScreen() {
  const { state, replaceState } = useApp();
  const moneda = state.moneda || '';

  const derived = useMemo(() => {
    const ahora = new Date();
    const mesActual = ahora.getMonth();
    const añoActual = ahora.getFullYear();
    const nombreMes = [
      'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
      'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre',
    ][mesActual];

    const saldosPorCuenta = calcularSaldosPorCuenta(state);
    const saldoActual = saldosPorCuenta.total || 0;
    const gastos = state.gastos || [];
    const ingresos = state.ingresos || [];
    const contribuciones = state.contribucionesMetas || [];

    const ingresosMesActual = ingresos
      .filter((i) => {
        const { mes, año } = obtenerMesAño(i.fecha);
        return mes === mesActual && año === añoActual;
      })
      .reduce((s, i) => s + i.cantidad, 0);

    const gastosMesActual = gastos
      .filter((g) => {
        const { mes, año } = obtenerMesAño(g.fecha);
        return mes === mesActual && año === añoActual;
      })
      .reduce((s, g) => s + montoGastoAfectaSaldo(g), 0);

    const totalGastos = gastos.reduce((s, g) => s + montoGastoAfectaSaldo(g), 0);
    const esUsuarioNuevo = saldoActual === 0 && ingresos.length === 0;
    const alertaTc = verificarAlertaTarjetaCredito(state);
    const presupuestoMensual = state.presupuestoMensual || 0;

    const gastosDelMes = gastos.filter((g) => {
      const { mes, año } = obtenerMesAño(g.fecha);
      return mes === mesActual && año === añoActual;
    });
    const mayorGasto =
      gastosDelMes.length > 0
        ? gastosDelMes.reduce((max, g) => (g.cantidad > max.cantidad ? g : max), gastosDelMes[0])
        : null;

    const categoriasData = (state.categorias || []).map(normalizarCategoria);
    const gastosMesPorCategoria = {};
    gastos
      .filter((g) => {
        const { mes, año } = obtenerMesAño(g.fecha);
        return mes === mesActual && año === añoActual;
      })
      .forEach((g) => {
        const cat = g.categoria || 'Otros';
        gastosMesPorCategoria[cat] = (gastosMesPorCategoria[cat] || 0) + g.cantidad;
      });
    const totalGastosMes = Object.values(gastosMesPorCategoria).reduce((s, v) => s + v, 0) || 1;

    const metasData = state.metas || [];
    const ultimosGastos = gastos.slice().reverse().slice(0, 8);

    let estadoMsg = '';
    let estadoDetalle = '';
    let estadoKind = 'info';
    if (presupuestoMensual <= 0) {
      estadoMsg = 'Define un presupuesto mensual para saber si vas bien o mal.';
    } else {
      const disponible = presupuestoMensual - gastosMesActual;
      const pctUsado = (gastosMesActual / presupuestoMensual) * 100;
      if (disponible > 0 && pctUsado < 80) {
        estadoMsg = '¡Vas bien!';
        estadoDetalle = `Te quedan ${formatearNumero(disponible)} ${moneda} disponibles.`;
        estadoKind = 'ok';
      } else if (disponible > 0 && pctUsado >= 80) {
        estadoMsg = 'Cuidado, te acercas al límite';
        estadoDetalle = `Te quedan ${formatearNumero(disponible)} ${moneda}.`;
        estadoKind = 'cuidado';
      } else if (disponible === 0) {
        estadoMsg = 'Has agotado tu presupuesto mensual';
        estadoDetalle = `Gastaste exactamente ${formatearNumero(presupuestoMensual)} ${moneda} este mes.`;
        estadoKind = 'alerta';
      } else {
        estadoMsg = 'Has superado tu presupuesto';
        estadoDetalle = `Te has pasado en ${formatearNumero(Math.abs(disponible))} ${moneda}.`;
        estadoKind = 'superado';
      }
    }

    return {
      saldosPorCuenta,
      saldoActual,
      gastos,
      ingresos,
      contribuciones,
      ingresosMesActual,
      gastosMesActual,
      totalGastos,
      esUsuarioNuevo,
      alertaTc,
      presupuestoMensual,
      mayorGasto,
      categoriasData,
      gastosMesPorCategoria,
      totalGastosMes,
      metasData,
      ultimosGastos,
      mesActual,
      añoActual,
      nombreMes,
      estadoMsg,
      estadoDetalle,
      estadoKind,
    };
  }, [state, moneda]);

  const pctPresupuesto =
    derived.presupuestoMensual > 0
      ? Math.min(100, (derived.gastosMesActual / derived.presupuestoMensual) * 100)
      : 0;

  function guardarPresupuesto(val) {
    const n = parseFloat(val) || 0;
    replaceState((s) => ({
      ...s,
      presupuestoMensual: n > 0 ? n : 0,
    }));
  }

  if (derived.esUsuarioNuevo) {
    return (
      <ScreenWrap contentStyle={{ paddingTop: spacing.sm }}>
        <Text style={typography.hero}>Bienvenido</Text>
        <Text style={[typography.subtitle, { marginTop: spacing.xs, marginBottom: spacing.lg }]}>
          Tu dinero, con claridad
        </Text>
        <UICard accent>
          <Text style={typography.label}>Primeros pasos</Text>
          <Step n={1} text="Saldo o ingresos (Saldo / Más → Ingresos)" />
          <Step n={2} text="Categorías en Más → Categorías" />
          <Step n={3} text="Registra gastos en la pestaña Gastos" />
        </UICard>
      </ScreenWrap>
    );
  }

  return (
    <ScreenWrap contentStyle={{ paddingTop: spacing.xs }}>
      <Text style={typography.label}>Resumen</Text>
      <Text style={typography.hero}>Inicio</Text>
      <Text style={[typography.subtitle, { marginBottom: spacing.lg }]}>
        {derived.nombreMes} {derived.añoActual}
      </Text>

      <LinearGradient
        colors={['rgba(124, 58, 237, 0.35)', 'rgba(8, 6, 14, 0.4)']}
        start={{ x: 0, y: 0 }}
        end={{ x: 1, y: 1 }}
        style={styles.heroCard}
      >
        <Text style={styles.heroLabel}>Patrimonio estimado</Text>
        <Text
          style={styles.heroAmount}
          adjustsFontSizeToFit
          minimumFontScale={0.55}
          maxFontSizeMultiplier={1.25}
          numberOfLines={2}
        >
          {formatearNumero(derived.saldoActual)} <Text style={styles.heroMoneda}>{moneda}</Text>
        </Text>
        <View style={styles.heroRow}>
          <Text style={styles.heroMeta}>
            +{formatearNumero(derived.ingresosMesActual)} {moneda} ingresos mes
          </Text>
          <Text style={styles.heroMetaDot}>·</Text>
          <Text style={styles.heroMeta}>
            −{formatearNumero(derived.gastosMesActual)} {moneda} gastos mes
          </Text>
        </View>
      </LinearGradient>

      {derived.alertaTc.mostrar && (
        <View style={styles.alerta}>
          <Text style={styles.alertaTit}>Tarjeta de crédito</Text>
          <Text style={styles.alertaBody}>
            {formatearNumero(derived.alertaTc.gastado)} {moneda} usados (
            {formatearNumero(derived.alertaTc.porcentaje, 1)}% del límite).
          </Text>
        </View>
      )}

      <UICard>
        <Text style={typography.label}>Por cuenta</Text>
        {CUENTAS.filter((c) => (derived.saldosPorCuenta[c.id] || 0) !== 0).map((c) => (
          <View key={c.id} style={layoutStyles.rowBetween}>
            <Text style={[typography.body, layoutStyles.rowLabel]}>{c.nombre}</Text>
            <Text style={[typography.monoAmount, layoutStyles.rowValue]}>
              {formatearNumero(derived.saldosPorCuenta[c.id])} {moneda}
            </Text>
          </View>
        ))}
        <Text style={[typography.small, { marginTop: spacing.sm }]}>
          {derived.nombreMes}: gastos {formatearNumero(derived.gastosMesActual)} {moneda}
          {derived.contribuciones.reduce((s, c) => s + c.cantidad, 0) > 0
            ? ` · Metas ${formatearNumero(derived.contribuciones.reduce((s, c) => s + c.cantidad, 0))} ${moneda}`
            : ''}
        </Text>
      </UICard>

      <UICard>
        <Text style={typography.label}>Presupuesto mensual</Text>
        {derived.presupuestoMensual > 0 ? (
          <>
            <View style={layoutStyles.rowBetween}>
              <Text style={[typography.body, layoutStyles.rowLabel]}>Planeado</Text>
              <Text style={[typography.monoAmount, layoutStyles.rowValue]}>
                {formatearNumero(derived.presupuestoMensual)} {moneda}
              </Text>
            </View>
            <View style={layoutStyles.rowBetween}>
              <Text style={[typography.body, layoutStyles.rowLabel]}>Disponible</Text>
              <Text style={[typography.monoAmount, layoutStyles.rowValue, { color: colors.mint }]}>
                {formatearNumero(Math.max(0, derived.presupuestoMensual - derived.gastosMesActual))}{' '}
                {moneda}
              </Text>
            </View>
            <View style={styles.barBg}>
              <LinearGradient
                colors={[colors.btnFrom, colors.accentDeep]}
                start={{ x: 0, y: 0 }}
                end={{ x: 1, y: 0 }}
                style={[styles.barFill, { width: `${pctPresupuesto}%` }]}
              />
            </View>
            <TouchableOpacity
              onPress={() => {
                Alert.alert('Presupuesto', '¿Eliminar presupuesto?', [
                  { text: 'Cancelar', style: 'cancel' },
                  {
                    text: 'Eliminar',
                    style: 'destructive',
                    onPress: () => replaceState((s) => ({ ...s, presupuestoMensual: 0 })),
                  },
                ]);
              }}
            >
              <Text style={styles.link}>Eliminar presupuesto</Text>
            </TouchableOpacity>
          </>
        ) : (
          <PresupuestoQuickInput onSave={guardarPresupuesto} />
        )}
      </UICard>

      <UICard>
        <Text style={typography.label}>Análisis</Text>
        {derived.mayorGasto ? (
          <Text style={typography.body}>
            Mayor gasto: <Text style={{ fontWeight: '700', color: colors.text }}>{derived.mayorGasto.nombre}</Text> —{' '}
            {formatearNumero(derived.mayorGasto.cantidad)} {moneda}
          </Text>
        ) : (
          <Text style={typography.small}>Sin gastos este mes.</Text>
        )}
        {derived.presupuestoMensual > 0 && (
          <Text style={[typography.small, { marginTop: spacing.xs }]}>
            {formatearNumero(derived.gastosMesActual)} / {formatearNumero(derived.presupuestoMensual)} {moneda}
          </Text>
        )}
        <Text
          style={[
            styles.estado,
            derived.estadoKind === 'ok' && styles.estadoOk,
            derived.estadoKind === 'cuidado' && styles.estadoCuidado,
            derived.estadoKind === 'alerta' && styles.estadoAlerta,
            derived.estadoKind === 'superado' && styles.estadoSuperado,
          ]}
        >
          {derived.estadoMsg}
        </Text>
        {!!derived.estadoDetalle && (
          <Text style={[typography.body, { marginTop: 4, flexShrink: 1 }]}>{derived.estadoDetalle}</Text>
        )}
      </UICard>

      <UICard>
        <Text style={typography.label}>Gastos por categoría</Text>
        {derived.categoriasData.length === 0 ? (
          <Text style={typography.small}>Crea categorías primero.</Text>
        ) : (
          derived.categoriasData
            .sort(
              (a, b) =>
                (derived.gastosMesPorCategoria[b.nombre] || 0) -
                (derived.gastosMesPorCategoria[a.nombre] || 0)
            )
            .slice(0, 6)
            .map((cat) => {
              const monto = derived.gastosMesPorCategoria[cat.nombre] || 0;
              const pct = derived.totalGastosMes > 0 ? (monto / derived.totalGastosMes) * 100 : 0;
              return (
                <View key={cat.nombre} style={{ marginBottom: spacing.md }}>
                  <View style={layoutStyles.rowBetween}>
                    <Text style={[typography.body, layoutStyles.rowLabel]}>
                      {cat.icono} {cat.nombre}
                    </Text>
                    <Text style={[typography.monoAmount, layoutStyles.rowValue]}>
                      {formatearNumero(monto)} {moneda}
                    </Text>
                  </View>
                  <View style={styles.barBg}>
                    <View
                      style={[
                        styles.barFillCat,
                        { width: `${Math.min(100, pct)}%`, backgroundColor: cat.color },
                      ]}
                    />
                  </View>
                </View>
              );
            })
        )}
      </UICard>

      <UICard>
        <Text style={typography.label}>Últimos movimientos</Text>
        {derived.gastos.length === 0 ? (
          <Text style={typography.small}>Aún no hay gastos.</Text>
        ) : (
          <>
            <Text style={[typography.small, { marginBottom: spacing.sm }]}>
              Histórico: {formatearNumero(derived.totalGastos)} {moneda}
            </Text>
            {derived.ultimosGastos.map((g, i) => (
              <View key={i} style={styles.moveRow}>
                <View style={styles.moveDot} />
                <Text style={styles.moveText}>
                  {g.nombre} · {formatearNumero(g.cantidad)} {moneda}{' '}
                  <Text style={{ color: colors.textFaint }}>({g.categoria})</Text>
                </Text>
              </View>
            ))}
          </>
        )}
      </UICard>

      <UICard style={{ marginBottom: 0 }}>
        <Text style={typography.label}>Metas</Text>
        {derived.metasData.length === 0 ? (
          <Text style={typography.small}>Sin metas aún.</Text>
        ) : (
          derived.metasData.slice(0, 4).map((meta) => {
            const acum = (state.contribucionesMetas || [])
              .filter((c) => c.metaId === meta.id)
              .reduce((s, c) => s + c.cantidad, 0);
            const obj = parseFloat(meta.objetivo) || 0;
            const pct = obj > 0 ? Math.min(100, (acum / obj) * 100) : 0;
            return (
              <View key={meta.id} style={{ marginBottom: spacing.md }}>
                <View style={layoutStyles.rowBetween}>
                  <Text style={[typography.body, layoutStyles.rowLabel]}>{meta.nombre}</Text>
                  <Text style={[typography.monoAmount, layoutStyles.rowValue]}>
                    {formatearNumero(acum)} / {formatearNumero(obj)}
                  </Text>
                </View>
                <View style={styles.barBg}>
                  <LinearGradient
                    colors={[colors.mint, colors.success]}
                    start={{ x: 0, y: 0 }}
                    end={{ x: 1, y: 0 }}
                    style={[styles.barFillCat, { width: `${pct}%` }]}
                  />
                </View>
              </View>
            );
          })
        )}
      </UICard>
    </ScreenWrap>
  );
}

function Step({ n, text }) {
  return (
    <View style={styles.stepRow}>
      <View style={styles.stepBadge}>
        <Text style={styles.stepNum}>{n}</Text>
      </View>
      <Text style={[typography.body, { flex: 1 }]}>{text}</Text>
    </View>
  );
}

function PresupuestoQuickInput({ onSave }) {
  const [v, setV] = React.useState('');
  const { width } = useWindowDimensions();
  const stackVertical = width < 368;
  return (
    <View style={[styles.presRow, stackVertical && styles.presRowStack]}>
      <TextInput
        style={[styles.input, stackVertical && styles.inputStacked]}
        value={v}
        onChangeText={setV}
        placeholder="Ej. 2000"
        placeholderTextColor={colors.textFaint}
        keyboardType="decimal-pad"
      />
      <PrimaryButton
        title="Guardar"
        style={stackVertical ? styles.presBtnStacked : { flexShrink: 0, minWidth: 108 }}
        onPress={() => {
          onSave(v);
          setV('');
        }}
      />
    </View>
  );
}

const styles = StyleSheet.create({
  heroCard: {
    borderRadius: radii.xl,
    padding: spacing.lg,
    marginBottom: spacing.md,
    borderWidth: 1,
    borderColor: colors.stroke,
  },
  heroLabel: { ...typography.label, marginBottom: spacing.xs },
  heroAmount: {
    fontSize: 32,
    fontWeight: '700',
    color: colors.text,
    letterSpacing: -1,
    maxWidth: '100%',
  },
  heroMoneda: { fontSize: 18, fontWeight: '600', color: colors.accent },
  heroRow: { flexDirection: 'row', alignItems: 'center', marginTop: spacing.sm, flexWrap: 'wrap' },
  heroMeta: { ...typography.small, fontSize: 12 },
  heroMetaDot: { color: colors.textFaint, marginHorizontal: 6 },
  alerta: {
    backgroundColor: colors.alertBg,
    borderWidth: 1,
    borderColor: colors.alertBorder,
    borderRadius: radii.lg,
    padding: spacing.md,
    marginBottom: spacing.md,
  },
  alertaTit: { color: colors.orange, fontWeight: '700', marginBottom: 4 },
  alertaBody: { color: colors.textSecondary, fontSize: 14, lineHeight: 20 },
  barBg: {
    height: 7,
    backgroundColor: colors.barTrack,
    borderRadius: radii.pill,
    marginTop: spacing.md,
    overflow: 'hidden',
  },
  barFill: { height: '100%', borderRadius: radii.pill },
  barFillCat: { height: '100%', borderRadius: radii.pill },
  link: { color: colors.accentBright, marginTop: spacing.md, fontWeight: '600', fontSize: 14 },
  estado: { marginTop: spacing.md, fontWeight: '700', fontSize: 16 },
  estadoOk: { color: colors.success },
  estadoCuidado: { color: colors.warning },
  estadoAlerta: { color: colors.orange },
  estadoSuperado: { color: colors.danger },
  moveRow: { flexDirection: 'row', alignItems: 'flex-start', marginTop: spacing.sm },
  moveDot: {
    width: 6,
    height: 6,
    borderRadius: 3,
    backgroundColor: colors.accentDeep,
    marginTop: 7,
    marginRight: spacing.sm,
  },
  moveText: { flex: 1, minWidth: 0, ...typography.small, color: colors.textSecondary },
  stepRow: { flexDirection: 'row', alignItems: 'center', marginTop: spacing.md },
  stepBadge: {
    width: 28,
    height: 28,
    borderRadius: 14,
    backgroundColor: colors.accentDeep,
    alignItems: 'center',
    justifyContent: 'center',
    marginRight: spacing.md,
  },
  stepNum: { color: '#fff', fontWeight: '800', fontSize: 13 },
  presRow: {
    flexDirection: 'row',
    alignItems: 'center',
    marginTop: spacing.sm,
    flexWrap: 'wrap',
    gap: spacing.sm,
  },
  presRowStack: { flexDirection: 'column', alignItems: 'stretch' },
  presBtnStacked: { alignSelf: 'stretch', width: '100%' },
  inputStacked: { marginRight: 0, width: '100%' },
  input: {
    flex: 1,
    minWidth: 0,
    marginRight: spacing.sm,
    borderWidth: 1,
    borderColor: colors.stroke,
    borderRadius: radii.md,
    paddingVertical: spacing.md,
    paddingHorizontal: spacing.md,
    color: colors.text,
    fontSize: 16,
    backgroundColor: 'rgba(0,0,0,0.2)',
  },
});
