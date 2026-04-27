import React, { useMemo } from 'react';
import { View, Text, StyleSheet, TextInput, TouchableOpacity, Alert, useWindowDimensions } from 'react-native';
import { LinearGradient } from 'expo-linear-gradient';
import { Ionicons } from '@expo/vector-icons';
import ScreenWrap from '../components/ScreenWrap';
import { HeaderConCampana } from '../components/HeaderConCampana';
import { NotificacionBell } from '../components/NotificacionBell';
import UICard from '../components/UICard';
import { PrimaryButton } from '../components/Buttons';
import { useApp, tieneDatosPrevios } from '../context/AppContext';
import {
  formatearNumero,
  CUENTAS,
  calcularSaldosPorCuenta,
  limiteTotalTarjetasCredito,
  totalCupoUtilizadoTarjetasCredito,
  obtenerMesAño,
  montoGastoAfectaSaldo,
  verificarAlertaTarjetaCredito,
  normalizarCategoria,
  normalizarMeta,
} from '../lib/finance';
import { colors, spacing, radii, typography, layoutStyles } from '../theme';

export default function HomeScreen() {
  const { state, ready, replaceState } = useApp();
  const moneda = state?.moneda || '';

  const derived = useMemo(() => {
    if (!state) {
      const nombreMes = [
        'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
        'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre',
      ][new Date().getMonth()];
      return {
        saldosPorCuenta: CUENTAS.reduce((acc, c) => {
          acc[c.id] = 0;
          return acc;
        }, {}),
        topeTarjeta: 0,
        deudaTarjeta: 0,
        saldoActual: 0,
        gastos: [],
        ingresos: [],
        contribuciones: [],
        ingresosMesActual: 0,
        gastosMesActual: 0,
        flujoMes: 0,
        totalGastos: 0,
        alertaTc: { mostrar: false, limite: 0, gastado: 0, porcentaje: 0, tarjetas: [] },
        presupuestoMensual: 0,
        mayorGasto: null,
        categoriasData: [],
        gastosMesPorCategoria: {},
        totalGastosMes: 1,
        metasData: [],
        ultimosGastos: [],
        mesActual: new Date().getMonth(),
        añoActual: new Date().getFullYear(),
        nombreMes,
        estadoMsg: '',
        estadoDetalle: '',
        estadoKind: 'info',
      };
    }

    const ahora = new Date();
    const mesActual = ahora.getMonth();
    const añoActual = ahora.getFullYear();
    const nombreMes = [
      'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
      'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre',
    ][mesActual];

    const saldosPorCuenta = calcularSaldosPorCuenta(state);
    const topeTarjeta = limiteTotalTarjetasCredito(state);
    const deudaTarjeta = totalCupoUtilizadoTarjetasCredito(state);
    const saldoActual = saldosPorCuenta.total || 0;
    const gastos = state.gastos || [];
    const ingresos = state.ingresos || [];
    const contribuciones = state.contribucionesMetas || [];

    const ingresosMesActual = ingresos
      .filter((i) => {
        const { mes, año } = obtenerMesAño(i.fecha);
        return mes === mesActual && año === añoActual;
      })
      .reduce((s, i) => s + (parseFloat(i.cantidad) || 0), 0);

    const gastosMesActual = gastos
      .filter((g) => {
        const { mes, año } = obtenerMesAño(g.fecha);
        return mes === mesActual && año === añoActual;
      })
      .reduce((s, g) => s + montoGastoAfectaSaldo(g), 0);

    const totalGastos = gastos.reduce((s, g) => s + montoGastoAfectaSaldo(g), 0);
    const flujoMes = ingresosMesActual - gastosMesActual;
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
      if (ingresosMesActual > 0 || gastosMesActual > 0) {
        estadoDetalle = `Ingresos del mes: +${formatearNumero(ingresosMesActual)} ${moneda}. Flujo: ${formatearNumero(flujoMes)} ${moneda} (ingresos − gastos).`;
      }
    } else {
      const disponible = presupuestoMensual - gastosMesActual;
      const pctUsado = (gastosMesActual / presupuestoMensual) * 100;
      const extraFlujo =
        ingresosMesActual > 0
          ? ` Ingresos del mes: +${formatearNumero(ingresosMesActual)} ${moneda}. Flujo: ${formatearNumero(flujoMes)} ${moneda} (ingresos − gastos).`
          : '';
      if (disponible > 0 && pctUsado < 80) {
        estadoMsg = '¡Vas bien!';
        estadoDetalle = `Te quedan ${formatearNumero(disponible)} ${moneda} del tope de gasto que fijaste.${extraFlujo}`;
        estadoKind = 'ok';
      } else if (disponible > 0 && pctUsado >= 80) {
        estadoMsg = 'Cuidado, te acercas al límite';
        estadoDetalle = `Te quedan ${formatearNumero(disponible)} ${moneda} del tope de gasto.${extraFlujo}`;
        estadoKind = 'cuidado';
      } else if (disponible === 0) {
        estadoMsg = 'Has agotado tu presupuesto mensual';
        estadoDetalle = `Gastaste exactamente ${formatearNumero(presupuestoMensual)} ${moneda} este mes, según tu tope.${extraFlujo}`;
        estadoKind = 'alerta';
      } else {
        estadoMsg = 'Has superado tu presupuesto';
        estadoDetalle = `Te has pasado en ${formatearNumero(Math.abs(disponible))} ${moneda} respecto al tope.${extraFlujo}`;
        estadoKind = 'superado';
      }
    }

    return {
      saldosPorCuenta,
      topeTarjeta,
      deudaTarjeta,
      saldoActual,
      gastos,
      ingresos,
      contribuciones,
      ingresosMesActual,
      gastosMesActual,
      flujoMes,
      totalGastos,
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
  }, [
    state,
    moneda,
    state?.saldosCuentas,
    state?.tarjetasCredito,
    state?.limiteTarjetaCredito,
    state?.gastos,
    state?.ingresos,
    state?.contribucionesMetas,
  ]);

  if (!ready || !state) {
    return null;
  }

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

  if (!tieneDatosPrevios(state)) {
    return (
      <ScreenWrap contentStyle={{ paddingTop: spacing.sm }}>
        <View
          style={{
            flexDirection: 'row',
            alignItems: 'flex-start',
            justifyContent: 'space-between',
            marginBottom: spacing.xs,
          }}
        >
          <Text style={[typography.hero, { flex: 1, minWidth: 0, paddingRight: spacing.md }]}>
            Bienvenido
          </Text>
          <NotificacionBell />
        </View>
        <Text style={[typography.subtitle, { marginBottom: spacing.lg }]}>Tu dinero, con claridad</Text>
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
      <HeaderConCampana
        label="Resumen"
        title="Inicio"
        subtitle={`${derived.nombreMes} ${derived.añoActual}`}
      />

      <LinearGradient
        colors={['rgba(75, 36, 108, 0.42)', 'rgba(12, 8, 18, 0.45)']}
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

      {derived.alertaTc.tarjetas && derived.alertaTc.tarjetas.length > 0 ? (
        <UICard style={{ marginBottom: spacing.md }}>
          <Text style={typography.label}>Tarjetas de crédito · reloj</Text>
          {derived.alertaTc.limite > 0 ? (
            <Text style={[typography.small, { marginBottom: spacing.sm, color: colors.textSecondary }]}>
              Gasto registrado en la app vs cupo total: {formatearNumero(derived.alertaTc.gastado)} {moneda} (
              {formatearNumero(derived.alertaTc.porcentaje, 1)}%)
            </Text>
          ) : null}
          {derived.alertaTc.tarjetas.map((t, i) => (
            <View
              key={t.id || `tc-${i}`}
              style={{
                marginTop: i > 0 ? spacing.md : spacing.xs,
                paddingTop: i > 0 ? spacing.md : 0,
                borderTopWidth: i > 0 ? 1 : 0,
                borderTopColor: colors.stroke,
              }}
            >
              <Text style={[typography.body, { fontWeight: '600' }]}>{t.nombreEntidad}</Text>
              <View style={styles.tcRelojRow}>
                <Text
                  style={[
                    typography.small,
                    { color: t.alertaCorte ? colors.warning : colors.textSecondary },
                  ]}
                >
                  Corte en {t.diasCorte} d
                </Text>
                <Text style={[typography.small, { color: colors.textFaint, marginHorizontal: 6 }]}>·</Text>
                <Text
                  style={[
                    typography.small,
                    { color: t.alertaPagoUrgente ? colors.danger : colors.textSecondary },
                  ]}
                >
                  Pago en {t.diasPago} d
                </Text>
              </View>
              {(t.etiquetaProxCorte || t.etiquetaProxPago) && (
                <Text style={[typography.small, { marginTop: 4, color: colors.textMuted }]}>
                  Próx. corte: {t.etiquetaProxCorte || '—'} · Próx. pago: {t.etiquetaProxPago || '—'}
                </Text>
              )}
              {t.cupoTotal > 0 ? (
                <Text style={[typography.small, { marginTop: 4, color: colors.textSecondary }]}>
                  Cupo declarado: {formatearNumero(t.cupoUtilizado)} / {formatearNumero(t.cupoTotal)} (
                  {formatearNumero(t.utilPct, 1)}% utilización)
                </Text>
              ) : null}
              {t.tasaEA > 0 ? (
                <Text style={[typography.small, { marginTop: 2, color: colors.textMuted }]}>
                  Tasa E.A.: {formatearNumero(t.tasaEA, 2)}%
                </Text>
              ) : null}
              {(t.alertaUtil || t.alertaPagoUrgente || t.alertaCorte) && (
                <View style={styles.tcChipsRow}>
                  {t.alertaPagoUrgente ? <Text style={styles.tcChipDanger}>Pago próximo</Text> : null}
                  {t.alertaCorte ? <Text style={styles.tcChipWarn}>Corte próximo</Text> : null}
                  {t.alertaUtil ? <Text style={styles.tcChipWarn}>Uso alto del cupo</Text> : null}
                </View>
              )}
            </View>
          ))}
        </UICard>
      ) : derived.alertaTc.mostrar ? (
        <View style={styles.alerta}>
          <Text style={styles.alertaTit}>Tarjeta de crédito</Text>
          <Text style={styles.alertaBody}>
            {formatearNumero(derived.alertaTc.gastado)} {moneda} usados (
            {formatearNumero(derived.alertaTc.porcentaje, 1)}% del límite).
          </Text>
        </View>
      ) : null}

      <UICard>
        <Text style={typography.label}>Por cuenta</Text>
        {CUENTAS.map((c) => (
          <View key={c.id} style={{ marginBottom: spacing.sm }}>
            <View style={layoutStyles.rowBetween}>
              <Text style={[typography.body, layoutStyles.rowLabel]}>
                {c.id === 'tarjetaCredito' ? `${c.nombre} (cupo libre)` : c.nombre}
              </Text>
              <Text style={[typography.monoAmount, layoutStyles.rowValue]}>
                {formatearNumero(derived.saldosPorCuenta[c.id] ?? 0)} {moneda}
              </Text>
            </View>
            {c.id === 'tarjetaCredito' && (derived.topeTarjeta > 0 || derived.deudaTarjeta > 0) ? (
              <Text style={[typography.small, { marginTop: 4, color: colors.textFaint, paddingRight: 4 }]}>
                Tope: {formatearNumero(derived.topeTarjeta)} {moneda} · Deuda:{' '}
                {formatearNumero(derived.deudaTarjeta)} {moneda} · arriba: tope − deuda (cupo usado) = cupo libre. Actualiza
                la deuda en Saldo cuando el banco cambie.
              </Text>
            ) : c.id === 'tarjetaCredito' && (state.tarjetasCredito || []).length > 0 ? (
              <Text style={[typography.small, { marginTop: 4, color: colors.textFaint, paddingRight: 4 }]}>
                En Saldo → Tarjeta: cupo total y cupo usado (deuda). El importe de arriba = cupo total − deuda.
              </Text>
            ) : null}
          </View>
        ))}
        <Text style={[typography.small, { marginTop: spacing.sm, color: colors.textFaint }]}>
          Incluye saldo 0,00: así ves cada caja aunque te hayas quedado sin plata; suma un ingreso en Ingresos o
          ajusta en Saldo.
        </Text>
        <Text style={[typography.small, { marginTop: spacing.sm }]}>
          {derived.nombreMes}: ingresos {formatearNumero(derived.ingresosMesActual)} {moneda} · gastos{' '}
          {formatearNumero(derived.gastosMesActual)} {moneda}
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
              <Text style={[typography.body, layoutStyles.rowLabel]}>Planeado (tope)</Text>
              <Text style={[typography.monoAmount, layoutStyles.rowValue]}>
                {formatearNumero(derived.presupuestoMensual)} {moneda}
              </Text>
            </View>
            <View style={layoutStyles.rowBetween}>
              <Text style={[typography.body, layoutStyles.rowLabel]}>Ingresos (mes)</Text>
              <Text style={[typography.monoAmount, layoutStyles.rowValue, { color: colors.mint }]}>
                +{formatearNumero(derived.ingresosMesActual)} {moneda}
              </Text>
            </View>
            <View style={layoutStyles.rowBetween}>
              <Text style={[typography.body, layoutStyles.rowLabel]}>Gastos (mes)</Text>
              <Text style={[typography.monoAmount, layoutStyles.rowValue, { color: colors.danger }]}>
                −{formatearNumero(derived.gastosMesActual)} {moneda}
              </Text>
            </View>
            <View style={layoutStyles.rowBetween}>
              <Text style={[typography.body, layoutStyles.rowLabel]}>Flujo (mes)</Text>
              <Text
                style={[
                  typography.monoAmount,
                  layoutStyles.rowValue,
                  { color: derived.flujoMes >= 0 ? colors.mint : colors.danger },
                ]}
              >
                {derived.flujoMes >= 0 ? '+' : ''}
                {formatearNumero(derived.flujoMes)} {moneda}
              </Text>
            </View>
            <View style={layoutStyles.rowBetween}>
              <Text style={[typography.body, layoutStyles.rowLabel]}>Disponible (tope)</Text>
              <Text style={[typography.monoAmount, layoutStyles.rowValue, { color: colors.mint }]}>
                {formatearNumero(Math.max(0, derived.presupuestoMensual - derived.gastosMesActual))}{' '}
                {moneda}
              </Text>
            </View>
            <Text style={[typography.small, { marginBottom: spacing.sm, color: colors.textFaint, marginTop: 2 }]}>
              Flujo: ingresos − gastos del mes. Disponible (tope): tope fijado − gastos; el ingreso no suma a esa
              cifra.
            </Text>
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
        {(derived.topeTarjeta > 0 ||
          derived.deudaTarjeta > 0 ||
          (state.tarjetasCredito || []).length > 0 ||
          (derived.saldosPorCuenta.tarjetaCredito || 0) > 0) && (
          <View style={{ marginTop: spacing.sm }}>
            <Text style={typography.body}>
              Cupo libre (tarjeta):{' '}
              <Text style={{ fontWeight: '700', color: colors.mint }}>
                {formatearNumero(derived.saldosPorCuenta.tarjetaCredito ?? 0)} {moneda}
              </Text>
            </Text>
            {derived.topeTarjeta > 0 || derived.deudaTarjeta > 0 ? (
              <Text style={[typography.small, { marginTop: 4, color: colors.textFaint }]}>
                Tope {formatearNumero(derived.topeTarjeta)} {moneda} · Deuda {formatearNumero(derived.deudaTarjeta)}{' '}
                {moneda} · cupo = tope − deuda
              </Text>
            ) : null}
          </View>
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
              const limiteCat =
                cat.limite != null && String(cat.limite).trim() !== ''
                  ? parseFloat(cat.limite)
                  : NaN;
              const tieneLimite = Number.isFinite(limiteCat) && limiteCat > 0;
              const superadoCategoria = tieneLimite && monto > limiteCat;
              return (
                <View
                  key={cat.nombre}
                  style={[
                    { marginBottom: spacing.md },
                    superadoCategoria && {
                      padding: spacing.sm,
                      borderRadius: radii.md,
                      backgroundColor: 'rgba(199, 123, 136, 0.16)',
                      borderWidth: 1,
                      borderColor: 'rgba(199, 123, 136, 0.55)',
                    },
                  ]}
                >
                  <View style={layoutStyles.rowBetween}>
                    <Text
                      style={[
                        typography.body,
                        layoutStyles.rowLabel,
                        superadoCategoria && { color: colors.danger, fontWeight: '700' },
                      ]}
                    >
                      {cat.icono} {cat.nombre}
                    </Text>
                    <Text
                      style={[
                        typography.monoAmount,
                        layoutStyles.rowValue,
                        superadoCategoria && { color: colors.danger },
                      ]}
                    >
                      {formatearNumero(monto)} {moneda}
                    </Text>
                  </View>
                  <View
                    style={[
                      styles.barBg,
                      superadoCategoria && { backgroundColor: 'rgba(199, 123, 136, 0.28)' },
                    ]}
                  >
                    <View
                      style={[
                        styles.barFillCat,
                        {
                          width: `${Math.min(100, pct)}%`,
                          backgroundColor: superadoCategoria ? colors.danger : cat.color,
                        },
                      ]}
                    />
                  </View>
                  {superadoCategoria ? (
                    <Text
                      style={[typography.small, { marginTop: spacing.xs, color: colors.danger, fontWeight: '600' }]}
                    >
                      Sobre límite ({formatearNumero(limiteCat)} {moneda}): +{formatearNumero(monto - limiteCat)}{' '}
                      {moneda}
                    </Text>
                  ) : null}
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
            const m = normalizarMeta(meta);
            const acum = (state.contribucionesMetas || [])
              .filter((c) => c.metaId === m.id)
              .reduce((s, c) => s + c.cantidad, 0);
            const obj = parseFloat(m.objetivo) || 0;
            const pct = obj > 0 ? Math.min(100, (acum / obj) * 100) : 0;
            return (
              <View key={m.id} style={{ marginBottom: spacing.md }}>
                <View style={layoutStyles.rowBetween}>
                  <View style={styles.metaInicioFila}>
                    <View style={styles.metaInicioIcono}>
                      <Ionicons name={m.icono} size={19} color={colors.mint} />
                    </View>
                    <Text
                      style={[typography.body, layoutStyles.rowLabel, { flex: 1, minWidth: 0 }]}
                      numberOfLines={2}
                    >
                      {m.nombre}
                    </Text>
                  </View>
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
  tcRelojRow: { flexDirection: 'row', flexWrap: 'wrap', alignItems: 'center', marginTop: 6 },
  tcChipsRow: { flexDirection: 'row', flexWrap: 'wrap', marginTop: spacing.sm },
  tcChipDanger: {
    borderRadius: radii.sm,
    paddingVertical: 4,
    paddingHorizontal: 8,
    backgroundColor: 'rgba(251, 113, 133, 0.22)',
    color: colors.danger,
    fontSize: 12,
    fontWeight: '600',
    marginRight: 8,
    marginBottom: 4,
    overflow: 'hidden',
  },
  tcChipWarn: {
    borderRadius: radii.sm,
    paddingVertical: 4,
    paddingHorizontal: 8,
    backgroundColor: 'rgba(251, 191, 36, 0.2)',
    color: colors.warning,
    fontSize: 12,
    fontWeight: '600',
    marginRight: 8,
    marginBottom: 4,
    overflow: 'hidden',
  },
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
  metaInicioFila: {
    flexDirection: 'row',
    alignItems: 'center',
    flex: 1,
    minWidth: 0,
    marginRight: spacing.sm,
  },
  metaInicioIcono: {
    width: 32,
    height: 32,
    borderRadius: 16,
    backgroundColor: 'rgba(125, 193, 145, 0.12)',
    borderWidth: 1,
    borderColor: 'rgba(125, 193, 145, 0.35)',
    alignItems: 'center',
    justifyContent: 'center',
    marginRight: spacing.sm,
    flexShrink: 0,
  },
});
