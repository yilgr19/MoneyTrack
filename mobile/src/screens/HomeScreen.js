import React, { useMemo, useState } from 'react';
import { View, Text, StyleSheet, TextInput, TouchableOpacity, Alert, useWindowDimensions } from 'react-native';
import { useNavigation } from '@react-navigation/native';
import { LinearGradient } from 'expo-linear-gradient';
import { Ionicons } from '@expo/vector-icons';
import ScreenWrap from '../components/ScreenWrap';
import { HeaderConCampana } from '../components/HeaderConCampana';
import { NotificacionBell } from '../components/NotificacionBell';
import UICard from '../components/UICard';
import DonutChart from '../components/charts/DonutChart';
import CategoriaGastoBarFun from '../components/CategoriaGastoBarFun';
import { PrimaryButton } from '../components/Buttons';
import { useApp, tieneDatosPrevios } from '../context/AppContext';
import {
  formatearNumero,
  CUENTAS,
  calcularSaldosPorCuenta,
  cuentaVisibleEnResumenInicio,
  limiteTotalTarjetasCredito,
  totalCupoUtilizadoTarjetasCredito,
  obtenerMesAño,
  montoGastoAfectaSaldoEnMes,
  verificarAlertaTarjetaCredito,
  normalizarCategoria,
  normalizarMeta,
  totalSaldoBolsillos,
} from '../lib/finance';
import {
  colors,
  spacing,
  radii,
  typography,
  layoutStyles,
  iconSemantic,
  colorIconoMetaDesdeNombre,
  shadows,
} from '../theme';

export default function HomeScreen() {
  const { state, ready, replaceState } = useApp();
  const navigation = useNavigation();
  const { width: winW } = useWindowDimensions();
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
        cuentasInicio: [],
        totalEnBolsillos: 0,
        totalAportesMetas: 0,
      };
    }

    const ahora = new Date();
    const totalEnBolsillos = totalSaldoBolsillos(state);
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
        if (i.esRetiroBolsillo) return false;
        const { mes, año } = obtenerMesAño(i.fecha);
        return mes === mesActual && año === añoActual;
      })
      .reduce((s, i) => s + (parseFloat(i.cantidad) || 0), 0);

    const gastosMesActual = gastos.reduce(
      (s, g) => s + montoGastoAfectaSaldoEnMes(g, state, mesActual, añoActual),
      0
    );

    const totalGastos = gastos.reduce(
      (s, g) => s + (g.esTransferenciaBolsillo ? 0 : parseFloat(g.cantidad) || 0),
      0
    );
    const flujoMes = ingresosMesActual - gastosMesActual;
    const alertaTc = verificarAlertaTarjetaCredito(state);
    const presupuestoMensual = state.presupuestoMensual || 0;

    const gastosConEfectoMes = gastos.map((g) => ({
      g,
      m: montoGastoAfectaSaldoEnMes(g, state, mesActual, añoActual),
    }));
    const conEfecto = gastosConEfectoMes.filter((x) => x.m > 0);
    const mayorGasto =
      conEfecto.length > 0
        ? conEfecto.reduce((max, x) => (x.m > max.m ? x : max), conEfecto[0]).g
        : null;

    const categoriasData = (state.categorias || []).map(normalizarCategoria);
    const gastosMesPorCategoria = {};
    gastosConEfectoMes.forEach(({ g, m }) => {
      if (m <= 0) return;
      const cat = g.categoria || 'Otros';
      gastosMesPorCategoria[cat] = (gastosMesPorCategoria[cat] || 0) + m;
    });
    const totalGastosMes = Object.values(gastosMesPorCategoria).reduce((s, v) => s + v, 0) || 1;

    const metasData = state.metas || [];
    const totalAportesMetas = contribuciones.reduce((s, c) => s + (parseFloat(c.cantidad) || 0), 0);
    const ultimosGastos = gastos.slice().reverse().slice(0, 8);

    let estadoMsg = '';
    let estadoDetalle = '';
    let estadoKind = 'info';
    if (presupuestoMensual <= 0) {
      estadoMsg = 'Sin tope mensual no hay semáforo (abajo o Saldo).';
      if (ingresosMesActual > 0 || gastosMesActual > 0) {
        estadoDetalle = `+${formatearNumero(ingresosMesActual)} / flujo ${formatearNumero(flujoMes)} ${moneda}`;
      }
    } else {
      const disponible = presupuestoMensual - gastosMesActual;
      const pctUsado = (gastosMesActual / presupuestoMensual) * 100;
      if (disponible > 0 && pctUsado < 80) {
        estadoMsg = '¡Dentro del tope!';
        estadoDetalle = `Quedan ${formatearNumero(disponible)} ${moneda} de tu límite de gasto.`;
        estadoKind = 'ok';
      } else if (disponible > 0 && pctUsado >= 80) {
        estadoMsg = 'Cerca del tope del mes';
        estadoDetalle = `Quedan ${formatearNumero(disponible)} ${moneda}.`;
        estadoKind = 'cuidado';
      } else if (disponible === 0) {
        estadoMsg = 'Límite de gasto alcanzado';
        estadoDetalle = `${formatearNumero(presupuestoMensual)} ${moneda} este mes.`;
        estadoKind = 'alerta';
      } else {
        estadoMsg = 'Sobre el tope fijado';
        estadoDetalle = `+${formatearNumero(Math.abs(disponible))} ${moneda} sobre límite.`;
        estadoKind = 'superado';
      }
    }

    const cuentasInicio = CUENTAS.filter((c) =>
      cuentaVisibleEnResumenInicio(c.id, state, saldosPorCuenta)
    );

    return {
      saldosPorCuenta,
      totalEnBolsillos,
      cuentasInicio,
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
      totalAportesMetas,
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
    state?.bancosDetalle,
    state?.plataformasDetalle,
    state?.tarjetasCredito,
    state?.limiteTarjetaCredito,
    state?.gastos,
    state?.ingresos,
    state?.contribucionesMetas,
    state?.bolsillos,
  ]);

  if (!ready || !state) {
    return null;
  }

  const pctPresupuesto =
    derived.presupuestoMensual > 0
      ? Math.min(100, (derived.gastosMesActual / derived.presupuestoMensual) * 100)
      : 0;

  const segmentosDonutCategorias = useMemo(() => {
    const entries = Object.entries(derived.gastosMesPorCategoria || {})
      .filter(([, v]) => (parseFloat(v) || 0) > 0)
      .sort((a, b) => (parseFloat(b[1]) || 0) - (parseFloat(a[1]) || 0));
    if (entries.length === 0) return [];
    const top = entries.slice(0, 5);
    const restSum = entries.slice(5).reduce((s, [, v]) => s + (parseFloat(v) || 0), 0);
    const out = top.map(([nombre, monto]) => {
      const c = (derived.categoriasData || []).find((x) => x.nombre === nombre);
      return {
        value: parseFloat(monto) || 0,
        color: c?.color || colors.textMuted,
        label: `${c?.icono ? `${c.icono} ` : ''}${nombre}`.trim(),
      };
    });
    if (restSum > 0.01) {
      out.push({ value: restSum, color: '#94a3b8', label: 'Otros' });
    }
    return out;
  }, [derived.gastosMesPorCategoria, derived.categoriasData]);

  const segmentosIngresoGasto = useMemo(() => {
    const ing = derived.ingresosMesActual || 0;
    const gas = derived.gastosMesActual || 0;
    if (ing <= 0 && gas <= 0) return [];
    return [
      { value: Math.max(0, ing), color: colors.mint, label: 'Ingresos' },
      { value: Math.max(0, gas), color: colors.danger, label: 'Gastos' },
    ];
  }, [derived.ingresosMesActual, derived.gastosMesActual]);

  const centroDonutCategorias = useMemo(() => {
    if (!segmentosDonutCategorias.length) return { line1: undefined, line2: undefined };
    const g = derived.gastosMesActual || 0;
    return {
      line1: `${formatearNumero(g)} ${moneda}`.trim(),
      line2: 'Total gastos (mes)',
    };
  }, [segmentosDonutCategorias, derived.gastosMesActual, moneda]);

  const centroDonutFlujo = useMemo(() => {
    const ing = derived.ingresosMesActual || 0;
    const gas = derived.gastosMesActual || 0;
    const t = ing + gas;
    if (t <= 0) return { line1: undefined, line2: undefined };
    const pi = Math.round((ing / t) * 100);
    const pg = Math.round((gas / t) * 100);
    return {
      line1: `${pi}%  ·  ${pg}%`,
      line2: 'Ingresos · gastos (mes)',
    };
  }, [derived.ingresosMesActual, derived.gastosMesActual]);

  const segmentosAhorroBolsillosMetas = useMemo(() => {
    const b = derived.totalEnBolsillos || 0;
    const m = derived.totalAportesMetas || 0;
    if (b <= 0 && m <= 0) return [];
    const out = [];
    if (b > 0) out.push({ value: b, color: '#2dd4bf', label: 'Bolsillos' });
    if (m > 0) out.push({ value: m, color: colors.accentGold, label: 'Metas (aportes)' });
    return out;
  }, [derived.totalEnBolsillos, derived.totalAportesMetas]);

  const totalAhorroBolsillosYMetas = (derived.totalEnBolsillos || 0) + (derived.totalAportesMetas || 0);
  const ingresoMesAhorro = derived.ingresosMesActual || 0;
  const porcentajeAhorroSobreIngreso =
    ingresoMesAhorro > 0 ? (totalAhorroBolsillosYMetas / ingresoMesAhorro) * 100 : 0;
  const ahorroSuperaTercioIngreso = ingresoMesAhorro > 0 && totalAhorroBolsillosYMetas >= 0.3 * ingresoMesAhorro;

  const centroDonutAhorro = useMemo(() => {
    if (ingresoMesAhorro <= 0) {
      return { line1: undefined, line2: undefined };
    }
    const p = Math.min(999, Math.round(porcentajeAhorroSobreIngreso));
    return {
      line1: `${p}%`,
      line2: 'Ahorro vs ingreso (mes)',
    };
  }, [ingresoMesAhorro, porcentajeAhorroSobreIngreso]);

  const indiceMensajeAhorro = useMemo(() => {
    const s = Math.floor(
      totalAhorroBolsillosYMetas * 3 + ingresoMesAhorro * 2 + (derived.totalAportesMetas || 0)
    );
    return Math.abs(s) % 5;
  }, [totalAhorroBolsillosYMetas, ingresoMesAhorro, derived.totalAportesMetas]);

  const chartsStack = winW < 400;
  const chartSize = chartsStack ? 148 : 136;
  const chartSizeAhorro = Math.min(200, Math.max(160, winW * 0.48));

  const FRASES_AHORRO_30_SI = [
    '¡Lo lograste! Tu ahorro en bolsillos y metas suma al menos el 30% de lo que entra este mes. Sigue reforzando el hábito: cada quincena cuenta.',
    'Excelente: estás en terreno sano. Superaste la guía del 30% frente a tus ingresos del mes. Mantén el ritmo y sube el listón cuando puedas.',
    'Tu disciplina se nota: bolsillos y metas ya cubren al menos un tercio de lo ingresado. Sigue y amplía la meta; la calma financiera se construye así.',
    'Bravo. Por encima del mínimo del 30%: separar ahorro del gasto te da margen. Un paso más hacia la estabilidad que buscas.',
    'Así se avanza: con este nivel de ahorro respecto al ingreso del mes, vas bien encaminado. Sigue priorizando el colchón y celebra el avance.',
  ];
  const FRASES_AHORRO_30_NO = [
    'Aún no llegas al 30%: no pasa nada. Pequeñas aportaciones a bolsillos o metas suman; apunta a reservar al menos un tercio de lo que entra.',
    'La guía es ahorrar al menos el 30% de tus ingresos del mes. Revisa bolsillos y metas: separa aunque sea poco, con constancia ganas.',
    'Sigue adelante: la estabilidad pasa por un colchón. Cada aporte a metas o movimiento a bolsillos te acerca al objetivo del 30%.',
    'Aún estás bajo el piso sugerido. Ajusta el ritmo: nada más cobrar, aparta; el 30% es guía, no presión, pero te protege en imprevistos.',
    'Puedes lograrlo: bolsillos y aportes a metas suman lo que se considera ahorro aquí. Sube poco a poco hasta al menos un tercio del ingreso.',
  ];
  const fraseAhorroMotiv = ahorroSuperaTercioIngreso
    ? FRASES_AHORRO_30_SI[indiceMensajeAhorro]
    : FRASES_AHORRO_30_NO[indiceMensajeAhorro];

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
        <Text style={[typography.subtitle, { marginBottom: spacing.lg }]}>Tres pasos</Text>
        <UICard accent>
          <Text style={typography.label}>Primeros pasos</Text>
          <Step n={1} text="Saldo o Ingresos (Más)" />
          <Step n={2} text="Categorías (Más)" />
          <Step n={3} text="Registrar en Gastos" />
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
        colors={['rgba(28, 26, 38, 0.98)', 'rgba(12, 11, 18, 1)']}
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
        {derived.totalEnBolsillos > 0 ? (
          <Text style={styles.heroBolsillos}>
            Bolsillos: {formatearNumero(derived.totalEnBolsillos)} {moneda} · no al total
          </Text>
        ) : null}
        <View style={styles.heroRow}>
          <View style={styles.heroStat}>
            <Text style={styles.heroStatLab}>Ingresos (mes)</Text>
            <Text style={styles.heroStatVal}>
              +{formatearNumero(derived.ingresosMesActual)} {moneda}
            </Text>
          </View>
          <View style={styles.heroStatSep} />
          <View style={styles.heroStat}>
            <Text style={styles.heroStatLab}>Gastos (mes)</Text>
            <Text style={[styles.heroStatVal, { color: colors.danger }]}>
              −{formatearNumero(derived.gastosMesActual)} {moneda}
            </Text>
          </View>
        </View>
      </LinearGradient>

      {derived.alertaTc.tarjetas && derived.alertaTc.tarjetas.length > 0 ? (
        <UICard style={{ marginBottom: spacing.md }}>
          <Text style={[typography.label, { marginBottom: spacing.sm }]}>TC · corte y pago</Text>
          {derived.alertaTc.tarjetas.map((t, i) => (
            <View
              key={t.id || `tc-${i}`}
              style={{
                marginTop: i > 0 ? spacing.md : 0,
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
                    {
                      color: t.pagoCorteMuestraAlDia
                        ? colors.mint
                        : t.alertaCorte
                          ? colors.warning
                          : colors.textSecondary,
                    },
                  ]}
                >
                  {t.pagoCorteMuestraAlDia
                    ? 'Corte: al día'
                    : t.corteHoy
                      ? 'Corte: hoy'
                      : `Corte en ${t.diasCorte} d`}
                </Text>
                <Text style={[typography.small, { color: colors.textFaint, marginHorizontal: 6 }]}>·</Text>
                <Text
                  style={[
                    typography.small,
                    {
                      color: t.pagoCorteMuestraAlDia
                        ? colors.mint
                        : t.alertaPagoUrgente
                          ? colors.danger
                          : colors.textSecondary,
                    },
                  ]}
                >
                  {t.pagoCorteMuestraAlDia ? 'Pago: al día' : `Pago en ${t.diasPago} d`}
                </Text>
              </View>
              {(t.etiquetaProxCorte || t.etiquetaProxPago) && (
                <Text style={[typography.small, { marginTop: 4, color: colors.textMuted }]}>
                  Próx. corte: {t.etiquetaProxCorte || '—'} · Próx. pago: {t.etiquetaProxPago || '—'}
                </Text>
              )}
              {(t.alertaUtil || t.alertaPagoUrgente || t.alertaCorte) && (
                <View style={styles.tcChipsRow}>
                  {t.alertaPagoUrgente ? <Text style={styles.tcChipDanger}>Pago próximo</Text> : null}
                  {t.alertaCorte ? <Text style={styles.tcChipWarn}>Corte próximo</Text> : null}
                  {t.alertaUtil ? <Text style={styles.tcChipWarn}>Uso alto del cupo</Text> : null}
                </View>
              )}
            </View>
          ))}
          <TouchableOpacity
            onPress={() => navigation.navigate('Mas', { screen: 'ExtractosTarjetas' })}
            style={styles.btnIrExtractos}
            activeOpacity={0.85}
          >
            <Ionicons
              name="receipt-outline"
              size={18}
              color={iconSemantic.moreMenu.ExtractosTarjetas.fg}
            />
            <Text style={styles.btnIrExtractosTxt}>Extractos de tarjeta</Text>
            <Ionicons name="chevron-forward" size={18} color={colors.textFaint} />
          </TouchableOpacity>
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
        {derived.cuentasInicio.length === 0 ? (
          <Text style={[typography.small, { color: colors.textFaint, marginTop: 4 }]}>
            Añade saldo o un ingreso (pestaña Saldo).
          </Text>
        ) : (
          derived.cuentasInicio.map((c) => (
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
                  Tope {formatearNumero(derived.topeTarjeta)} · deuda {formatearNumero(derived.deudaTarjeta)}{' '}
                  {moneda} · ajusta en Saldo
                </Text>
              ) : c.id === 'tarjetaCredito' && (state.tarjetasCredito || []).length > 0 ? (
                <Text style={[typography.small, { marginTop: 4, color: colors.textFaint, paddingRight: 4 }]}>
                  Saldo: cupo total − deuda (Saldo → Tarjeta)
                </Text>
              ) : null}
            </View>
          ))
        )}
        <Text style={[typography.small, { marginTop: spacing.sm }]}>
          {derived.nombreMes}: ingresos {formatearNumero(derived.ingresosMesActual)} {moneda} · gastos{' '}
          {formatearNumero(derived.gastosMesActual)} {moneda}
          {derived.contribuciones.reduce((s, c) => s + c.cantidad, 0) > 0
            ? ` · Metas ${formatearNumero(derived.contribuciones.reduce((s, c) => s + c.cantidad, 0))} ${moneda}`
            : ''}
        </Text>
      </UICard>

      <UICard>
        <Text style={styles.analisisSectionTit}>Análisis</Text>
        <Text style={styles.analisisSectionSub}>
          Periodo: mes en curso. Ahorro = saldo en bolsillos + aportes acumulados a metas; se compara con el ingreso
          del mes y la guía del 30%.
        </Text>
        <View style={[styles.chartsRow, chartsStack && styles.chartsRowStack]}>
          <View style={[styles.chartPanel, chartsStack && styles.chartPanelFull]}>
            <DonutChart
              segments={segmentosDonutCategorias}
              title="Gastos por categoría"
              emptyHint="Sin gastos el mes"
              size={chartSize}
              centerLine1={centroDonutCategorias.line1}
              centerLine2={centroDonutCategorias.line2}
            />
          </View>
          <View style={[styles.chartPanel, chartsStack && styles.chartPanelFull]}>
            <DonutChart
              segments={segmentosIngresoGasto}
              title="Ingreso y gasto"
              emptyHint="Sin ingresos ni gastos"
              size={chartSize}
              centerLine1={centroDonutFlujo.line1}
              centerLine2={centroDonutFlujo.line2}
            />
          </View>
        </View>

        <View style={[styles.chartPanel, styles.chartPanelFull, styles.ahorroPanel]}>
          <Text style={styles.ahorroPanelTit}>Ahorro: bolsillos · metas</Text>
          <Text style={styles.ahorroPanelSub}>
            Ahorro considerado: {formatearNumero(derived.totalEnBolsillos)} {moneda} (bolsillos) + {formatearNumero(derived.totalAportesMetas)} {moneda} (metas) ={' '}
            <Text style={{ fontWeight: '700', color: colors.text }}>{formatearNumero(totalAhorroBolsillosYMetas)} {moneda}</Text>
            {ingresoMesAhorro > 0
              ? ` · Guía: al menos 30% del ingreso del mes = ${formatearNumero(ingresoMesAhorro * 0.3)} ${moneda}`
              : ' · Añade ingresos del mes para calcular el % y la guía del 30%.'}
          </Text>
          {ingresoMesAhorro > 0 && !segmentosAhorroBolsillosMetas.length ? (
            <View style={styles.ahorroCeroBloq}>
              <Ionicons name="wallet-outline" size={36} color={colors.textFaint} style={{ marginBottom: spacing.sm }} />
              <Text style={styles.ahorroCeroPct}>0%</Text>
              <Text style={styles.ahorroCeroLey} numberOfLines={2}>
                del ingreso (mes) en bolsillos + metas
              </Text>
              <Text style={styles.ahorroMetaG}>
                Mínimo sugerido 30%: {formatearNumero(ingresoMesAhorro * 0.3)} {moneda}
              </Text>
            </View>
          ) : (
            <DonutChart
              segments={segmentosAhorroBolsillosMetas}
              title=""
              emptyHint={
                ingresoMesAhorro > 0
                  ? 'Carga ahorro en bolsillos o aporta a metas (Más).'
                  : 'Sin ahorro visible ni ingreso del mes; revisa bolsillos, metas o ingresos (Más).'
              }
              size={chartSizeAhorro}
              centerLine1={ingresoMesAhorro > 0 ? centroDonutAhorro.line1 : undefined}
              centerLine2={ingresoMesAhorro > 0 ? centroDonutAhorro.line2 : undefined}
            />
          )}
          {ingresoMesAhorro > 0 ? (
            <Text
              style={[
                styles.ahorroMens,
                ahorroSuperaTercioIngreso ? styles.ahorroMensOk : styles.ahorroMensTip,
              ]}
            >
              {fraseAhorroMotiv}
            </Text>
          ) : (
            <Text style={styles.ahorroMensTip}>
              Registra en <Text style={{ fontWeight: '700' }}>Más → Ingresos</Text> lo que entra este mes; así
              podemos comparar bolsillos+metas con el 30% recomendado.
            </Text>
          )}
        </View>
        {derived.mayorGasto ? (
          <Text style={typography.body}>
            Mayor gasto: <Text style={{ fontWeight: '700', color: colors.text }}>{derived.mayorGasto.nombre}</Text> —{' '}
            {formatearNumero(derived.mayorGasto.cantidad)} {moneda}
          </Text>
        ) : (
          <Text style={typography.small}>Sin gastos este mes.</Text>
        )}
        {(derived.topeTarjeta > 0 ||
          derived.deudaTarjeta > 0 ||
          (state.tarjetasCredito || []).length > 0 ||
          (derived.saldosPorCuenta.tarjetaCredito || 0) > 0) && (
          <View style={{ marginTop: spacing.sm }}>
            <Text style={typography.body}>
              Cupo libre:{' '}
              <Text style={{ fontWeight: '700', color: colors.mint }}>
                {formatearNumero(derived.saldosPorCuenta.tarjetaCredito ?? 0)} {moneda}
              </Text>
            </Text>
            {derived.topeTarjeta > 0 || derived.deudaTarjeta > 0 ? (
              <Text style={[typography.small, { marginTop: 4, color: colors.textFaint }]}>
                Tope {formatearNumero(derived.topeTarjeta)} · deuda {formatearNumero(derived.deudaTarjeta)} {moneda}
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
        <Text style={[typography.small, { color: colors.textMuted, marginBottom: spacing.md, lineHeight: 20 }]}>
          Cómo repartieron tus salidas: medallas, colores y barras con ritmo. ¡A competir con el bolsillo!
        </Text>
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
            .map((cat, idx) => {
              const monto = derived.gastosMesPorCategoria[cat.nombre] || 0;
              const pct = derived.totalGastosMes > 0 ? (monto / derived.totalGastosMes) * 100 : 0;
              const limiteCat =
                cat.limite != null && String(cat.limite).trim() !== ''
                  ? parseFloat(cat.limite)
                  : NaN;
              const tieneLimite = Number.isFinite(limiteCat) && limiteCat > 0;
              const superadoCategoria = tieneLimite && monto > limiteCat;
              return (
                <CategoriaGastoBarFun
                  key={cat.nombre}
                  cat={cat}
                  monto={monto}
                  pct={pct}
                  moneda={moneda}
                  index={idx}
                  superadoCategoria={superadoCategoria}
                  limiteCat={limiteCat}
                  formatearNumero={formatearNumero}
                />
              );
            })
        )}
      </UICard>

      <UICard>
        <Text style={typography.label}>Presupuesto mensual</Text>
        <View style={styles.guiaPresup}>
          <Text style={styles.guiaPresupTit}>Guía rápida</Text>
          <Text style={styles.guiaPresupTxt}>
            <Text style={styles.guiaPresupBold}>Tope</Text> = límite de gasto del mes (lo fijas tú).{' '}
            <Text style={styles.guiaPresupBold}>Disponible</Text> = tope − gastos del mes.{' '}
            <Text style={styles.guiaPresupBold}>Ingreso</Text> y <Text style={styles.guiaPresupBold}>flujo</Text> son
            informativos (entradas y entradas − gastos); no suman al tope.
          </Text>
        </View>
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
              Disponible = tope − gastos (el ingreso no sube este saldo)
            </Text>
            <View style={styles.barBg}>
              <LinearGradient
                colors={['rgba(88, 82, 108, 0.85)', 'rgba(60, 52, 78, 0.95)']}
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
        <Text style={typography.label}>Últimos movimientos</Text>
        {derived.gastos.length === 0 ? (
          <Text style={typography.small}>Aún no hay gastos.</Text>
        ) : (
          <>
            <Text style={[typography.small, { marginBottom: spacing.sm, color: colors.textFaint }]}>
              Total: {formatearNumero(derived.totalGastos)} {moneda}
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
                      <Ionicons name={m.icono} size={19} color={colorIconoMetaDesdeNombre(m.icono)} />
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
  guiaPresup: {
    backgroundColor: 'rgba(0,0,0,0.2)',
    borderRadius: radii.md,
    padding: spacing.md,
    marginBottom: spacing.md,
    borderWidth: 1,
    borderColor: 'rgba(255,255,255,0.06)',
    borderLeftWidth: 3,
    borderLeftColor: 'rgba(160, 140, 190, 0.55)',
  },
  guiaPresupTit: {
    fontSize: 10,
    fontWeight: '600',
    letterSpacing: 1.1,
    textTransform: 'uppercase',
    marginBottom: spacing.xs,
    color: colors.textMuted,
  },
  guiaPresupTxt: { ...typography.small, color: colors.textSecondary, lineHeight: 20 },
  guiaPresupBold: { fontWeight: '600', color: colors.text },
  analisisSectionTit: {
    fontSize: 10,
    fontWeight: '600',
    letterSpacing: 1.2,
    textTransform: 'uppercase',
    color: colors.textMuted,
    marginBottom: 6,
  },
  analisisSectionSub: {
    ...typography.small,
    color: colors.textFaint,
    marginBottom: spacing.md,
    lineHeight: 18,
  },
  chartPanel: {
    flex: 1,
    minWidth: 148,
    maxWidth: '48%',
    padding: spacing.md,
    borderRadius: radii.lg,
    backgroundColor: 'rgba(0,0,0,0.16)',
    borderWidth: 1,
    borderColor: 'rgba(255,255,255,0.05)',
  },
  chartPanelFull: { maxWidth: '100%', width: '100%' },
  ahorroPanel: {
    marginTop: spacing.sm,
    borderColor: 'rgba(125, 193, 145, 0.22)',
    backgroundColor: 'rgba(0,0,0,0.18)',
  },
  ahorroPanelTit: {
    fontSize: 10,
    fontWeight: '600',
    color: colors.textMuted,
    letterSpacing: 1.1,
    textTransform: 'uppercase',
    marginBottom: spacing.sm,
  },
  ahorroPanelSub: {
    ...typography.small,
    color: colors.textSecondary,
    lineHeight: 20,
    marginBottom: spacing.md,
  },
  ahorroCeroBloq: {
    alignItems: 'center',
    paddingVertical: spacing.lg,
    marginBottom: spacing.sm,
  },
  ahorroCeroPct: { fontSize: 36, fontWeight: '800', color: colors.text, fontVariant: ['tabular-nums'] },
  ahorroCeroLey: { fontSize: 12, color: colors.textFaint, textAlign: 'center', marginTop: 2 },
  ahorroMetaG: { fontSize: 13, fontWeight: '600', color: colors.warning, marginTop: spacing.md },
  ahorroMens: { fontSize: 14, lineHeight: 22, marginTop: spacing.md, padding: spacing.md, borderRadius: radii.md },
  ahorroMensOk: { backgroundColor: 'rgba(125, 193, 145, 0.12)', color: colors.text, borderWidth: 1, borderColor: 'rgba(125, 193, 145, 0.35)' },
  ahorroMensTip: { backgroundColor: 'rgba(0,0,0,0.2)', color: colors.textSecondary, borderWidth: 1, borderColor: 'rgba(199, 195, 227, 0.12)' },
  chartsRow: {
    width: '100%',
    flexDirection: 'row',
    flexWrap: 'wrap',
    justifyContent: 'space-between',
    alignItems: 'flex-start',
    marginBottom: spacing.md,
    gap: spacing.md,
  },
  chartsRowStack: {
    flexDirection: 'column',
    alignItems: 'stretch',
  },
  heroCard: {
    borderRadius: radii.xl,
    padding: spacing.lg,
    marginBottom: spacing.md,
    borderWidth: 1,
    borderColor: 'rgba(255,255,255,0.08)',
    ...shadows.soft,
  },
  heroLabel: {
    fontSize: 10,
    fontWeight: '600',
    color: colors.textMuted,
    letterSpacing: 1.4,
    textTransform: 'uppercase',
    marginBottom: spacing.xs,
  },
  heroAmount: {
    fontSize: 32,
    fontWeight: '700',
    color: colors.text,
    letterSpacing: -0.5,
    maxWidth: '100%',
  },
  heroMoneda: { fontSize: 17, fontWeight: '500', color: colors.textSecondary },
  heroBolsillos: {
    ...typography.small,
    color: colors.textFaint,
    marginTop: 8,
    lineHeight: 18,
  },
  heroRow: {
    flexDirection: 'row',
    alignItems: 'stretch',
    marginTop: spacing.md,
    paddingTop: spacing.md,
    borderTopWidth: StyleSheet.hairlineWidth,
    borderTopColor: 'rgba(255,255,255,0.08)',
  },
  heroStat: { flex: 1, minWidth: 0 },
  heroStatLab: { fontSize: 10, fontWeight: '600', color: colors.textFaint, letterSpacing: 0.6, marginBottom: 4 },
  heroStatVal: { fontSize: 14, fontWeight: '600', color: colors.mint, fontVariant: ['tabular-nums'] },
  heroStatSep: { width: StyleSheet.hairlineWidth, backgroundColor: 'rgba(255,255,255,0.1)', marginHorizontal: spacing.md },
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
  btnIrExtractos: {
    flexDirection: 'row',
    alignItems: 'center',
    marginTop: spacing.md,
    paddingVertical: spacing.md,
    paddingHorizontal: spacing.sm,
    borderRadius: radii.md,
    borderWidth: 1,
    borderColor: colors.stroke,
    backgroundColor: 'rgba(0,0,0,0.2)',
  },
  btnIrExtractosTxt: {
    flex: 1,
    minWidth: 0,
    color: colors.accent,
    fontSize: 14,
    fontWeight: '600',
    marginLeft: 8,
    marginRight: 4,
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
