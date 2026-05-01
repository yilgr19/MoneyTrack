import React, { useEffect, useRef } from 'react';
import { View, Text, StyleSheet, Animated, Easing } from 'react-native';
import {
  GRUPO_NECESIDADES,
  GRUPO_DESEOS,
  GRUPO_AHORRO_DEUDA,
  ETIQUETA_GRUPO_503020,
  META_FRACCION_GRUPO,
  fraccionGastoSobreIngreso,
} from '../lib/categorias503020';
import { colors, spacing, radii, typography } from '../theme';

const ORDEN = [GRUPO_NECESIDADES, GRUPO_DESEOS, GRUPO_AHORRO_DEUDA];

/**
 * Barras de gasto vs ingreso del mes con línea de meta 50/30/20.
 * Si el gasto del grupo supera la meta, la barra se muestra en rojo.
 */
export default function Salud503020Barras({ montosPorGrupo, ingresoMes, moneda, formatearNumero }) {
  const ing = Math.max(0, parseFloat(ingresoMes) || 0);

  return (
    <View style={styles.wrap}>
      {ORDEN.map((grupo, index) => (
        <FilaPilar
          key={grupo}
          grupo={grupo}
          monto={Math.max(0, parseFloat(montosPorGrupo?.[grupo]) || 0)}
          ingreso={ing}
          moneda={moneda}
          formatearNumero={formatearNumero}
          index={index}
        />
      ))}
    </View>
  );
}

function FilaPilar({ grupo, monto, ingreso, moneda, formatearNumero, index }) {
  const metaFrac = META_FRACCION_GRUPO[grupo] || 0;
  const gastoFrac = fraccionGastoSobreIngreso(monto, ingreso);
  /** Altura de barra en % del ancho del carril (0–100+). */
  const barPct = ingreso > 0 ? Math.min(100, gastoFrac * 100) : 0;
  const targetPct = metaFrac * 100;
  const superaMeta = ingreso > 0 && gastoFrac > metaFrac + 0.001;
  const anim = useRef(new Animated.Value(0)).current;

  useEffect(() => {
    anim.setValue(0);
    Animated.timing(anim, {
      toValue: barPct,
      duration: 720,
      delay: index * 90,
      easing: Easing.out(Easing.cubic),
      useNativeDriver: false,
    }).start();
  }, [barPct, index, anim]);

  const widthAnim = anim.interpolate({
    inputRange: [0, 100],
    outputRange: ['0%', '100%'],
    extrapolate: 'clamp',
  });

  return (
    <View style={styles.fila}>
      <View style={styles.filaHead}>
        <Text style={styles.filaTit} numberOfLines={2}>
          {ETIQUETA_GRUPO_503020[grupo] || grupo}
        </Text>
        <Text style={[styles.filaMonto, superaMeta && styles.filaMontoAlert]} numberOfLines={1}>
          {formatearNumero(monto)} {moneda}
          {ingreso > 0 ? (
            <Text style={styles.filaPct}> · {Math.round(gastoFrac * 100)}%</Text>
          ) : null}
        </Text>
      </View>
      <View style={styles.track}>
        <View style={styles.trackBg} />
        {/* Línea de meta */}
        <View style={[styles.metaLine, { left: `${targetPct}%` }]} />
        <View style={styles.barClip}>
          <Animated.View
            style={[
              styles.barFill,
              {
                width: widthAnim,
                backgroundColor: superaMeta ? colors.danger : colors.mint,
              },
            ]}
          />
        </View>
      </View>
      <Text style={styles.leyMeta}>
        Meta: ≤{Math.round(metaFrac * 100)}% del ingreso
        {superaMeta ? ' · por encima' : ingreso > 0 ? ' · dentro o por debajo' : ''}
      </Text>
    </View>
  );
}

const styles = StyleSheet.create({
  wrap: {},
  fila: {
    marginBottom: spacing.sm,
  },
  filaHead: {
    flexDirection: 'row',
    justifyContent: 'space-between',
    alignItems: 'flex-start',
    gap: spacing.sm,
    marginBottom: spacing.xs,
  },
  filaTit: {
    ...typography.small,
    color: colors.textSecondary,
    fontWeight: '700',
    flex: 1,
    minWidth: 0,
  },
  filaMonto: {
    ...typography.monoAmount,
    fontSize: 13,
    fontWeight: '800',
    color: colors.text,
    maxWidth: '48%',
  },
  filaMontoAlert: {
    color: colors.danger,
  },
  filaPct: {
    fontWeight: '600',
    color: colors.textMuted,
    fontSize: 12,
  },
  track: {
    height: 22,
    borderRadius: radii.sm,
    position: 'relative',
    overflow: 'hidden',
  },
  trackBg: {
    ...StyleSheet.absoluteFillObject,
    backgroundColor: 'rgba(125, 193, 145, 0.12)',
    borderRadius: radii.sm,
  },
  metaLine: {
    position: 'absolute',
    top: 0,
    bottom: 0,
    width: 3,
    marginLeft: -1.5,
    backgroundColor: 'rgba(255,255,255,0.85)',
    zIndex: 2,
    borderRadius: 1,
  },
  barClip: {
    ...StyleSheet.absoluteFillObject,
    zIndex: 1,
  },
  barFill: {
    height: '100%',
    borderRadius: radii.sm,
    opacity: 0.92,
  },
  leyMeta: {
    marginTop: 4,
    fontSize: 11,
    color: colors.textFaint,
    lineHeight: 15,
  },
});
