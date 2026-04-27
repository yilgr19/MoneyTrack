import React, { useEffect, useRef, useMemo } from 'react';
import { View, Text, StyleSheet, Animated, Easing } from 'react-native';
import { LinearGradient } from 'expo-linear-gradient';
import { colors, radii, spacing, typography } from '../theme';

const MEDAL = ['🥇', '🥈', '🥉'];

function mixWithWhite(hex, t) {
  if (!hex || String(hex).charAt(0) !== '#') {
    return { from: colors.accentDeep, to: colors.accentBright };
  }
  const h = String(hex).replace('#', '');
  const r = parseInt(h.slice(0, 2), 16);
  const g = parseInt(h.slice(2, 4), 16);
  const b = parseInt(h.slice(4, 6), 16);
  if (Number.isNaN(r) || h.length < 6) {
    return { from: colors.accentDeep, to: colors.chartBlue };
  }
  const m = (c) => Math.round(c + (255 - c) * t);
  return {
    from: hex,
    to: `rgb(${m(r)},${m(g)},${m(b)})`,
  };
}

/**
 * Fila viva: barra con gradiente, medallas top-3, icono con pop y % en pastilla.
 */
export default function CategoriaGastoBarFun({
  cat,
  monto,
  pct,
  moneda,
  index,
  superadoCategoria,
  limiteCat,
  formatearNumero,
}) {
  const wAnim = useRef(new Animated.Value(0)).current;
  const popIcon = useRef(new Animated.Value(0.4)).current;
  const pulseRef = useRef(new Animated.Value(0)).current;

  const grad = useMemo(
    () =>
      superadoCategoria
        ? { from: 'rgba(248, 113, 113, 0.95)', to: 'rgba(199, 123, 136, 0.65)' }
        : mixWithWhite(cat.color, 0.5),
    [cat.color, superadoCategoria]
  );

  const hasMonto = monto > 0;
  const wPct = Math.min(100, Math.max(0, pct));
  const delay = index * 80;

  useEffect(() => {
    wAnim.setValue(0);
    if (!hasMonto) return undefined;
    const t = setTimeout(() => {
      Animated.spring(wAnim, {
        toValue: wPct,
        friction: 7,
        tension: 64,
        useNativeDriver: false,
      }).start();
    }, delay);
    return () => clearTimeout(t);
  }, [wPct, delay, hasMonto, wAnim]);

  useEffect(() => {
    popIcon.setValue(0.4);
    const t = setTimeout(() => {
      Animated.spring(popIcon, {
        toValue: 1,
        friction: 5,
        tension: 120,
        useNativeDriver: true,
      }).start();
    }, delay + 40);
    return () => clearTimeout(t);
  }, [delay, popIcon, index]);

  useEffect(() => {
    if (!hasMonto || index > 2) return undefined;
    const loop = Animated.loop(
      Animated.sequence([
        Animated.timing(pulseRef, { toValue: 1, duration: 1200, easing: Easing.inOut(Easing.ease), useNativeDriver: true }),
        Animated.timing(pulseRef, { toValue: 0, duration: 1200, easing: Easing.inOut(Easing.ease), useNativeDriver: true }),
      ])
    );
    loop.start();
    return () => loop.stop();
  }, [hasMonto, pulseRef, index]);

  const widthPct = wAnim.interpolate({ inputRange: [0, 100], outputRange: ['0%', '100%'] });
  const barPulse = pulseRef.interpolate({
    inputRange: [0, 1],
    outputRange: [0.9, 1],
  });

  return (
    <View
      style={[
        styles.wrap,
        superadoCategoria && styles.wrapAlerta,
        hasMonto && { borderLeftColor: superadoCategoria ? colors.danger : cat.color },
      ]}
    >
      <View style={styles.rowTop}>
        {index < 3 && hasMonto ? (
          <Text style={styles.medal} accessibilityLabel={`Puesto ${index + 1}`}>
            {MEDAL[index]}
          </Text>
        ) : (
          <View style={styles.medalGap} />
        )}

        <Animated.View
          style={[
            styles.iconBubble,
            { backgroundColor: (superadoCategoria ? colors.danger : cat.color) + '33' },
            { transform: [{ scale: popIcon }] },
          ]}
        >
          <Text style={styles.iconEmoji} allowFontScaling>
            {cat.icono || '•'}
          </Text>
        </Animated.View>

        <View style={styles.titular}>
          <Text
            style={[typography.body, styles.nombre, superadoCategoria && { color: colors.danger, fontWeight: '800' }]}
            numberOfLines={1}
          >
            {cat.nombre}
          </Text>
          {hasMonto && (
            <View style={styles.pctPill}>
              <LinearGradient
                colors={['rgba(199, 195, 227, 0.35)', 'rgba(167, 139, 250, 0.25)']}
                start={{ x: 0, y: 0 }}
                end={{ x: 1, y: 1 }}
                style={StyleSheet.absoluteFill}
              />
              <Text style={styles.pctTxt}>{Math.round(pct)}%</Text>
            </View>
          )}
        </View>

        <View style={styles.montoBox}>
          <Text
            style={[
              typography.monoAmount,
              { fontSize: 15, fontWeight: '800' },
              superadoCategoria && { color: colors.danger },
            ]}
            numberOfLines={1}
          >
            {formatearNumero(monto)} {moneda}
          </Text>
        </View>
      </View>

      <View style={[styles.track, superadoCategoria && styles.trackAlerta]}>
        <View style={styles.trackInner}>
          {hasMonto && (
            <Animated.View style={[styles.fillShell, { width: widthPct }]}>
              {index < 3 ? (
                <Animated.View style={{ flex: 1, opacity: barPulse }}>
                  <LinearGradient
                    colors={[grad.from, grad.to]}
                    start={{ x: 0, y: 0.5 }}
                    end={{ x: 1, y: 0.5 }}
                    style={StyleSheet.absoluteFill}
                  />
                  <LinearGradient
                    colors={['transparent', 'rgba(255,255,255,0.4)', 'transparent']}
                    start={{ x: 0, y: 0.5 }}
                    end={{ x: 1, y: 0.5 }}
                    style={styles.shine}
                  />
                </Animated.View>
              ) : (
                <View style={{ flex: 1 }}>
                  <LinearGradient
                    colors={[grad.from, grad.to]}
                    start={{ x: 0, y: 0.5 }}
                    end={{ x: 1, y: 0.5 }}
                    style={StyleSheet.absoluteFill}
                  />
                  <LinearGradient
                    colors={['transparent', 'rgba(255,255,255,0.4)', 'transparent']}
                    start={{ x: 0, y: 0.5 }}
                    end={{ x: 1, y: 0.5 }}
                    style={styles.shine}
                  />
                </View>
              )}
            </Animated.View>
          )}
        </View>
      </View>

      {superadoCategoria && (
        <View style={styles.filaAviso}>
          <Text style={styles.avisitTxt}>
            ⚡ Sobre tope: +{formatearNumero(monto - limiteCat)} {moneda} (lím. {formatearNumero(limiteCat)} {moneda})
          </Text>
        </View>
      )}
    </View>
  );
}

const styles = StyleSheet.create({
  wrap: {
    marginBottom: spacing.lg,
    borderRadius: radii.lg,
    borderLeftWidth: 4,
    borderLeftColor: 'transparent',
    backgroundColor: 'rgba(0,0,0,0.2)',
    padding: spacing.md,
    overflow: 'hidden',
  },
  wrapAlerta: {
    backgroundColor: 'rgba(199, 123, 136, 0.14)',
    borderWidth: 1,
    borderColor: 'rgba(248, 113, 113, 0.4)',
  },
  rowTop: {
    flexDirection: 'row',
    alignItems: 'center',
    marginBottom: spacing.sm,
  },
  medal: { fontSize: 22, marginRight: 6, width: 32, textAlign: 'center' },
  medalGap: { width: 32, marginRight: 6 },
  iconBubble: {
    width: 44,
    height: 44,
    borderRadius: 22,
    alignItems: 'center',
    justifyContent: 'center',
    marginRight: spacing.sm,
    borderWidth: 1,
    borderColor: 'rgba(255,255,255,0.12)',
  },
  iconEmoji: { fontSize: 22, lineHeight: 26 },
  titular: { flex: 1, minWidth: 0, flexDirection: 'row', alignItems: 'center', flexWrap: 'wrap' },
  nombre: { flexShrink: 1, fontWeight: '700', marginRight: 6 },
  pctPill: {
    borderRadius: radii.pill,
    paddingHorizontal: 8,
    paddingVertical: 3,
    overflow: 'hidden',
    borderWidth: 1,
    borderColor: 'rgba(199, 195, 227, 0.35)',
  },
  pctTxt: {
    fontSize: 11,
    fontWeight: '800',
    color: colors.accentGold,
    fontVariant: ['tabular-nums'],
  },
  montoBox: { marginLeft: 6, maxWidth: '40%' },
  track: {
    height: 18,
    borderRadius: 12,
    backgroundColor: 'rgba(255,255,255,0.07)',
    padding: 3,
    borderWidth: 1,
    borderColor: 'rgba(255,255,255,0.08)',
  },
  trackAlerta: { borderColor: 'rgba(248, 113, 113, 0.35)' },
  trackInner: { height: 12, borderRadius: 6, overflow: 'hidden' },
  fillShell: {
    height: 12,
    borderRadius: 6,
    overflow: 'hidden',
    minWidth: 4,
    shadowColor: '#fff',
    shadowOffset: { width: 0, height: 0 },
    shadowRadius: 8,
    elevation: 2,
  },
  shine: {
    ...StyleSheet.absoluteFillObject,
    opacity: 0.5,
  },
  filaAviso: { marginTop: spacing.xs },
  avisitTxt: { fontSize: 11, color: colors.danger, fontWeight: '700' },
});
