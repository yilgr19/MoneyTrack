import React, { useMemo, useState, useEffect, useLayoutEffect, useRef } from 'react';
import { View, Text, StyleSheet, Animated, Easing } from 'react-native';
import Svg, { G, Path } from 'react-native-svg';
import { colors, spacing } from '../../theme';
import { ringSlicePath, fullDonutPaths } from './donutPaths';

const AnimatedPath = Animated.createAnimatedComponent(Path);

function fmtPorcentaje(frac) {
  if (frac == null || Number.isNaN(frac) || frac < 0) return '0%';
  const p = Math.min(100, Math.max(0, frac * 100));
  if (p < 10) return `${p.toFixed(1).replace('.', ',')}%`;
  return `${Math.round(p)}%`;
}

function buildSlicePaths(segments, total, cx, cy, rIn, rOut) {
  if (!segments?.length || total <= 0) return [];
  let a = 0;
  const paths = [];
  for (let i = 0; i < segments.length; i += 1) {
    const v = parseFloat(segments[i].value) || 0;
    if (v <= 0) continue;
    const sweep = (v / total) * 360;
    if (sweep < 0.1) continue;
    const col = segments[i].color;
    const a0 = a;
    const a1 = a + sweep;
    a = a1;
    if (sweep >= 359) {
      fullDonutPaths(cx, cy, rIn, rOut).forEach((d, j) => {
        if (d) paths.push({ d, color: col, key: `s-${i}-h${j}` });
      });
    } else {
      paths.push({
        d: ringSlicePath(cx, cy, rIn, rOut, a0, a1 - 0.01),
        color: col,
        key: `s-${i}`,
      });
    }
  }
  return paths;
}

/**
 * @param {Array<{ value: number, color: string, label: string }>} segments
 * @param {string} [title]
 * @param {string} [emptyHint]
 * @param {number} [size=150]
 * @param {string} [centerLine1]
 * @param {string} [centerLine2]
 */
export default function DonutChart({
  segments,
  title,
  emptyHint = 'Sin datos',
  size = 150,
  centerLine1,
  centerLine2,
}) {
  const total = (segments || []).reduce((s, x) => s + (parseFloat(x.value) || 0), 0);
  const cx = size / 2;
  const cy = size / 2;
  const rOut = size * 0.38;
  const rIn = size * 0.29;
  const track = 'rgba(255, 255, 255, 0.05)';

  const animKey = useMemo(
    () =>
      `${total}|${(segments || [])
        .map((s) => `${s.label}:${Number(s.value)}`)
        .join(';')}|${size}`,
    [segments, total, size]
  );

  const slicePaths = useMemo(
    () => buildSlicePaths(segments, total, cx, cy, rIn, rOut),
    [segments, total, cx, cy, rIn, rOut]
  );

  const entry = useRef(new Animated.Value(0)).current;
  const [pathAnims, setPathAnims] = useState(null);

  useLayoutEffect(() => {
    entry.setValue(0);
    if (slicePaths.length === 0) {
      setPathAnims(null);
    } else {
      setPathAnims(slicePaths.map(() => new Animated.Value(0)));
    }
  }, [animKey, slicePaths.length, entry]);

  useEffect(() => {
    if (slicePaths.length === 0) {
      Animated.spring(entry, {
        toValue: 1,
        friction: 7,
        tension: 78,
        useNativeDriver: true,
      }).start();
      return undefined;
    }
    if (!pathAnims || pathAnims.length !== slicePaths.length) return undefined;

    const springEntrada = Animated.spring(entry, {
      toValue: 1,
      friction: 7,
      tension: 78,
      useNativeDriver: true,
    });

    // fillOpacity en Path SVG: driver JS (no mezclar con spring en un solo parallel)
    const segmentAnims = pathAnims.map((v) =>
      Animated.timing(v, {
        toValue: 1,
        duration: 400,
        easing: Easing.out(Easing.cubic),
        useNativeDriver: false,
      })
    );
    const stagger = Animated.stagger(48, segmentAnims);

    springEntrada.start();
    stagger.start();
    return () => {
      entry.stopAnimation();
      (pathAnims || []).forEach((v) => v.stopAnimation());
    };
  }, [animKey, pathAnims, slicePaths.length, entry]);

  const filasLey = useMemo(() => {
    if (!total || total <= 0) return [];
    return (segments || [])
      .map((s) => {
        const v = parseFloat(s.value) || 0;
        if (v <= 0) return null;
        return {
          key: s.label,
          label: s.label,
          color: s.color,
          pct: fmtPorcentaje(v / total),
        };
      })
      .filter(Boolean)
      .slice(0, 5);
  }, [segments, total]);

  const leyOpacity = entry.interpolate({
    inputRange: [0, 0.4, 0.75, 1],
    outputRange: [0, 0, 0.45, 1],
  });

  const centerOpacity = entry.interpolate({
    inputRange: [0, 0.45, 1],
    outputRange: [0, 0.15, 1],
  });

  const renderTrack = () => (
    <G>
      {fullDonutPaths(cx, cy, rIn, rOut).map((d, j) => (
        <Path key={`tr-${j}`} d={d} fill={track} />
      ))}
    </G>
  );

  const ringScale = entry.interpolate({
    inputRange: [0, 1],
    outputRange: [0.8, 1],
  });
  const ringRotate = entry.interpolate({
    inputRange: [0, 1],
    outputRange: ['-15deg', '0deg'],
  });

  const hasPathAnims = pathAnims && pathAnims.length === slicePaths.length;
  const chartBody = hasPathAnims ? (
    <G>
      {slicePaths.map((p, i) => {
        const v = pathAnims[i];
        const fillOp = v.interpolate({
          inputRange: [0, 1],
          outputRange: [0, 0.88],
        });
        return <AnimatedPath key={p.key} d={p.d} fill={p.color} fillOpacity={fillOp} />;
      })}
    </G>
  ) : (
    <G>
      {slicePaths.map((p) => (
        <Path key={p.key} d={p.d} fill={p.color} fillOpacity={0.88} />
      ))}
    </G>
  );

  const animWrap = (children) => (
    <Animated.View
      style={{
        width: size,
        height: size,
        alignSelf: 'center',
        opacity: entry,
        transform: [{ scale: ringScale }, { rotate: ringRotate }],
      }}
    >
      {children}
    </Animated.View>
  );

  if (!segments || segments.length === 0 || total <= 0) {
    return (
      <View style={styles.card}>
        {!!title && <Text style={styles.tit}>{title}</Text>}
        <View style={{ width: size, height: size, alignSelf: 'center', position: 'relative' }}>
          {animWrap(
            <Svg width={size} height={size} viewBox={`0 0 ${size} ${size}`}>
              {renderTrack()}
            </Svg>
          )}
        </View>
        <Text style={styles.hint}>{emptyHint}</Text>
      </View>
    );
  }

  if (slicePaths.length === 0) {
    return (
      <View style={styles.card}>
        {!!title && <Text style={styles.tit}>{title}</Text>}
        <View style={{ width: size, height: size, alignSelf: 'center' }}>
          {animWrap(
            <Svg width={size} height={size} viewBox={`0 0 ${size} ${size}`}>
              {renderTrack()}
            </Svg>
          )}
        </View>
        <Text style={styles.hint}>{emptyHint}</Text>
      </View>
    );
  }

  return (
    <View style={styles.card}>
      {!!title && <Text style={styles.tit}>{title}</Text>}
      <View style={{ width: size, height: size, alignSelf: 'center', position: 'relative' }}>
        {animWrap(
          <Svg width={size} height={size} viewBox={`0 0 ${size} ${size}`} style={{ alignSelf: 'center' }}>
            {chartBody}
          </Svg>
        )}
        {(centerLine1 || centerLine2) && (
          <Animated.View style={[styles.centerBox, { opacity: centerOpacity }]} pointerEvents="none">
            {!!centerLine1 && (
              <Text style={styles.center1} numberOfLines={1}>
                {centerLine1}
              </Text>
            )}
            {!!centerLine2 && (
              <Text style={styles.center2} numberOfLines={2}>
                {centerLine2}
              </Text>
            )}
          </Animated.View>
        )}
      </View>
      <Animated.View style={[styles.leyend, { opacity: leyOpacity }]}>
        {filasLey.map((f) => (
          <View key={f.key} style={styles.leyFila}>
            <View style={styles.leyFilaL}>
              <View style={[styles.leyc, { backgroundColor: f.color }]} />
              <Text style={styles.leyLabel} numberOfLines={1}>
                {f.label}
              </Text>
            </View>
            <Text style={styles.leyPct}>{f.pct}</Text>
          </View>
        ))}
      </Animated.View>
    </View>
  );
}

const styles = StyleSheet.create({
  card: {
    alignItems: 'stretch',
    width: '100%',
    minWidth: 160,
  },
  tit: {
    fontSize: 10,
    fontWeight: '600',
    color: colors.textMuted,
    letterSpacing: 1.2,
    textTransform: 'uppercase',
    marginBottom: spacing.md,
    textAlign: 'left',
  },
  hint: {
    fontSize: 12,
    color: colors.textFaint,
    textAlign: 'center',
    marginTop: spacing.sm,
  },
  centerBox: {
    ...StyleSheet.absoluteFillObject,
    justifyContent: 'center',
    alignItems: 'center',
    paddingHorizontal: 8,
  },
  center1: {
    fontSize: 12,
    fontWeight: '600',
    color: colors.text,
    textAlign: 'center',
    fontVariant: ['tabular-nums'],
  },
  center2: {
    fontSize: 9,
    fontWeight: '500',
    color: colors.textFaint,
    textAlign: 'center',
    marginTop: 3,
    letterSpacing: 0.2,
  },
  leyend: { marginTop: spacing.md, alignSelf: 'stretch' },
  leyFila: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    marginBottom: 8,
  },
  leyFilaL: { flexDirection: 'row', alignItems: 'center', flex: 1, minWidth: 0, marginRight: spacing.sm },
  leyc: { width: 6, height: 6, borderRadius: 3, marginRight: 8, opacity: 0.95 },
  leyLabel: {
    fontSize: 12,
    color: colors.textSecondary,
    flex: 1,
    fontWeight: '500',
  },
  leyPct: {
    fontSize: 12,
    fontWeight: '600',
    color: colors.text,
    fontVariant: ['tabular-nums'],
  },
});
