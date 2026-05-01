import React, { useEffect, useRef } from "react";
import { Animated, Easing, StyleSheet, View } from "react-native";
import { colors } from "../theme";

export function MoneyLoadingAnimation() {
  const spin = useRef(new Animated.Value(0)).current;
  const b1 = useRef(new Animated.Value(0)).current;
  const b2 = useRef(new Animated.Value(0.33)).current;
  const b3 = useRef(new Animated.Value(0.66)).current;

  useEffect(() => {
    const coinLoop = Animated.loop(
      Animated.timing(spin, {
        toValue: 1,
        duration: 2000,
        easing: Easing.linear,
        useNativeDriver: true,
      })
    );
    coinLoop.start();

    const barWave = (v, delayMs) =>
      Animated.loop(
        Animated.sequence([
          Animated.delay(delayMs),
          Animated.timing(v, {
            toValue: 1,
            duration: 450,
            easing: Easing.inOut(Easing.quad),
            useNativeDriver: false,
          }),
          Animated.timing(v, {
            toValue: 0,
            duration: 450,
            easing: Easing.inOut(Easing.quad),
            useNativeDriver: false,
          }),
        ])
      );

    const w1 = barWave(b1, 0);
    const w2 = barWave(b2, 150);
    const w3 = barWave(b3, 300);
    w1.start();
    w2.start();
    w3.start();

    return () => {
      coinLoop.stop();
      w1.stop();
      w2.stop();
      w3.stop();
    };
  }, [spin, b1, b2, b3]);

  const rotate = spin.interpolate({
    inputRange: [0, 1],
    outputRange: ["0deg", "360deg"],
  });

  const h1 = b1.interpolate({ inputRange: [0, 1], outputRange: [10, 36] });
  const h2 = b2.interpolate({ inputRange: [0, 1], outputRange: [14, 40] });
  const h3 = b3.interpolate({ inputRange: [0, 1], outputRange: [8, 32] });

  return (
    <View style={styles.row} accessibilityLabel="Cargando">
      <Animated.View style={[styles.coinWrap, { transform: [{ rotate }] }]}>
        <View style={styles.coinOuter}>
          <View style={styles.coinInner} />
          <View style={styles.coinShine} />
        </View>
      </Animated.View>
      <View style={styles.bars}>
        <Animated.View style={[styles.bar, { height: h1, backgroundColor: colors.chartBlue }]} />
        <Animated.View style={[styles.bar, { height: h2, backgroundColor: colors.mint }]} />
        <Animated.View style={[styles.bar, { height: h3, backgroundColor: colors.accentGold }]} />
      </View>
    </View>
  );
}

const styles = StyleSheet.create({
  row: {
    flexDirection: "row",
    alignItems: "center",
    justifyContent: "center",
    marginTop: 8,
  },
  coinWrap: {
    width: 56,
    height: 56,
    alignItems: "center",
    justifyContent: "center",
    marginRight: 22,
  },
  coinOuter: {
    width: 48,
    height: 48,
    borderRadius: 24,
    backgroundColor: colors.accentGold,
    borderWidth: 3,
    borderColor: "#f5d76e",
    alignItems: "center",
    justifyContent: "center",
    overflow: "hidden",
  },
  coinInner: {
    width: 22,
    height: 22,
    borderRadius: 11,
    borderWidth: 2,
    borderColor: "rgba(62, 31, 90, 0.35)",
    backgroundColor: "rgba(255,255,255,0.15)",
  },
  coinShine: {
    position: "absolute",
    top: 6,
    right: 8,
    width: 12,
    height: 8,
    borderRadius: 4,
    backgroundColor: "rgba(255,255,255,0.35)",
    transform: [{ rotate: "-28deg" }],
  },
  bars: {
    flexDirection: "row",
    alignItems: "flex-end",
    height: 44,
  },
  bar: {
    width: 8,
    borderRadius: 4,
    opacity: 0.92,
    marginHorizontal: 3.5,
  },
});
