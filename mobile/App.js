import React, { useEffect, useRef } from 'react';
import {
  View,
  StyleSheet,
  Text,
  AppState,
  Platform,
  useWindowDimensions,
  Animated,
  Easing,
} from 'react-native';
import { LinearGradient } from 'expo-linear-gradient';
import * as NavigationBar from 'expo-navigation-bar';
import * as SplashScreen from 'expo-splash-screen';
import * as SystemUI from 'expo-system-ui';
import { StatusBar } from 'expo-status-bar';
import { SafeAreaProvider } from 'react-native-safe-area-context';
import { AppProvider, useApp } from './src/context/AppContext';
import { ThemeProvider } from './src/context/ThemeContext';
import { NotificacionLecturaProvider } from './src/context/NotificacionLecturaContext';
import AppNavigator from './src/navigation/AppNavigator';
import { notificacionesSistemaDisponibles } from './src/lib/notificacionesLocalesEntorno';
import OnboardingScreen from './src/screens/OnboardingScreen';
import { MoneyTrackMark } from './src/components/MoneyTrackMark';
import { spacing, typography } from './src/theme';
import { paletteOriginal } from './src/theme/colorPalettes';

SplashScreen.preventAutoHideAsync().catch(() => {});

/** Android: oculta la barra de navegación del sistema (vuelve a mostrarse con gesto desde abajo). iOS no tiene esa barra. */
function useImmersiveSystemChrome() {
  useEffect(() => {
    if (Platform.OS === 'web') return;

    async function applyAndroid() {
      if (Platform.OS !== 'android') return;
      try {
        await SystemUI.setBackgroundColorAsync(paletteOriginal.bg);
        await NavigationBar.setBehaviorAsync('overlay-swipe');
        await NavigationBar.setVisibilityAsync('hidden');
        await NavigationBar.setButtonStyleAsync('light');
      } catch (e) {
        if (__DEV__) console.warn('[MoneyTrack] System UI:', e);
      }
    }

    applyAndroid();
    const sub = AppState.addEventListener('change', (state) => {
      if (state === 'active') applyAndroid();
    });
    return () => sub.remove();
  }, []);
}

const loadingGradientColors = [
  paletteOriginal.gradTop,
  paletteOriginal.gradMid,
  paletteOriginal.gradBottom,
  paletteOriginal.bg,
];
const loadingGradientLocations = [0, 0.28, 0.62, 1];

/** Moneda dorada girando + pulso suave (tema dinero). */
function SpinningCoin() {
  const spin = useRef(new Animated.Value(0)).current;
  const pulse = useRef(new Animated.Value(1)).current;
  useEffect(() => {
    const loopSpin = Animated.loop(
      Animated.timing(spin, {
        toValue: 1,
        duration: 2600,
        easing: Easing.linear,
        useNativeDriver: true,
      })
    );
    const loopPulse = Animated.loop(
      Animated.sequence([
        Animated.timing(pulse, {
          toValue: 1.07,
          duration: 700,
          easing: Easing.inOut(Easing.ease),
          useNativeDriver: true,
        }),
        Animated.timing(pulse, {
          toValue: 1,
          duration: 700,
          easing: Easing.inOut(Easing.ease),
          useNativeDriver: true,
        }),
      ])
    );
    loopSpin.start();
    loopPulse.start();
    return () => {
      loopSpin.stop();
      loopPulse.stop();
    };
  }, [spin, pulse]);
  const rotate = spin.interpolate({
    inputRange: [0, 1],
    outputRange: ['0deg', '360deg'],
  });
  return (
    <Animated.View style={[styles.coinWrap, { transform: [{ rotate }, { scale: pulse }] }]}>
      <LinearGradient
        colors={[paletteOriginal.accentGold, '#b8922f', paletteOriginal.accentGold]}
        start={{ x: 0.15, y: 0 }}
        end={{ x: 0.85, y: 1 }}
        style={styles.coinCircle}
      >
        <Text style={styles.coinSymbol}>$</Text>
      </LinearGradient>
    </Animated.View>
  );
}

/** Barras tipo mini gráfico con altura animada. */
function MiniChartBars() {
  const h0 = useRef(new Animated.Value(14)).current;
  const h1 = useRef(new Animated.Value(22)).current;
  const h2 = useRef(new Animated.Value(18)).current;
  const h3 = useRef(new Animated.Value(26)).current;
  const heights = [h0, h1, h2, h3];
  useEffect(() => {
    const makeLoop = (h, delay, hi) =>
      Animated.loop(
        Animated.sequence([
          Animated.delay(delay),
          Animated.timing(h, {
            toValue: hi,
            duration: 480,
            easing: Easing.inOut(Easing.ease),
            useNativeDriver: false,
          }),
          Animated.timing(h, {
            toValue: 10,
            duration: 480,
            easing: Easing.inOut(Easing.ease),
            useNativeDriver: false,
          }),
        ])
      );
    const loops = [
      makeLoop(h0, 0, 34),
      makeLoop(h1, 120, 40),
      makeLoop(h2, 60, 36),
      makeLoop(h3, 180, 38),
    ];
    loops.forEach((l) => l.start());
    return () => loops.forEach((l) => l.stop());
  }, [h0, h1, h2, h3]);
  const barColors = [
    paletteOriginal.mint,
    paletteOriginal.chartBlue,
    paletteOriginal.mint,
    paletteOriginal.chartBlue,
  ];
  return (
    <View style={styles.chartRow}>
      {heights.map((h, i) => (
        <Animated.View
          key={i}
          style={[
            styles.chartBar,
            {
              height: h,
              backgroundColor: barColors[i],
            },
          ]}
        />
      ))}
    </View>
  );
}

function LoadingSplash() {
  const { width, height } = useWindowDimensions();
  const shortSide = Math.min(width, height);
  /** Logo con margen visual (PNG con transparencia se ve mejor sobre el gradiente). */
  const box = Math.max(156, Math.min(shortSide * 0.52, width * 0.82, 300));
  const innerPad = Math.max(16, Math.round(box * 0.1));
  const markSize = Math.max(88, Math.round(box - innerPad * 2));

  return (
    <View style={styles.loading}>
      <View style={[StyleSheet.absoluteFill, { backgroundColor: paletteOriginal.bg }]} />
      <LinearGradient
        colors={loadingGradientColors}
        locations={loadingGradientLocations}
        start={{ x: 0.15, y: 0 }}
        end={{ x: 0.85, y: 1 }}
        style={StyleSheet.absoluteFill}
      />
      <View
        style={[styles.logoCard, { width: box, height: box, padding: innerPad }]}
        accessible
        accessibilityLabel="MoneyTrack"
        accessibilityRole="image"
      >
        <View style={styles.logoImageWrap}>
          <MoneyTrackMark size={markSize} />
        </View>
      </View>
      <View style={styles.loadAnimRow}>
        <MiniChartBars />
        <SpinningCoin />
      </View>
      <Text style={styles.loadingBrand}>MoneyTrack</Text>
      <Text style={styles.loadingSub}>Cargando…</Text>
    </View>
  );
}

function Root() {
  const { ready, mostrarOnboarding, state } = useApp();
  const pushInicializado = useRef(false);
  const permisosNotifPreguntados = useRef(false);

  useEffect(() => {
    if (ready) {
      SplashScreen.hideAsync().catch(() => {});
    }
  }, [ready]);

  /** Tras el tutorial (o si ya estaba hecho): POST_NOTIFICATIONS / iOS alertas + vibración en canales. */
  useEffect(() => {
    if (!ready || state == null || mostrarOnboarding) return;
    if (Platform.OS === 'web') return;
    if (!notificacionesSistemaDisponibles()) return;
    if (permisosNotifPreguntados.current) return;
    permisosNotifPreguntados.current = true;
    import('./src/lib/notificacionesPermisos').then((p) =>
      p.solicitarPermisosNotificacionesAlIniciar().catch(() => {})
    );
  }, [ready, state, mostrarOnboarding]);

  useEffect(() => {
    if (!ready || state == null || mostrarOnboarding) return;
    if (!notificacionesSistemaDisponibles()) return;
    if (pushInicializado.current) return;
    pushInicializado.current = true;
    import('./src/lib/pushNotifications').then((m) => m.inicializarPushNotificaciones().catch(() => {}));
  }, [ready, state, mostrarOnboarding]);

  /** Quita el splash nativo aunque JS tarde; evita “solo logo” de Expo encima del contenido. */
  useEffect(() => {
    const t = setTimeout(() => {
      SplashScreen.hideAsync().catch(() => {});
    }, 8000);
    return () => clearTimeout(t);
  }, []);

  if (!ready || state == null) {
    return <LoadingSplash />;
  }

  return (
    <>
      <StatusBar style="light" />
      {mostrarOnboarding ? <OnboardingScreen /> : <AppNavigator />}
    </>
  );
}

function ThemeBridge({ children }) {
  const { state } = useApp();
  return <ThemeProvider temaId={state?.temaId}>{children}</ThemeProvider>;
}

export default function App() {
  useImmersiveSystemChrome();
  useEffect(() => {
    if (Platform.OS === 'web') return;
    if (!notificacionesSistemaDisponibles()) return;
    import('./src/lib/notificacionesLocalesPagosProgramados').then((m) =>
      m.registrarHandlerNotificacionesLocales()
    );
  }, []);
  return (
    <SafeAreaProvider>
      <AppProvider>
        <ThemeBridge>
          <NotificacionLecturaProvider>
            <Root />
          </NotificacionLecturaProvider>
        </ThemeBridge>
      </AppProvider>
    </SafeAreaProvider>
  );
}

const styles = StyleSheet.create({
  loading: {
    flex: 1,
    backgroundColor: paletteOriginal.bg,
    alignItems: 'center',
    justifyContent: 'center',
    paddingHorizontal: spacing.lg,
  },
  logoCard: {
    backgroundColor: 'rgba(247, 245, 251, 0.06)',
    borderRadius: 28,
    borderWidth: 1,
    borderColor: paletteOriginal.stroke,
    overflow: 'hidden',
    alignItems: 'center',
    justifyContent: 'center',
    ...Platform.select({
      ios: {
        shadowColor: '#000',
        shadowOffset: { width: 0, height: 10 },
        shadowOpacity: 0.22,
        shadowRadius: 24,
      },
      android: { elevation: 12 },
      default: {},
    }),
  },
  logoImageWrap: {
    flex: 1,
    width: '100%',
    alignItems: 'center',
    justifyContent: 'center',
  },
  loadAnimRow: {
    flexDirection: 'row',
    alignItems: 'flex-end',
    justifyContent: 'center',
    marginTop: spacing.lg + 4,
    gap: 28,
  },
  chartRow: {
    flexDirection: 'row',
    alignItems: 'flex-end',
    height: 44,
    gap: 8,
  },
  chartBar: {
    width: 10,
    borderRadius: 5,
    minHeight: 8,
  },
  coinWrap: {
    marginBottom: 2,
  },
  coinCircle: {
    width: 64,
    height: 64,
    borderRadius: 32,
    alignItems: 'center',
    justifyContent: 'center',
    borderWidth: 2,
    borderColor: 'rgba(255, 255, 255, 0.35)',
  },
  coinSymbol: {
    fontSize: 30,
    fontWeight: '900',
    color: paletteOriginal.accentDeep,
    marginTop: -2,
  },
  loadingBrand: {
    ...typography.title,
    marginTop: spacing.lg,
    color: paletteOriginal.text,
    letterSpacing: -0.3,
  },
  loadingSub: {
    ...typography.small,
    marginTop: spacing.sm,
    color: paletteOriginal.textMuted,
  },
});
