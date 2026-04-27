import React, { useEffect } from 'react';
import {
  View,
  StyleSheet,
  Text,
  AppState,
  Platform,
  Image,
  useWindowDimensions,
} from 'react-native';
import { LinearGradient } from 'expo-linear-gradient';
import * as NavigationBar from 'expo-navigation-bar';
import * as SplashScreen from 'expo-splash-screen';
import * as SystemUI from 'expo-system-ui';
import { StatusBar } from 'expo-status-bar';
import { SafeAreaProvider } from 'react-native-safe-area-context';
import { AppProvider, useApp } from './src/context/AppContext';
import AppNavigator from './src/navigation/AppNavigator';
import OnboardingScreen from './src/screens/OnboardingScreen';
import { colors, spacing, typography } from './src/theme';

SplashScreen.preventAutoHideAsync().catch(() => {});

/** Android: oculta la barra de navegación del sistema (vuelve a mostrarse con gesto desde abajo). iOS no tiene esa barra. */
function useImmersiveSystemChrome() {
  useEffect(() => {
    if (Platform.OS === 'web') return;

    async function applyAndroid() {
      if (Platform.OS !== 'android') return;
      try {
        await SystemUI.setBackgroundColorAsync(colors.bg);
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

const loadingGradientColors = [colors.gradTop, colors.gradMid, colors.gradBottom, colors.bg];
const loadingGradientLocations = [0, 0.28, 0.62, 1];

function LoadingSplash() {
  const { width, height } = useWindowDimensions();
  const shortSide = Math.min(width, height);
  /** Cuadro blanco proporcional al dispositivo: ni demasiado pequeño ni desbordado */
  const box = Math.max(168, Math.min(shortSide * 0.56, width * 0.88, 320));
  const innerPad = Math.max(10, Math.round(box * 0.07));

  return (
    <View style={styles.loading}>
      <View style={[StyleSheet.absoluteFill, { backgroundColor: colors.bg }]} />
      <LinearGradient
        colors={loadingGradientColors}
        locations={loadingGradientLocations}
        start={{ x: 0.15, y: 0 }}
        end={{ x: 0.85, y: 1 }}
        style={StyleSheet.absoluteFill}
      />
      <View style={[styles.logoCard, { width: box, height: box, padding: innerPad }]}>
        <View style={styles.logoImageWrap}>
          <Image
            accessibilityLabel="MoneyTrack"
            source={require('./assets/moneytrack-logo.jpg')}
            style={styles.logoImage}
            resizeMode="contain"
          />
        </View>
      </View>
      <Text style={styles.loadingBrand}>MoneyTrack</Text>
      <Text style={styles.loadingSub}>Cargando…</Text>
    </View>
  );
}

function Root() {
  const { ready, mostrarOnboarding } = useApp();

  useEffect(() => {
    if (ready) {
      SplashScreen.hideAsync().catch(() => {});
    }
  }, [ready]);

  if (!ready) {
    return <LoadingSplash />;
  }

  return (
    <>
      <StatusBar style="light" />
      {mostrarOnboarding ? <OnboardingScreen /> : <AppNavigator />}
    </>
  );
}

export default function App() {
  useImmersiveSystemChrome();
  return (
    <SafeAreaProvider>
      <AppProvider>
        <Root />
      </AppProvider>
    </SafeAreaProvider>
  );
}

const styles = StyleSheet.create({
  loading: {
    flex: 1,
    backgroundColor: colors.bg,
    alignItems: 'center',
    justifyContent: 'center',
    paddingHorizontal: spacing.lg,
  },
  logoCard: {
    backgroundColor: '#FFFFFF',
    borderRadius: 28,
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
  logoImage: {
    width: '100%',
    height: '100%',
    backgroundColor: '#FFFFFF',
  },
  loadingBrand: {
    ...typography.title,
    marginTop: spacing.lg,
    color: colors.text,
    letterSpacing: -0.3,
  },
  loadingSub: {
    ...typography.small,
    marginTop: spacing.sm,
    color: colors.textMuted,
  },
});
