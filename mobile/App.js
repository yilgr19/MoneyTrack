import React, { useEffect } from 'react';
import { View, ActivityIndicator, StyleSheet, Text, AppState, Platform } from 'react-native';
import { LinearGradient } from 'expo-linear-gradient';
import * as NavigationBar from 'expo-navigation-bar';
import * as SystemUI from 'expo-system-ui';
import { StatusBar } from 'expo-status-bar';
import { SafeAreaProvider } from 'react-native-safe-area-context';
import { AppProvider, useApp } from './src/context/AppContext';
import AppNavigator from './src/navigation/AppNavigator';
import { colors, typography } from './src/theme';

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

function Root() {
  const { ready } = useApp();
  if (!ready) {
    return (
      <View style={styles.loading}>
        <LinearGradient
          colors={[colors.gradTop, colors.gradBottom]}
          style={StyleSheet.absoluteFill}
        />
        <ActivityIndicator size="large" color={colors.accentBright} />
        <Text style={styles.loadingBrand}>MoneyTrack</Text>
        <Text style={styles.loadingSub}>Cargando…</Text>
      </View>
    );
  }
  return (
    <>
      <StatusBar style="light" />
      <AppNavigator />
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
  },
  loadingBrand: {
    ...typography.title,
    marginTop: 20,
    color: colors.text,
  },
  loadingSub: {
    ...typography.small,
    marginTop: 6,
    color: colors.textMuted,
  },
});
