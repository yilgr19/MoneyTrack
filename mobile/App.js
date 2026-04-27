import React from 'react';
import { View, ActivityIndicator, StyleSheet, Text } from 'react-native';
import { LinearGradient } from 'expo-linear-gradient';
import { StatusBar } from 'expo-status-bar';
import { SafeAreaProvider } from 'react-native-safe-area-context';
import { AppProvider, useApp } from './src/context/AppContext';
import AppNavigator from './src/navigation/AppNavigator';
import { colors, typography } from './src/theme';

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
