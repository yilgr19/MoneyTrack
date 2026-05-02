import React, { useCallback, useEffect, useMemo, useState } from 'react';
import { Platform, View } from 'react-native';
import { NavigationContainer, DefaultTheme, getFocusedRouteNameFromRoute } from '@react-navigation/native';
import { useApp } from '../context/AppContext';
import { useTheme } from '../context/ThemeContext';
import { rootNavigationRef } from './rootNavigationRef';
import { createBottomTabNavigator } from '@react-navigation/bottom-tabs';
import { createNativeStackNavigator } from '@react-navigation/native-stack';
import { Ionicons } from '@expo/vector-icons';
import HomeScreen from '../screens/HomeScreen';
import GastosScreen from '../screens/GastosScreen';
import SaldoScreen from '../screens/SaldoScreen';
import MoreMenuScreen from '../screens/MoreMenuScreen';
import IngresosScreen from '../screens/IngresosScreen';
import CategoriasScreen from '../screens/CategoriasScreen';
import MetasScreen from '../screens/MetasScreen';
import ExtractosTarjetasScreen from '../screens/ExtractosTarjetasScreen';
import MisBolsillosScreen from '../screens/MisBolsillosScreen';
import PagosScreen from '../screens/PagosScreen';
import ReportesScreen from '../screens/ReportesScreen';
import InformesMensualesScreen from '../screens/InformesMensualesScreen';
import AdminScreen from '../screens/AdminScreen';
import AsistenteComprasScreen from '../screens/AsistenteComprasScreen';
import FabRegistrarGastos from '../components/FabRegistrarGastos';

const Tab = createBottomTabNavigator();
const Stack = createNativeStackNavigator();

/** Tras el tutorial, lleva al usuario a la pestaña Saldo para que ingrese allí moneda y saldos iniciales. */
function NavegacionTrasOnboarding() {
  const { postOnboardingIrASaldo, clearPostOnboardingIrASaldo } = useApp();
  useEffect(() => {
    if (!postOnboardingIrASaldo) return;
    let n = 0;
    const max = 50;
    const t = setInterval(() => {
      if (rootNavigationRef.isReady()) {
        rootNavigationRef.navigate('Saldo');
        clearPostOnboardingIrASaldo();
        clearInterval(t);
        return;
      }
      n += 1;
      if (n >= max) {
        clearPostOnboardingIrASaldo();
        clearInterval(t);
      }
    }, 32);
    return () => clearInterval(t);
  }, [postOnboardingIrASaldo, clearPostOnboardingIrASaldo]);

  return null;
}

function tabActivaDesdeEstadoRoot(state) {
  if (!state || !Array.isArray(state.routes) || state.index == null) return 'Inicio';
  return state.routes[state.index]?.name ?? 'Inicio';
}

/**
 * El FAB (+) solo en Inicio/Saldo y en la rejilla Más. En pantallas empujadas (Bolsillos, Asistente…)
 * se oculta: en varios Android el overlay absoluto tapaba el scroll aunque el botón fuera pequeño.
 *
 * `getFocusedRouteNameFromRoute` puede ser `undefined` si el estado del stack aún no está hidratado;
 * antes caíamos en `?? 'MoreMenu'` y el FAB seguía visible encima de Asistente/Movimientos.
 */
function calcularFabVisible(navState) {
  if (!navState?.routes) return true;
  if (tabActivaDesdeEstadoRoot(navState) === 'Gastos') return false;
  const tab = navState.routes[navState.index];
  if (tab?.name !== 'Mas') return true;
  const nested = getFocusedRouteNameFromRoute(tab);
  /**
   * Si el estado del stack aún no está listo, `nested` puede ser `undefined` y antes el FAB quedaba oculto
   * en la rejilla Más (solo `nested === 'MoreMenu'`). En Inicio/Saldo no aplica.
   */
  if (nested === undefined) return true;
  return nested === 'MoreMenu';
}

export default function AppNavigator() {
  const { colors, iconSemantic } = useTheme();
  const [fabVisible, setFabVisible] = useState(true);

  const navTheme = useMemo(
    () => ({
      ...DefaultTheme,
      colors: {
        ...DefaultTheme.colors,
        primary: colors.accentBright,
        background: colors.bg,
        card: colors.headerBg,
        text: colors.text,
        border: colors.stroke,
      },
    }),
    [colors]
  );

  const stackOptions = useMemo(
    () => ({
      headerStyle: {
        backgroundColor: colors.headerBg,
        borderBottomWidth: 1,
        borderBottomColor: colors.stroke,
      },
      headerTintColor: colors.accentBright,
      headerTitleStyle: {
        fontWeight: '700',
        fontSize: 17,
        letterSpacing: -0.3,
        color: colors.text,
      },
      headerShadowVisible: false,
      contentStyle: { flex: 1, backgroundColor: colors.bg },
    }),
    [colors]
  );

  function MoreStack() {
    return (
      <Stack.Navigator screenOptions={stackOptions}>
        <Stack.Screen name="MoreMenu" component={MoreMenuScreen} options={{ headerShown: false }} />
        <Stack.Screen name="Ingresos" component={IngresosScreen} options={{ title: 'Ingresos' }} />
        <Stack.Screen
          name="ExtractosTarjetas"
          component={ExtractosTarjetasScreen}
          options={{ title: 'Extractos de tarjeta' }}
        />
        <Stack.Screen
          name="InformesMensuales"
          component={InformesMensualesScreen}
          options={{ title: 'Reportes' }}
        />
        <Stack.Screen name="MisBolsillos" component={MisBolsillosScreen} options={{ title: 'Mis bolsillos' }} />
        <Stack.Screen name="Categorias" component={CategoriasScreen} options={{ title: 'Categorías' }} />
        <Stack.Screen name="Metas" component={MetasScreen} options={{ title: 'Metas' }} />
        <Stack.Screen name="PagosProgramados" component={PagosScreen} options={{ title: 'Pagos programados' }} />
        <Stack.Screen
          name="AsistenteCompras"
          component={AsistenteComprasScreen}
          options={{ title: 'Asistente de compras' }}
        />
        <Stack.Screen name="Movimientos" component={ReportesScreen} options={{ title: 'Movimientos' }} />
        <Stack.Screen name="Administrar" component={AdminScreen} options={{ title: 'Administrar' }} />
      </Stack.Navigator>
    );
  }

  const sincronizarNavegacion = useCallback((state) => {
    if (state == null) return;
    setFabVisible(calcularFabVisible(state));
  }, []);

  return (
    <NavigationContainer
      ref={rootNavigationRef}
      theme={navTheme}
      onReady={() => {
        if (rootNavigationRef.isReady()) {
          const s = rootNavigationRef.getRootState();
          if (s) sincronizarNavegacion(s);
        }
      }}
      onStateChange={sincronizarNavegacion}
    >
      <NavegacionTrasOnboarding />
      <View style={{ flex: 1 }}>
        <Tab.Navigator
          sceneContainerStyle={{ flex: 1, backgroundColor: colors.bg }}
          screenOptions={({ route }) => ({
            headerStyle: {
              backgroundColor: colors.headerBg,
              borderBottomWidth: 1,
              borderBottomColor: colors.stroke,
            },
            headerTintColor: colors.text,
            headerTitleStyle: {
              fontWeight: '700',
              fontSize: 18,
              letterSpacing: -0.4,
              color: colors.text,
            },
            headerShadowVisible: false,
            tabBarHideOnKeyboard: true,
            tabBarStyle: {
              backgroundColor: colors.tabBar,
              borderTopWidth: 1,
              borderTopColor: colors.tabBorder,
              height: Platform.select({ ios: 88, web: 64, default: 68 }),
              paddingBottom: Platform.select({ ios: 28, web: 10, default: 12 }),
              paddingTop: 10,
            },
            tabBarActiveTintColor: colors.accentBright,
            tabBarInactiveTintColor: colors.textFaint,
            tabBarLabelStyle: { fontSize: 11, fontWeight: '600', letterSpacing: 0.2 },
            tabBarIcon: ({ color, size, focused }) => {
              let icon = 'ellipse';
              if (route.name === 'Inicio') icon = 'home';
              if (route.name === 'Gastos') icon = 'card';
              if (route.name === 'Saldo') icon = 'wallet';
              if (route.name === 'Mas') icon = 'apps';
              const tint = focused ? iconSemantic.tabActive[route.name] ?? color : color;
              return <Ionicons name={icon} size={size} color={tint} />;
            },
          })}
        >
          <Tab.Screen name="Inicio" component={HomeScreen} options={{ headerShown: false }} />
          <Tab.Screen name="Gastos" component={GastosScreen} options={{ headerShown: false }} />
          <Tab.Screen name="Saldo" component={SaldoScreen} options={{ headerShown: false }} />
          <Tab.Screen name="Mas" component={MoreStack} options={{ headerShown: false }} />
        </Tab.Navigator>
        <FabRegistrarGastos visible={fabVisible} />
      </View>
    </NavigationContainer>
  );
}
