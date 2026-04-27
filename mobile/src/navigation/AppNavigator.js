import React from 'react';
import { Platform } from 'react-native';
import { NavigationContainer, DefaultTheme } from '@react-navigation/native';
import { createBottomTabNavigator } from '@react-navigation/bottom-tabs';
import { createNativeStackNavigator } from '@react-navigation/native-stack';
import { Ionicons } from '@expo/vector-icons';
import { colors } from '../theme';
import HomeScreen from '../screens/HomeScreen';
import GastosScreen from '../screens/GastosScreen';
import SaldoScreen from '../screens/SaldoScreen';
import MoreMenuScreen from '../screens/MoreMenuScreen';
import IngresosScreen from '../screens/IngresosScreen';
import CategoriasScreen from '../screens/CategoriasScreen';
import MetasScreen from '../screens/MetasScreen';
import PagosScreen from '../screens/PagosScreen';
import ReportesScreen from '../screens/ReportesScreen';
import AdminScreen from '../screens/AdminScreen';

const Tab = createBottomTabNavigator();
const Stack = createNativeStackNavigator();

const navTheme = {
  ...DefaultTheme,
  colors: {
    ...DefaultTheme.colors,
    primary: colors.accentBright,
    background: colors.bg,
    card: colors.headerBg,
    text: colors.text,
    border: colors.stroke,
  },
};

const stackOptions = {
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
  contentStyle: { backgroundColor: colors.bg },
};

function MoreStack() {
  return (
    <Stack.Navigator screenOptions={stackOptions}>
      <Stack.Screen name="MoreMenu" component={MoreMenuScreen} options={{ title: 'Más' }} />
      <Stack.Screen name="Ingresos" component={IngresosScreen} options={{ title: 'Ingresos' }} />
      <Stack.Screen name="Categorias" component={CategoriasScreen} options={{ title: 'Categorías' }} />
      <Stack.Screen name="Metas" component={MetasScreen} options={{ title: 'Metas' }} />
      <Stack.Screen name="PagosProgramados" component={PagosScreen} options={{ title: 'Pagos programados' }} />
      <Stack.Screen name="Reportes" component={ReportesScreen} options={{ title: 'Reportes' }} />
      <Stack.Screen name="Administrar" component={AdminScreen} options={{ title: 'Administrar' }} />
    </Stack.Navigator>
  );
}

export default function AppNavigator() {
  return (
    <NavigationContainer theme={navTheme}>
      <Tab.Navigator
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
          tabBarStyle: {
            backgroundColor: colors.tabBar,
            borderTopWidth: 1,
            borderTopColor: colors.tabBorder,
            /* Si cambias altura, actualiza TAB_BAR_SCROLL_PADDING en theme.js (scroll del contenido) */
            height: Platform.OS === 'ios' ? 88 : 68,
            paddingBottom: Platform.OS === 'ios' ? 28 : 12,
            paddingTop: 10,
          },
          tabBarActiveTintColor: colors.accentBright,
          tabBarInactiveTintColor: colors.textFaint,
          tabBarLabelStyle: { fontSize: 11, fontWeight: '600', letterSpacing: 0.2 },
          tabBarIcon: ({ color, size }) => {
            let icon = 'ellipse';
            if (route.name === 'Inicio') icon = 'home';
            if (route.name === 'Gastos') icon = 'card';
            if (route.name === 'Saldo') icon = 'wallet';
            if (route.name === 'Mas') icon = 'apps';
            return <Ionicons name={icon} size={size} color={color} />;
          },
        })}
      >
        <Tab.Screen name="Inicio" component={HomeScreen} options={{ headerShown: false }} />
        <Tab.Screen name="Gastos" component={GastosScreen} options={{ headerShown: false }} />
        <Tab.Screen name="Saldo" component={SaldoScreen} options={{ headerShown: false }} />
        <Tab.Screen name="Mas" component={MoreStack} options={{ headerShown: false }} />
      </Tab.Navigator>
    </NavigationContainer>
  );
}
