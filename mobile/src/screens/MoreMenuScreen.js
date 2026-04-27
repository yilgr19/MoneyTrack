import React from 'react';
import { View, Text, TouchableOpacity, StyleSheet } from 'react-native';
import { Ionicons } from '@expo/vector-icons';
import ScreenWrap from '../components/ScreenWrap';
import { colors, spacing, radii, typography } from '../theme';

const ITEMS = [
  { key: 'Ingresos', title: 'Ingresos', sub: 'Registrar entradas de dinero', icon: 'trending-up' },
  { key: 'Categorias', title: 'Categorías', sub: 'Organiza tus gastos', icon: 'grid' },
  { key: 'Metas', title: 'Metas', sub: 'Ahorro y objetivos', icon: 'flag' },
  { key: 'PagosProgramados', title: 'Pagos programados', sub: 'Suscripciones y recordatorios', icon: 'calendar' },
  { key: 'Reportes', title: 'Reportes', sub: 'Resumen por mes', icon: 'bar-chart' },
  { key: 'Administrar', title: 'Administrar', sub: 'Resets y diagnóstico', icon: 'settings-outline' },
];

export default function MoreMenuScreen({ navigation }) {
  return (
    <ScreenWrap contentStyle={{ paddingTop: spacing.xs }}>
      <Text style={typography.label}>Explora</Text>
      <Text style={typography.hero}>Más</Text>
      <Text style={[typography.subtitle, { marginBottom: spacing.lg }]}>Accesos rápidos</Text>

      {ITEMS.map((it) => (
        <TouchableOpacity
          key={it.key}
          style={styles.row}
          onPress={() => navigation.navigate(it.key)}
          activeOpacity={0.72}
        >
          <View style={styles.iconCircle}>
            <Ionicons name={it.icon} size={22} color={colors.accentBright} />
          </View>
          <View style={styles.textBlock}>
            <Text style={styles.title}>{it.title}</Text>
            <Text style={styles.sub}>{it.sub}</Text>
          </View>
          <View style={{ flexShrink: 0 }}>
            <Ionicons name="chevron-forward" size={22} color={colors.textFaint} />
          </View>
        </TouchableOpacity>
      ))}
    </ScreenWrap>
  );
}

const styles = StyleSheet.create({
  row: {
    flexDirection: 'row',
    alignItems: 'center',
    backgroundColor: colors.surface,
    borderRadius: radii.lg,
    padding: spacing.md,
    marginBottom: spacing.sm,
    borderWidth: 1,
    borderColor: colors.stroke,
  },
  iconCircle: {
    width: 48,
    height: 48,
    borderRadius: radii.md,
    backgroundColor: colors.surfaceHighlight,
    alignItems: 'center',
    justifyContent: 'center',
    marginRight: spacing.md,
    borderWidth: 1,
    borderColor: colors.stroke,
    flexShrink: 0,
  },
  textBlock: { flex: 1, minWidth: 0 },
  title: { color: colors.text, fontSize: 17, fontWeight: '700', letterSpacing: -0.2, flexShrink: 1 },
  sub: { color: colors.textMuted, fontSize: 13, marginTop: 3, lineHeight: 18, flexShrink: 1 },
});
