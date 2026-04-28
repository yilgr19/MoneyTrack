import React from 'react';
import { View, Text, TouchableOpacity, StyleSheet } from 'react-native';
import { Ionicons } from '@expo/vector-icons';
import ScreenWrap from '../components/ScreenWrap';
import { HeaderConCampana } from '../components/HeaderConCampana';
import { colors, spacing, radii, typography, iconSemantic } from '../theme';

const ITEMS = [
  { key: 'Ingresos', title: 'Ingresos', sub: 'Registrar entradas de dinero', icon: 'trending-up' },
  {
    key: 'ExtractosTarjetas',
    title: 'Extractos de tarjeta',
    sub: 'Historial por entidad y mes',
    icon: 'receipt',
  },
  { key: 'Categorias', title: 'Categorías', sub: 'Organiza tus gastos', icon: 'grid' },
  { key: 'MisBolsillos', title: 'Mis bolsillos', sub: 'Ahorro separado de tu patrimonio', icon: 'wallet' },
  {
    key: 'AsistenteCompras',
    title: 'Asistente de compras',
    sub: 'Checklist y cupo mensual antes de comprar',
    icon: 'basket-outline',
  },
  { key: 'Metas', title: 'Metas', sub: 'Ahorro y objetivos', icon: 'flag' },
  { key: 'PagosProgramados', title: 'Pagos programados', sub: 'Suscripciones y recordatorios', icon: 'calendar' },
  { key: 'Movimientos', title: 'Movimientos', sub: 'Ingresos, gastos, tarjeta y metas', icon: 'swap-vertical' },
  { key: 'Administrar', title: 'Administrar', sub: 'Resets y diagnóstico', icon: 'settings-outline' },
];

export default function MoreMenuScreen({ navigation }) {
  return (
    <ScreenWrap contentStyle={{ paddingTop: spacing.xs }}>
      <HeaderConCampana label="Explora" title="Más" subtitle="Accesos rápidos" />

      {ITEMS.map((it) => (
        <TouchableOpacity
          key={it.key}
          style={styles.row}
          onPress={() => navigation.navigate(it.key)}
          activeOpacity={0.72}
        >
          <View
            style={[
              styles.iconCircle,
              { backgroundColor: iconSemantic.moreMenu[it.key]?.bg ?? colors.surfaceHighlight },
            ]}
          >
            <Ionicons
              name={it.icon}
              size={22}
              color={iconSemantic.moreMenu[it.key]?.fg ?? colors.accentBright}
            />
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
