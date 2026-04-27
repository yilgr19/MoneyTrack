import React from 'react';
import { View, ScrollView, StyleSheet, useWindowDimensions } from 'react-native';
import { LinearGradient } from 'expo-linear-gradient';
import { useSafeAreaInsets } from 'react-native-safe-area-context';
import { colors, spacing, screenPadding, TAB_BAR_SCROLL_PADDING } from '../theme';

/**
 * Fondo con gradiente + ScrollView. Contenido con padding horizontal uniforme.
 */
/** includeTopInset: false cuando ya hay header de navegación encima */
export default function ScreenWrap({ children, contentStyle, scrollProps, includeTopInset = true }) {
  const insets = useSafeAreaInsets();
  const { width } = useWindowDimensions();
  const horizontalPad = width < 360 ? spacing.md : spacing.lg;
  /** Área segura + aire; nunca dejar que contentStyle.paddingTop lo sobrescriba (antes tapaba el texto bajo el status bar en Inicio/Gastos/Saldo). */
  const topBase = includeTopInset ? spacing.lg + insets.top : spacing.md;
  const flatContent = contentStyle != null ? StyleSheet.flatten(contentStyle) : {};
  const topExtra = typeof flatContent.paddingTop === 'number' ? flatContent.paddingTop : 0;
  const { paddingTop: _ignorePt, ...contentRest } = flatContent;
  const bottomPad = screenPadding.paddingBottom + TAB_BAR_SCROLL_PADDING;
  return (
    <View style={styles.root}>
      <LinearGradient
        colors={[colors.gradTop, colors.gradMid, colors.gradBottom]}
        locations={[0, 0.38, 1]}
        start={{ x: 0.2, y: 0 }}
        end={{ x: 0.9, y: 1 }}
        style={StyleSheet.absoluteFill}
      />
      <ScrollView
        style={styles.scroll}
        contentContainerStyle={[
          styles.content,
          {
            paddingHorizontal: horizontalPad,
            paddingTop: topBase + topExtra,
            paddingBottom: bottomPad,
          },
          contentRest,
        ]}
        showsVerticalScrollIndicator={false}
        keyboardShouldPersistTaps="handled"
        keyboardDismissMode="on-drag"
        {...(scrollProps || {})}
      >
        {children}
      </ScrollView>
    </View>
  );
}

const styles = StyleSheet.create({
  root: { flex: 1, backgroundColor: colors.bg },
  scroll: { flex: 1 },
  content: { flexGrow: 1 },
});
