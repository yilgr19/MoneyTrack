import React from 'react';
import { View, ScrollView, StyleSheet, useWindowDimensions, KeyboardAvoidingView, Platform } from 'react-native';
import { LinearGradient } from 'expo-linear-gradient';
import { useSafeAreaInsets } from 'react-native-safe-area-context';
import { colors, spacing, screenPadding, TAB_BAR_SCROLL_PADDING } from '../theme';
import { useKeyboardHeight } from '../hooks/useKeyboardHeight';

/**
 * Fondo con gradiente + ScrollView. Contenido con padding horizontal uniforme.
 * `scrollEnabled={false}`: solo envuelve en View (mismo relleno) — útil si un hijo ya hace scroll
 * (p. ej. FlatList) y evita anidar listas virtualizadas dentro de un ScrollView.
 */
/** includeTopInset: false cuando ya hay header de navegación encima */
export default function ScreenWrap({
  children,
  contentStyle,
  scrollProps,
  includeTopInset = true,
  scrollEnabled = true,
}) {
  const insets = useSafeAreaInsets();
  const { width } = useWindowDimensions();
  const horizontalPad = width < 360 ? spacing.md : spacing.lg;
  /** Área segura + aire; nunca dejar que contentStyle.paddingTop lo sobrescriba (antes tapaba el texto bajo el status bar en Inicio/Gastos/Saldo). */
  const topBase = includeTopInset ? spacing.lg + insets.top : spacing.md;
  const flatContent = contentStyle != null ? StyleSheet.flatten(contentStyle) : {};
  const topExtra = typeof flatContent.paddingTop === 'number' ? flatContent.paddingTop : 0;
  const { paddingTop: _ignorePt, ...contentRest } = flatContent;
  const bottomPad = screenPadding.paddingBottom + TAB_BAR_SCROLL_PADDING;

  const keyboardHeight = useKeyboardHeight();

  const paddingStyle = {
    paddingHorizontal: horizontalPad,
    paddingTop: topBase + topExtra,
    /** Android: espacio extra al final (inmersivo); iOS: KAV + automaticallyAdjustKeyboardInsets. */
    paddingBottom: bottomPad + (Platform.OS === 'android' ? keyboardHeight : 0),
  };

  /** Offset: status bar + un poco de cabecera cuando el contenido lleva header propio. */
  const keyboardOffset =
    Platform.OS === 'web'
      ? 0
      : includeTopInset
        ? insets.top + spacing.md
        : spacing.md + Math.round(TAB_BAR_SCROLL_PADDING * 0.15);
  const keyboardBody = scrollEnabled ? (
    <ScrollView
      style={styles.scroll}
      contentContainerStyle={[styles.content, paddingStyle, contentRest]}
      showsVerticalScrollIndicator={false}
      keyboardShouldPersistTaps="handled"
      keyboardDismissMode="on-drag"
      automaticallyAdjustKeyboardInsets={Platform.OS !== 'web'}
      {...(scrollProps || {})}
    >
      {children}
    </ScrollView>
  ) : (
    <View style={[styles.scroll, styles.content, styles.fillColumn, paddingStyle, contentRest]}>
      {children}
    </View>
  );

  return (
    <View style={styles.root}>
      <View style={[StyleSheet.absoluteFill, styles.baseFill]} />
      <LinearGradient
        colors={[colors.gradTop, colors.gradMid, colors.gradBottom, colors.bg]}
        locations={[0, 0.28, 0.62, 1]}
        start={{ x: 0.15, y: 0 }}
        end={{ x: 0.85, y: 1 }}
        style={StyleSheet.absoluteFill}
      />
      {Platform.OS === 'web' ? (
        keyboardBody
      ) : (
        <KeyboardAvoidingView
          style={styles.keyboardAvoid}
          behavior="padding"
          keyboardVerticalOffset={keyboardOffset}
        >
          {keyboardBody}
        </KeyboardAvoidingView>
      )}
    </View>
  );
}

const styles = StyleSheet.create({
  root: { flex: 1, backgroundColor: colors.bg },
  keyboardAvoid: { flex: 1 },
  baseFill: { backgroundColor: colors.bg },
  scroll: { flex: 1, backgroundColor: 'transparent' },
  content: { flexGrow: 1 },
  /** Hijos (p. ej. FlatList con flex:1) pueden repartir altura sin ScrollView. */
  fillColumn: { flex: 1, minHeight: 0 },
});
