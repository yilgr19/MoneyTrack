import React, { useMemo } from 'react';
import { View, ScrollView, StyleSheet, useWindowDimensions, KeyboardAvoidingView, Platform } from 'react-native';
import { LinearGradient } from 'expo-linear-gradient';
import { useSafeAreaInsets } from 'react-native-safe-area-context';
import { spacing } from '../theme';
import { useTheme } from '../context/ThemeContext';
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
  const { colors, screenPadding, TAB_BAR_SCROLL_PADDING } = useTheme();
  const insets = useSafeAreaInsets();
  const { width } = useWindowDimensions();
  const horizontalPad = width < 360 ? spacing.md : spacing.lg;
  /** Área segura + aire; nunca dejar que contentStyle.paddingTop lo sobrescriba (antes tapaba el texto bajo el status bar en Inicio/Gastos/Saldo). */
  const topBase = includeTopInset ? spacing.lg + insets.top : spacing.md;
  const flatContent = contentStyle != null ? StyleSheet.flatten(contentStyle) : {};
  const topExtra = typeof flatContent.paddingTop === 'number' ? flatContent.paddingTop : 0;
  const { paddingTop: _ignorePt, ...contentRest } = flatContent;
  /**
   * El área útil de cada pestaña ya termina *encima* de la barra inferior; no hace falta reservar otra vez
   * la altura completa del tab bar. Ese padding extra en el contenedor con `scrollEnabled={false}` dejaba
   * una franja vacía (transparente) que en Android se dibujaba negra y cubría el final del formulario/lista.
   */
  /** Solo aire bajo el último control; el “home indicator” en iOS queda ya fuera del área útil respecto al tab bar. */
  const bottomPad = screenPadding.paddingBottom;

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

  const layoutStyles = useMemo(
    () =>
      StyleSheet.create({
        root: { flex: 1, backgroundColor: colors.bg },
        keyboardAvoid: { flex: 1 },
        baseFill: { backgroundColor: colors.bg },
        scroll: { flex: 1, backgroundColor: 'transparent' },
        content: { flexGrow: 1 },
        /** Hijos (p. ej. FlatList con flex:1) pueden repartir altura sin ScrollView. */
        fillColumn: { flex: 1, minHeight: 0 },
      }),
    [colors.bg]
  );

  const keyboardBody = scrollEnabled ? (
    <ScrollView
      style={layoutStyles.scroll}
      contentContainerStyle={[layoutStyles.content, paddingStyle, contentRest]}
      showsVerticalScrollIndicator={false}
      keyboardShouldPersistTaps="handled"
      keyboardDismissMode="on-drag"
      automaticallyAdjustKeyboardInsets={Platform.OS !== 'web'}
      {...(scrollProps || {})}
    >
      {children}
    </ScrollView>
  ) : (
    <View style={[layoutStyles.scroll, layoutStyles.fillColumn, paddingStyle, contentRest]}>
      {children}
    </View>
  );

  return (
    <View style={layoutStyles.root}>
      <View style={[StyleSheet.absoluteFill, layoutStyles.baseFill]} />
      <LinearGradient
        colors={[colors.gradTop, colors.gradMid, colors.gradBottom, colors.bg]}
        locations={[0, 0.28, 0.62, 1]}
        start={{ x: 0.15, y: 0 }}
        end={{ x: 0.85, y: 1 }}
        style={StyleSheet.absoluteFill}
      />
      {Platform.OS === 'web' ? (
        keyboardBody
      ) : Platform.OS === 'android' ? (
        /** Android: `adjustResize` en el manifest + padding con teclado; KAV+padding suele dejar franja negra. */
        keyboardBody
      ) : (
        <KeyboardAvoidingView
          style={layoutStyles.keyboardAvoid}
          behavior="padding"
          keyboardVerticalOffset={keyboardOffset}
        >
          {keyboardBody}
        </KeyboardAvoidingView>
      )}
    </View>
  );
}

export { TAB_BAR_SCROLL_PADDING } from '../theme';
