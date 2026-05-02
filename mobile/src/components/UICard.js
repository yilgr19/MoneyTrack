import React, { useMemo } from 'react';
import { View, StyleSheet } from 'react-native';
import { radii, spacing } from '../theme';
import { useTheme } from '../context/ThemeContext';

export default function UICard({ children, style, accent }) {
  const { colors, shadows } = useTheme();
  const styles = useMemo(
    () =>
      StyleSheet.create({
        card: {
          backgroundColor: colors.surface,
          borderRadius: radii.lg,
          padding: spacing.lg,
          marginBottom: spacing.md,
          borderWidth: 1,
          borderColor: colors.stroke,
        },
        cardAccent: {
          borderColor: colors.strokeStrong,
          backgroundColor: colors.surfaceHighlight,
        },
      }),
    [colors]
  );
  return (
    <View style={[styles.card, accent && styles.cardAccent, shadows.card, style]}>{children}</View>
  );
}
