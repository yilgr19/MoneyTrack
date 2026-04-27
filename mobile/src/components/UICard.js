import React from 'react';
import { View, StyleSheet } from 'react-native';
import { colors, radii, shadows, spacing } from '../theme';

export default function UICard({ children, style, accent }) {
  return (
    <View
      style={[
        styles.card,
        accent && styles.cardAccent,
        shadows.card,
        style,
      ]}
    >
      {children}
    </View>
  );
}

const styles = StyleSheet.create({
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
});
