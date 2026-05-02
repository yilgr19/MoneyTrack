import React, { useMemo } from 'react';
import { Text, TouchableOpacity, StyleSheet } from 'react-native';
import { LinearGradient } from 'expo-linear-gradient';
import { radii, spacing } from '../theme';
import { useTheme } from '../context/ThemeContext';

export function PrimaryButton({ title, onPress, style, disabled }) {
  const { colors, shadows } = useTheme();
  const styles = useMemo(
    () =>
      StyleSheet.create({
        touch: {
          borderRadius: radii.md,
          overflow: 'hidden',
          ...shadows.soft,
        },
        gradient: {
          paddingVertical: spacing.md,
          paddingHorizontal: spacing.lg,
          alignItems: 'center',
          justifyContent: 'center',
        },
        primaryText: {
          color: '#fff',
          fontSize: 16,
          fontWeight: '700',
          letterSpacing: 0.3,
        },
      }),
    [shadows]
  );
  return (
    <TouchableOpacity
      activeOpacity={0.88}
      onPress={onPress}
      disabled={disabled}
      style={[styles.touch, style, disabled && { opacity: 0.45 }]}
    >
      <LinearGradient
        colors={[colors.btnFrom, colors.btnTo]}
        start={{ x: 0, y: 0 }}
        end={{ x: 1, y: 1 }}
        style={styles.gradient}
      >
        <Text style={styles.primaryText}>{title}</Text>
      </LinearGradient>
    </TouchableOpacity>
  );
}

export function GhostButton({ title, onPress, style }) {
  const { colors } = useTheme();
  const styles = useMemo(
    () =>
      StyleSheet.create({
        ghost: {
          paddingVertical: spacing.sm,
          paddingHorizontal: spacing.md,
          borderRadius: radii.md,
          borderWidth: 1,
          borderColor: colors.stroke,
          backgroundColor: 'transparent',
        },
        ghostText: {
          color: colors.accentBright,
          fontWeight: '600',
          fontSize: 15,
          textAlign: 'center',
        },
      }),
    [colors]
  );
  return (
    <TouchableOpacity style={[styles.ghost, style]} onPress={onPress} activeOpacity={0.75}>
      <Text style={styles.ghostText}>{title}</Text>
    </TouchableOpacity>
  );
}
