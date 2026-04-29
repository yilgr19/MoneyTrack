import React from 'react';
import { View, Text, StyleSheet } from 'react-native';
import { DeseosCompraBell } from './DeseosCompraBell';
import { NotificacionBell } from './NotificacionBell';
import { spacing, typography } from '../theme';

/**
 * Título de pantalla alineado con la campana de notificaciones (esquina superior derecha).
 */
export function HeaderConCampana({ label, title, subtitle }) {
  return (
    <View style={styles.wrap}>
      <View style={styles.row}>
        <View style={styles.titles}>
          {label ? <Text style={typography.label}>{label}</Text> : null}
          {title ? <Text style={typography.hero}>{title}</Text> : null}
        </View>
        <View style={styles.iconsRow}>
          <DeseosCompraBell />
          <NotificacionBell />
        </View>
      </View>
      {subtitle ? (
        <Text
          style={[
            typography.subtitle,
            { marginTop: label || title ? spacing.xs : 0, marginBottom: spacing.lg },
          ]}
        >
          {subtitle}
        </Text>
      ) : null}
    </View>
  );
}

const styles = StyleSheet.create({
  wrap: { marginBottom: 0 },
  row: { flexDirection: 'row', alignItems: 'flex-start', justifyContent: 'space-between' },
  titles: { flex: 1, minWidth: 0, paddingRight: spacing.md },
  iconsRow: { flexDirection: 'row', alignItems: 'flex-start' },
});
