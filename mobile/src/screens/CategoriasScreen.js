import React, { useState } from 'react';
import { View, Text, StyleSheet, TextInput, TouchableOpacity, Alert } from 'react-native';
import ScreenWrap from '../components/ScreenWrap';
import UICard from '../components/UICard';
import { PrimaryButton, GhostButton } from '../components/Buttons';
import { useApp } from '../context/AppContext';
import { formatearNumero, normalizarCategoria } from '../lib/finance';
import { colors, spacing, radii, typography } from '../theme';

const CATEGORIAS_POR_DEFECTO = [
  { nombre: 'Supermercado', color: '#22c55e', icono: '🛒' },
  { nombre: 'Transporte', color: '#3b82f6', icono: '🚗' },
  { nombre: 'Hogar', color: '#f59e0b', icono: '🏠' },
  { nombre: 'Entretenimiento', color: '#4B246C', icono: '🎬' },
  { nombre: 'Restaurantes', color: '#ec4899', icono: '🍽️' },
  { nombre: 'Ropa', color: '#06b6d4', icono: '👕' },
  { nombre: 'Salud', color: '#ef4444', icono: '💊' },
  { nombre: 'Educación', color: '#6366f1', icono: '📚' },
  { nombre: 'Servicios', color: '#f97316', icono: '💡' },
  { nombre: 'Regalos', color: '#14b8a6', icono: '🎁' },
  { nombre: 'Viajes', color: '#0ea5e9', icono: '✈️' },
  { nombre: 'Tecnología', color: '#64748b', icono: '📱' },
];

export default function CategoriasScreen() {
  const { state, replaceState } = useApp();
  const [nombre, setNombre] = useState('');
  const [color, setColor] = useState('#22c55e');
  const [limite, setLimite] = useState('');

  const categorias = (state.categorias || []).map(normalizarCategoria);

  function agregar() {
    const n = nombre.trim();
    if (!n) return;
    if (categorias.some((c) => c.nombre.toLowerCase() === n.toLowerCase())) {
      Alert.alert('Duplicado', 'Ya existe esa categoría.');
      return;
    }
    const l = limite.trim() ? parseFloat(limite) : null;
    replaceState((s) => ({
      ...s,
      categorias: [...(s.categorias || []), { nombre: n, color, limite: l, icono: '📋' }],
    }));
    setNombre('');
    setLimite('');
  }

  function eliminar(index) {
    Alert.alert('Eliminar', '¿Quitar esta categoría?', [
      { text: 'Cancelar', style: 'cancel' },
      {
        text: 'Eliminar',
        style: 'destructive',
        onPress: () =>
          replaceState((s) => {
            const cats = [...(s.categorias || [])].map(normalizarCategoria);
            cats.splice(index, 1);
            return { ...s, categorias: cats };
          }),
      },
    ]);
  }

  function addDefaults() {
    const nombres = categorias.map((c) => c.nombre.toLowerCase());
    const nuevas = [...(state.categorias || [])].map(normalizarCategoria);
    CATEGORIAS_POR_DEFECTO.forEach((def) => {
      if (!nombres.includes(def.nombre.toLowerCase())) {
        nuevas.push({ ...def, limite: null });
        nombres.push(def.nombre.toLowerCase());
      }
    });
    replaceState((s) => ({ ...s, categorias: nuevas }));
  }

  return (
    <ScreenWrap includeTopInset={false} contentStyle={{ paddingTop: spacing.xs }}>
      <Text style={typography.label}>Organización</Text>
      <Text style={typography.hero}>Categorías</Text>
      <Text style={[typography.subtitle, { marginBottom: spacing.lg }]}>Etiquetas para tus gastos</Text>

      <UICard>
        <Text style={typography.label}>Nueva categoría</Text>
        <Text style={styles.lab}>Nombre</Text>
        <TextInput style={styles.input} value={nombre} onChangeText={setNombre} placeholderTextColor={colors.textFaint} />
        <Text style={styles.lab}>Color (hex)</Text>
        <TextInput
          style={styles.input}
          value={color}
          onChangeText={setColor}
          placeholderTextColor={colors.textFaint}
          autoCapitalize="none"
        />
        <Text style={styles.lab}>Límite mensual (opcional)</Text>
        <TextInput
          style={styles.input}
          value={limite}
          onChangeText={setLimite}
          keyboardType="decimal-pad"
          placeholderTextColor={colors.textFaint}
        />
        <PrimaryButton title="Agregar categoría" onPress={agregar} style={{ marginTop: spacing.md }} />
        <GhostButton title="Añadir categorías sugeridas" onPress={addDefaults} style={{ marginTop: spacing.sm }} />
      </UICard>

      <UICard style={{ marginBottom: 0 }}>
        <Text style={typography.label}>Lista</Text>
        {categorias.length === 0 ? (
          <Text style={typography.small}>Sin categorías aún.</Text>
        ) : (
          categorias.map((c, i) => (
            <View key={c.nombre + i} style={styles.row}>
              <View style={[styles.swatch, { backgroundColor: c.color }]} />
              <Text style={styles.catText}>
                {c.icono} {c.nombre}
                {c.limite ? ` · lím. ${formatearNumero(parseFloat(c.limite))}` : ''}
              </Text>
              <TouchableOpacity onPress={() => eliminar(i)} hitSlop={12}>
                <Text style={styles.del}>Quitar</Text>
              </TouchableOpacity>
            </View>
          ))
        )}
      </UICard>
    </ScreenWrap>
  );
}

const styles = StyleSheet.create({
  lab: {
    ...typography.label,
    marginTop: spacing.md,
    marginBottom: spacing.xs,
    color: colors.textMuted,
    letterSpacing: 0.8,
  },
  input: {
    borderWidth: 1,
    borderColor: colors.stroke,
    borderRadius: radii.md,
    padding: spacing.md,
    color: colors.text,
    fontSize: 16,
    backgroundColor: 'rgba(0,0,0,0.18)',
  },
  row: {
    flexDirection: 'row',
    alignItems: 'center',
    paddingVertical: spacing.md,
    borderBottomWidth: 1,
    borderBottomColor: colors.stroke,
  },
  swatch: { width: 10, height: 36, borderRadius: 5, marginRight: spacing.md },
  catText: { color: colors.textSecondary, flex: 1, minWidth: 0, fontSize: 15, flexShrink: 1 },
  del: { color: colors.danger, fontWeight: '600', fontSize: 13, flexShrink: 0, marginLeft: spacing.sm },
});
