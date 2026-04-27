import React, { useState } from 'react';
import { View, Text, StyleSheet, TextInput, TouchableOpacity, Alert, useWindowDimensions } from 'react-native';
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

/** Colores predefinidos: el usuario elige a la vista, sin introducir códigos hex. */
const PALETA_COLORES = [
  '#22c55e',
  '#3b82f6',
  '#f59e0b',
  '#4B246C',
  '#ec4899',
  '#06b6d4',
  '#ef4444',
  '#6366f1',
  '#f97316',
  '#14b8a6',
  '#0ea5e9',
  '#64748b',
  '#84cc16',
  '#a855f7',
  '#d946ef',
  '#10b981',
  '#f43f5e',
  '#0891b2',
  '#eab308',
  '#78716c',
  '#b45309',
  '#15803d',
  '#1d4ed8',
  '#7c3aed',
  '#be123c',
];

const ICONOS_CATEGORIA = [
  '📋',
  '🛒',
  '🚗',
  '🏠',
  '🎬',
  '🍽️',
  '👕',
  '💊',
  '📚',
  '💡',
  '🎁',
  '✈️',
  '📱',
  '☕',
  '🏋️',
  '🎮',
  '💰',
  '💳',
  '🏦',
  '⚡',
  '🐕',
  '👶',
  '💻',
  '🔧',
  '🎨',
  '🌮',
  '☀️',
  '🌙',
  '📦',
  '✏️',
  '🎵',
  '⚽',
  '🧾',
  '💼',
  '🏥',
  '🧹',
  '🪴',
  '🧃',
  '🧀',
  '🥗',
  '🚕',
  '🅿️',
];

export default function CategoriasScreen() {
  const { state, replaceState } = useApp();
  const { width: winW } = useWindowDimensions();
  const [nombre, setNombre] = useState('');
  const [color, setColor] = useState(PALETA_COLORES[0]);
  const [icono, setIcono] = useState(ICONOS_CATEGORIA[0]);
  const [limite, setLimite] = useState('');

  const anchoPaletas = Math.min(winW, 480);
  const margenPunto = 6;
  const columnas = 6;
  const tamColor = Math.max(32, Math.floor((anchoPaletas - 48 - columnas * margenPunto * 2) / columnas) - 2);
  const columnasIcono = 5;

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
      categorias: [...(s.categorias || []), { nombre: n, color, limite: l, icono }],
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
        <Text style={styles.lab}>Color</Text>
        <Text style={styles.hint}>Toca el color que quieras para esta categoría.</Text>
        <View style={styles.paletaColores}>
          {PALETA_COLORES.map((c, i) => {
            const sel = c === color;
            return (
              <TouchableOpacity
                key={c}
                onPress={() => setColor(c)}
                activeOpacity={0.88}
                style={[
                  styles.chipColor,
                  {
                    width: tamColor,
                    height: tamColor,
                    borderRadius: tamColor / 2,
                    backgroundColor: c,
                    margin: margenPunto,
                  },
                  sel && styles.chipColorSel,
                ]}
                accessibilityLabel={`Elegir color, opción ${i + 1} de ${PALETA_COLORES.length}`}
                accessibilityState={{ selected: sel }}
              />
            );
          })}
        </View>
        <Text style={styles.lab}>Icono</Text>
        <Text style={styles.hint}>Elige un icono para identificar la categoría en la app.</Text>
        <View style={styles.paletaIconos}>
          {ICONOS_CATEGORIA.map((ic, i) => {
            const sel = ic === icono;
            return (
              <TouchableOpacity
                key={`${i}-${ic}`}
                onPress={() => setIcono(ic)}
                activeOpacity={0.88}
                style={[
                  styles.celdaIcono,
                  {
                    width: Math.min(56, Math.max(48, (anchoPaletas - 48) / columnasIcono - 4)),
                    minHeight: 48,
                  },
                  sel && styles.celdaIconoSel,
                ]}
                accessibilityLabel={`Elegir icono, opción ${i + 1} de ${ICONOS_CATEGORIA.length}`}
                accessibilityState={{ selected: sel }}
              >
                <Text style={styles.iconoGly}>{ic}</Text>
              </TouchableOpacity>
            );
          })}
        </View>
        <View style={styles.vistaPrevia} accessibilityRole="text">
          <View style={[styles.vistaPreviaIconWrap, { backgroundColor: color }]}>
            <Text style={styles.vistaPreviaIcon}>{icono}</Text>
          </View>
          <View style={{ flex: 1, minWidth: 0 }}>
            <Text style={typography.label}>Vista previa</Text>
            <Text style={styles.hint} numberOfLines={2}>
              Así se verá en listas y gráficos.
            </Text>
          </View>
        </View>
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
              <View style={[styles.swatch, { backgroundColor: c.color }]}>
                <Text style={styles.swatchIcono}>{c.icono}</Text>
              </View>
              <Text style={styles.catText}>
                {c.nombre}
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
  swatch: {
    width: 40,
    height: 40,
    borderRadius: radii.md,
    marginRight: spacing.md,
    alignItems: 'center',
    justifyContent: 'center',
  },
  swatchIcono: { fontSize: 20 },
  catText: { color: colors.textSecondary, flex: 1, minWidth: 0, fontSize: 15, flexShrink: 1 },
  del: { color: colors.danger, fontWeight: '600', fontSize: 13, flexShrink: 0, marginLeft: spacing.sm },
  hint: {
    ...typography.small,
    color: colors.textFaint,
    marginBottom: spacing.xs,
    lineHeight: 18,
  },
  paletaColores: {
    flexDirection: 'row',
    flexWrap: 'wrap',
    marginTop: spacing.xs,
    marginLeft: -6,
    marginRight: -6,
  },
  chipColor: {
    borderWidth: 2,
    borderColor: 'transparent',
  },
  chipColorSel: {
    borderColor: colors.accentGold,
    borderWidth: 3,
  },
  paletaIconos: {
    flexDirection: 'row',
    flexWrap: 'wrap',
    marginTop: spacing.xs,
    marginLeft: -2,
    marginRight: -2,
  },
  celdaIcono: {
    alignItems: 'center',
    justifyContent: 'center',
    margin: 2,
    borderRadius: radii.md,
    borderWidth: 2,
    borderColor: 'transparent',
    backgroundColor: 'rgba(0,0,0,0.12)',
  },
  celdaIconoSel: {
    borderColor: colors.accentGold,
    backgroundColor: 'rgba(217, 180, 74, 0.12)',
  },
  iconoGly: {
    fontSize: 24,
  },
  vistaPrevia: {
    flexDirection: 'row',
    alignItems: 'center',
    marginTop: spacing.md,
    padding: spacing.md,
    borderRadius: radii.md,
    backgroundColor: 'rgba(0,0,0,0.12)',
    borderWidth: 1,
    borderColor: colors.stroke,
  },
  vistaPreviaIconWrap: {
    width: 48,
    height: 48,
    borderRadius: radii.md,
    alignItems: 'center',
    justifyContent: 'center',
    marginRight: spacing.md,
  },
  vistaPreviaIcon: {
    fontSize: 26,
  },
});
