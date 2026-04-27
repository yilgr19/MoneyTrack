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
  /** Nombre de la fila al abrir "Editar" (estable, para guardar o renombrar y migrar movimientos). */
  const [categoriaEnEdicion, setCategoriaEnEdicion] = useState(null);

  const anchoPaletas = Math.min(winW, 480);
  const margenPunto = 6;
  const columnas = 6;
  const tamColor = Math.max(32, Math.floor((anchoPaletas - 48 - columnas * margenPunto * 2) / columnas) - 2);
  const columnasIcono = 5;

  const categorias = (state.categorias || []).map(normalizarCategoria);

  function formularioLimpioNuevo() {
    setCategoriaEnEdicion(null);
    setNombre('');
    setLimite('');
    setColor(PALETA_COLORES[0]);
    setIcono(ICONOS_CATEGORIA[0]);
  }

  function abrirEdicion(c) {
    const n = normalizarCategoria(c);
    setCategoriaEnEdicion(n.nombre);
    setNombre(n.nombre);
    setColor(n.color);
    setIcono(n.icono);
    setLimite(n.limite != null && n.limite !== '' ? String(n.limite) : '');
  }

  function guardarOAgregar() {
    const n = nombre.trim();
    if (!n) {
      Alert.alert('Nombre', 'Escribe un nombre para la categoría.');
      return;
    }
    const l = limite.trim() ? parseFloat(limite) : null;
    if (limite.trim() && Number.isNaN(l)) {
      Alert.alert('Límite', 'Escribe un número válido o deja el límite vacío.');
      return;
    }

    if (categoriaEnEdicion) {
      const otros = categorias.filter((c) => c.nombre !== categoriaEnEdicion);
      if (otros.some((c) => c.nombre.toLowerCase() === n.toLowerCase())) {
        Alert.alert('Duplicado', 'Ya existe otra categoría con ese nombre.');
        return;
      }
      const nombreAnterior = categoriaEnEdicion;
      replaceState((s) => {
        const raw = s.categorias || [];
        const nextCats = raw.map((row) => {
          const c = normalizarCategoria(row);
          if (c.nombre === nombreAnterior) {
            return { nombre: n, color, limite: l, icono };
          }
          return row;
        });
        const gastos = (s.gastos || []).map((g) => {
          if (g.categoria === nombreAnterior) return { ...g, categoria: n };
          return g;
        });
        const pagosProgramados = (s.pagosProgramados || []).map((p) => {
          if (p.categoria === nombreAnterior) return { ...p, categoria: n };
          return p;
        });
        return { ...s, categorias: nextCats, gastos, pagosProgramados };
      });
      formularioLimpioNuevo();
      return;
    }

    if (categorias.some((c) => c.nombre.toLowerCase() === n.toLowerCase())) {
      Alert.alert('Duplicado', 'Ya existe esa categoría.');
      return;
    }
    replaceState((s) => ({
      ...s,
      categorias: [...(s.categorias || []), { nombre: n, color, limite: l, icono }],
    }));
    setNombre('');
    setLimite('');
  }

  function eliminar(index) {
    const c = categorias[index];
    Alert.alert('Eliminar', '¿Quitar esta categoría?', [
      { text: 'Cancelar', style: 'cancel' },
      {
        text: 'Eliminar',
        style: 'destructive',
        onPress: () => {
          if (c && c.nombre === categoriaEnEdicion) formularioLimpioNuevo();
          replaceState((s) => {
            const cats = [...(s.categorias || [])].map(normalizarCategoria);
            cats.splice(index, 1);
            return { ...s, categorias: cats };
          });
        },
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
        <Text style={typography.label}>{categoriaEnEdicion ? 'Editar categoría' : 'Nueva categoría'}</Text>
        {categoriaEnEdicion ? (
          <Text style={[typography.small, { color: colors.textFaint, marginBottom: spacing.sm }]}>
            Estás editando “{categoriaEnEdicion}”. Si cambias el nombre, los gastos y pagos programados
            vinculados se actualizan a nombre nuevo.
          </Text>
        ) : null}
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
        <PrimaryButton
          title={categoriaEnEdicion ? 'Guardar cambios' : 'Agregar categoría'}
          onPress={guardarOAgregar}
          style={{ marginTop: spacing.md }}
        />
        {categoriaEnEdicion ? (
          <GhostButton title="Cancelar edición" onPress={formularioLimpioNuevo} style={{ marginTop: spacing.sm }} />
        ) : null}
        {!categoriaEnEdicion ? (
          <GhostButton title="Añadir categorías sugeridas" onPress={addDefaults} style={{ marginTop: spacing.sm }} />
        ) : null}
      </UICard>

      <UICard style={{ marginBottom: 0 }}>
        <Text style={typography.label}>Lista</Text>
        {categorias.length === 0 ? (
          <Text style={typography.small}>Sin categorías aún.</Text>
        ) : (
          categorias.map((c, i) => {
            const esEditando = c.nombre === categoriaEnEdicion;
            return (
              <View key={c.nombre + i} style={[styles.row, esEditando && styles.rowEdicion]}>
                <View style={[styles.swatch, { backgroundColor: c.color }]}>
                  <Text style={styles.swatchIcono}>{c.icono}</Text>
                </View>
                <Text style={styles.catText} numberOfLines={2}>
                  {c.nombre}
                  {c.limite ? ` · lím. ${formatearNumero(parseFloat(c.limite))}` : ''}
                </Text>
                <View style={styles.rowAcc}>
                  <TouchableOpacity onPress={() => abrirEdicion(c)} hitSlop={10}>
                    <Text style={styles.ed}>Editar</Text>
                  </TouchableOpacity>
                  <TouchableOpacity onPress={() => eliminar(i)} hitSlop={10}>
                    <Text style={styles.del}>Quitar</Text>
                  </TouchableOpacity>
                </View>
              </View>
            );
          })
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
  rowEdicion: { backgroundColor: 'rgba(217, 180, 74, 0.08)' },
  catText: { color: colors.textSecondary, flex: 1, minWidth: 0, fontSize: 15, flexShrink: 1 },
  rowAcc: { flexDirection: 'row', alignItems: 'center', flexShrink: 0, marginLeft: spacing.xs },
  ed: { color: colors.accent, fontWeight: '600', fontSize: 13, marginRight: spacing.sm },
  del: { color: colors.danger, fontWeight: '600', fontSize: 13 },
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
