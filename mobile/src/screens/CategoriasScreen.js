import React, { useMemo, useState } from 'react';
import { View, Text, StyleSheet, TextInput, TouchableOpacity, Alert, useWindowDimensions } from 'react-native';
import { Ionicons } from '@expo/vector-icons';
import ScreenWrap from '../components/ScreenWrap';
import UICard from '../components/UICard';
import { PrimaryButton, GhostButton } from '../components/Buttons';
import { useApp } from '../context/AppContext';
import { formatearNumero, normalizarCategoria } from '../lib/finance';
import {
  CATEGORIAS_POR_DEFECTO_503020,
  CATALOGO_ICONOS_ION_CATEGORIA,
  GRUPO503020_TODOS,
  ETIQUETA_GRUPO_503020,
  GRUPO_DESEOS,
} from '../lib/categorias503020';
import { colors, spacing, radii, typography } from '../theme';

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

function construirListaIonUnica() {
  const fromDef = CATEGORIAS_POR_DEFECTO_503020.map((c) => c.iconoIon).filter(Boolean);
  return [...new Set([...fromDef, ...CATALOGO_ICONOS_ION_CATEGORIA])];
}

const LISTA_ION_TODOS = construirListaIonUnica();

export default function CategoriasScreen() {
  const { state, replaceState } = useApp();
  const { width: winW } = useWindowDimensions();
  const [nombre, setNombre] = useState('');
  const [color, setColor] = useState(PALETA_COLORES[0]);
  const [icono, setIcono] = useState(ICONOS_CATEGORIA[0]);
  /** Icono Ionicons (opcional); si hay uno, tiene prioridad visual sobre el emoji. */
  const [iconoIon, setIconoIon] = useState(null);
  const [grupo503020, setGrupo503020] = useState(GRUPO_DESEOS);
  const [limite, setLimite] = useState('');
  /** Nombre de la fila al abrir "Editar" (estable, para guardar o renombrar y migrar movimientos). */
  const [categoriaEnEdicion, setCategoriaEnEdicion] = useState(null);

  const anchoPaletas = Math.min(winW, 480);
  const margenPunto = 6;
  const columnas = 6;
  const tamColor = Math.max(32, Math.floor((anchoPaletas - 48 - columnas * margenPunto * 2) / columnas) - 2);
  const columnasIcono = 5;
  const columnasIon = 5;

  const categorias = (state.categorias || []).map(normalizarCategoria);

  const tamCeldaIon = useMemo(
    () => Math.min(52, Math.max(44, (anchoPaletas - 48) / columnasIon - 4)),
    [anchoPaletas]
  );

  function formularioLimpioNuevo() {
    setCategoriaEnEdicion(null);
    setNombre('');
    setLimite('');
    setColor(PALETA_COLORES[0]);
    setIcono(ICONOS_CATEGORIA[0]);
    setIconoIon(null);
    setGrupo503020(GRUPO_DESEOS);
  }

  function abrirEdicion(c) {
    const n = normalizarCategoria(c);
    setCategoriaEnEdicion(n.nombre);
    setNombre(n.nombre);
    setColor(n.color_hex || n.color);
    setIconoIon(n.iconoIon || null);
    setIcono(n.iconoIon && (!n.icono || !String(n.icono).trim()) ? ICONOS_CATEGORIA[0] : n.icono || ICONOS_CATEGORIA[0]);
    setGrupo503020(n.grupo503020 || GRUPO_DESEOS);
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
            return {
              nombre: n,
              color,
              color_hex: color,
              limite: l,
              icono: iconoIon ? ' ' : icono,
              iconoIon: iconoIon || undefined,
              grupo503020,
            };
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
      categorias: [
        ...(s.categorias || []),
        {
          nombre: n,
          color,
          color_hex: color,
          limite: l,
          icono: iconoIon ? ' ' : icono,
          iconoIon: iconoIon || undefined,
          grupo503020,
        },
      ],
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
    const nuevas = [...(state.categorias || [])];
    CATEGORIAS_POR_DEFECTO_503020.forEach((def) => {
      if (!nombres.includes(def.nombre.toLowerCase())) {
        nuevas.push({ ...def, limite: def.limite ?? null });
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
        <Text style={styles.lab}>Pilar 50 / 30 / 20</Text>
        <Text style={styles.hint}>
          Agrupa la categoría para el tablero de Inicio (Necesidades ~50%, Deseos ~30%, Ahorro/deuda ~20%).
        </Text>
        <View style={styles.grupoRow}>
          {GRUPO503020_TODOS.map((g) => {
            const sel = grupo503020 === g;
            return (
              <TouchableOpacity
                key={g}
                onPress={() => setGrupo503020(g)}
                activeOpacity={0.88}
                style={[styles.grupoChip, sel && styles.grupoChipSel]}
              >
                <Text style={[styles.grupoChipTxt, sel && styles.grupoChipTxtSel]} numberOfLines={2}>
                  {ETIQUETA_GRUPO_503020[g] || g}
                </Text>
              </TouchableOpacity>
            );
          })}
        </View>

        <Text style={styles.lab}>Icono (emoji)</Text>
        <Text style={styles.hint}>Opcional si usas un icono de sistema abajo.</Text>
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

        <Text style={styles.lab}>Icono (Ionicons)</Text>
        <Text style={styles.hint}>Toca para elegir; si eliges uno, sustituye al emoji en gráficos.</Text>
        <View style={styles.paletaIconos}>
          <TouchableOpacity
            onPress={() => setIconoIon(null)}
            activeOpacity={0.88}
            style={[
              styles.celdaIcono,
              { width: tamCeldaIon, minHeight: tamCeldaIon },
              !iconoIon && styles.celdaIconoSel,
            ]}
          >
            <Text style={styles.sinIonTxt}>Ninguno</Text>
          </TouchableOpacity>
          {LISTA_ION_TODOS.map((ion) => {
            const sel = iconoIon === ion;
            return (
              <TouchableOpacity
                key={ion}
                onPress={() => setIconoIon(ion)}
                activeOpacity={0.88}
                style={[
                  styles.celdaIcono,
                  { width: tamCeldaIon, minHeight: tamCeldaIon },
                  sel && styles.celdaIconoSel,
                ]}
              >
                <Ionicons name={ion} size={22} color={sel ? colors.accentGold : colors.text} />
              </TouchableOpacity>
            );
          })}
        </View>

        <View style={styles.vistaPrevia} accessibilityRole="text">
          <View style={[styles.vistaPreviaIconWrap, { backgroundColor: color }]}>
            {iconoIon ? (
              <Ionicons name={iconoIon} size={26} color="#fff" />
            ) : (
              <Text style={styles.vistaPreviaIcon}>{icono}</Text>
            )}
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
          <GhostButton
            title="Añadir categorías 50/30/20 sugeridas"
            onPress={addDefaults}
            style={{ marginTop: spacing.sm }}
          />
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
                  {c.iconoIon ? (
                    <Ionicons name={c.iconoIon} size={20} color="#fff" />
                  ) : (
                    <Text style={styles.swatchIcono}>{c.icono}</Text>
                  )}
                </View>
                <Text style={styles.catText} numberOfLines={3}>
                  {c.nombre}
                  {c.limite ? ` · lím. ${formatearNumero(parseFloat(c.limite))}` : ''}
                  {'\n'}
                  <Text style={styles.catGrupo}>
                    {ETIQUETA_GRUPO_503020[c.grupo503020] || c.grupo503020}
                  </Text>
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
  grupoRow: {
    flexDirection: 'row',
    flexWrap: 'wrap',
    gap: spacing.sm,
    marginTop: spacing.xs,
  },
  grupoChip: {
    flex: 1,
    minWidth: '28%',
    paddingVertical: spacing.sm,
    paddingHorizontal: spacing.xs,
    borderRadius: radii.md,
    borderWidth: 1,
    borderColor: colors.stroke,
    backgroundColor: 'rgba(0,0,0,0.12)',
  },
  grupoChipSel: {
    borderColor: colors.accentGold,
    backgroundColor: 'rgba(217, 180, 74, 0.14)',
  },
  grupoChipTxt: {
    fontSize: 10,
    color: colors.textSecondary,
    textAlign: 'center',
    fontWeight: '600',
    lineHeight: 14,
  },
  grupoChipTxtSel: {
    color: colors.accentGold,
  },
  sinIonTxt: {
    fontSize: 10,
    color: colors.textMuted,
    fontWeight: '700',
    textAlign: 'center',
  },
  catGrupo: {
    fontSize: 11,
    color: colors.textFaint,
    fontWeight: '600',
  },
});
