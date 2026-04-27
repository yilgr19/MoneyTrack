import React, { useState, useCallback } from 'react';
import { View, Text, StyleSheet, ScrollView, useWindowDimensions } from 'react-native';
import { LinearGradient } from 'expo-linear-gradient';
import { Ionicons } from '@expo/vector-icons';
import { useSafeAreaInsets } from 'react-native-safe-area-context';
import { PrimaryButton, GhostButton } from '../components/Buttons';
import { useApp } from '../context/AppContext';
import { colors, spacing, radii, typography } from '../theme';

const TOTAL_PASOS = 6;

export default function OnboardingScreen() {
  const insets = useSafeAreaInsets();
  const { width } = useWindowDimensions();
  const { completarOnboarding } = useApp();

  const [paso, setPaso] = useState(0);

  const sig = useCallback(() => {
    if (paso < TOTAL_PASOS - 1) {
      setPaso((p) => p + 1);
    } else {
      completarOnboarding();
    }
  }, [paso, completarOnboarding]);

  const atr = useCallback(() => {
    if (paso > 0) setPaso((p) => p - 1);
  }, [paso]);

  const isLast = paso === TOTAL_PASOS - 1;
  const tituloPaso = [
    'Bienvenida',
    'Saldos iniciales',
    'Cómo se usan',
    'La app en 4 pestañas',
    'Categorías y metas',
    'Listo',
  ];

  return (
    <View style={styles.root}>
      <LinearGradient
        colors={[colors.gradTop, colors.gradMid, colors.gradBottom, colors.bg]}
        locations={[0, 0.28, 0.62, 1]}
        start={{ x: 0.15, y: 0 }}
        end={{ x: 0.85, y: 1 }}
        style={StyleSheet.absoluteFill}
      />
      <ScrollView
        contentContainerStyle={[
          styles.scrollContent,
          {
            paddingTop: insets.top + spacing.md,
            paddingBottom: insets.bottom + spacing.lg,
            minHeight: '100%',
          },
        ]}
        showsVerticalScrollIndicator={false}
        keyboardShouldPersistTaps="handled"
      >
        <Text style={styles.pasoLabel}>
          {paso + 1} / {TOTAL_PASOS} · {tituloPaso[paso]}
        </Text>
        <View style={styles.dots}>
          {Array.from({ length: TOTAL_PASOS }).map((_, i) => (
            <View
              key={i}
              style={[styles.dot, i === paso && styles.dotOn, i < paso && styles.dotHecho]}
            />
          ))}
        </View>

        {paso === 0 && <PasoBienvenida ancho={width} />}
        {paso === 1 && <PasoSaldoDondeYQue />}
        {paso === 2 && <PasoPorQueSaldo />}
        {paso === 3 && <PasoTabs />}
        {paso === 4 && <PasoExtras />}
        {paso === 5 && <PasoListo />}

        <View style={styles.actions}>
          {paso > 0 && <GhostButton title="Atrás" onPress={atr} style={styles.btnGhost} />}
          <PrimaryButton
            title={isLast ? 'Ir a Saldos iniciales' : 'Continuar'}
            onPress={sig}
            style={{ marginTop: paso > 0 ? spacing.sm : 0 }}
          />
        </View>
      </ScrollView>
    </View>
  );
}

function PasoBienvenida({ ancho }) {
  const box = Math.min(200, ancho * 0.5);
  return (
    <View>
      <View style={[styles.card, { marginBottom: spacing.lg }]}>
        <Ionicons name="wallet-outline" size={56} color={colors.mint} style={{ marginBottom: spacing.md }} />
        <Text style={typography.hero}>MoneyTrack</Text>
        <Text style={[typography.subtitle, { marginTop: spacing.md, lineHeight: 22 }]}>
          Te explicamos lo esencial: para qué sirve el saldo base, qué hace cada sección y dónde encontrar
          categorías, metas e ingresos. Los números de dinero <Text style={{ fontWeight: '600' }}>no se escriben
          en este recorrido</Text>: al terminar te llevaremos a la pestaña <Text style={{ fontWeight: '600' }}>Saldo
          </Text> para que allí indiques moneda y saldos iniciales.
        </Text>
      </View>
      <View style={[styles.logoHint, { maxWidth: box + 32 }]}>
        <Text style={typography.small}>
          Siguiente: qué significa «Saldos iniciales» y dónde los cargarás.
        </Text>
      </View>
    </View>
  );
}

/** Sin campos: solo deja claro que la digitación es en la pantalla Saldo tras el tutorial. */
function PasoSaldoDondeYQue() {
  return (
    <View style={styles.card}>
      <Ionicons name="create-outline" size={48} color={colors.mint} style={{ marginBottom: spacing.md }} />
      <Text style={typography.title}>Dónde ingresarás tus saldos</Text>
      <Text style={[typography.body, { marginTop: spacing.md, lineHeight: 24 }]}>
        No escribes montos en este tutorial. Cuando cierres el recorrido, la app abrirá la pestaña <Text
        style={{ fontWeight: '600' }}>Saldo</Text>: allí elige <Text style={{ fontWeight: '600' }}>moneda
        </Text>, <Text style={{ fontWeight: '600' }}>efectivo</Text>, <Text style={{ fontWeight: '600' }}>bancos
        </Text>, billeteras, tarjeta, presupuesto y nota, según lo que uses.
      </Text>
      <Text style={[typography.small, { marginTop: spacing.lg, lineHeight: 20, color: colors.textFaint }]}>
        Tómate el tiempo: es la base para Inicio, alertas, gastos y metas. El resto del tutorial solo describe la
        app; los datos van solo en Saldo.
      </Text>
    </View>
  );
}

function PasoPorQueSaldo() {
  return (
    <View style={styles.card}>
      <Ionicons name="trending-up-outline" size={48} color={colors.chartBlue} style={{ marginBottom: spacing.md }} />
      <Text style={typography.title}>Para qué sirve lo que cargues en Saldo</Text>
      <Text style={[typography.body, { marginTop: spacing.md, lineHeight: 24 }]}>
        El total que verás en Inicio, el seguimiento de categorías, alertas y metas se alimentan de los saldos y la
        moneda que definas. Si un número no cuadra, cámbialo cuando quieras en <Text
        style={{ fontWeight: '600' }}>Saldo</Text>.
      </Text>
    </View>
  );
}

function RowIcon({ icon, text }) {
  return (
    <View style={styles.bulletRow}>
      <Ionicons name={icon} size={22} color={colors.accentBright} style={styles.bulletIcon} />
      <Text style={[typography.body, { flex: 1, minWidth: 0, lineHeight: 22 }]}>{text}</Text>
    </View>
  );
}

function PasoTabs() {
  return (
    <View style={styles.card}>
      <Text style={typography.title}>Cuatro secciones principales</Text>
      <Text style={[typography.subtitle, { marginTop: spacing.sm, marginBottom: spacing.md, lineHeight: 20 }]}>
        Toca el icono inferior para cambiar de pestaña.
      </Text>
      <RowIcon icon="home-outline" text="Inicio: resumen del mes, movimientos recientes y progreso de metas." />
      <RowIcon icon="card-outline" text="Gastos: registra cada salida, categoría, cuenta y fecha." />
      <RowIcon
        icon="wallet-outline"
        text="Saldo: aquí cargarás y actualizarás moneda, cuentas, plataformas, tarjeta, presupuesto y nota."
      />
      <RowIcon icon="apps-outline" text="Más: ingresos, categorías, metas, pagos programados, reportes y administrar." />
    </View>
  );
}

function PasoExtras() {
  return (
    <View style={styles.card}>
      <Text style={typography.title}>Luego profundizas</Text>
      <Text style={[typography.subtitle, { marginTop: spacing.sm, marginBottom: spacing.md, lineHeight: 20 }]}>
        Abajo, en Más, está casi todo lo demás.
      </Text>
      <RowIcon icon="grid-outline" text="Categorías: colores, iconos y límites para organizar gastos." />
      <RowIcon
        icon="trending-up-outline"
        text="Ingresos: añade entradas de dinero; afecta el saldo según la app."
      />
      <RowIcon icon="flag-outline" text="Metas: define objetivos y aporta desde cuentas con saldo." />
      <RowIcon icon="bar-chart-outline" text="Reportes: resúmenes y tendencias (sin Excel en móvil)." />
    </View>
  );
}

function PasoListo() {
  return (
    <View style={styles.card}>
      <Ionicons name="checkmark-circle-outline" size={56} color={colors.mint} style={{ marginBottom: spacing.md }} />
      <Text style={typography.title}>Siguiente: pestaña Saldo</Text>
      <Text style={[typography.body, { marginTop: spacing.md, lineHeight: 24 }]}>
        Al tocar <Text style={{ fontWeight: '600' }}>«Ir a Saldos iniciales»</Text> se cierra el tutorial y
        se abre <Text style={{ fontWeight: '600' }}>Saldo</Text> para que indiques allí moneda y montos. Este
        recorrido no se volverá a mostrar salvo que restablezcas todo en Administrar.
      </Text>
    </View>
  );
}

const styles = StyleSheet.create({
  root: { flex: 1, backgroundColor: colors.bg },
  scrollContent: { paddingHorizontal: spacing.lg },
  pasoLabel: {
    ...typography.label,
    textAlign: 'center',
    marginBottom: spacing.sm,
  },
  dots: { flexDirection: 'row', justifyContent: 'center', marginBottom: spacing.lg, flexWrap: 'wrap' },
  dot: {
    width: 8,
    height: 8,
    borderRadius: 4,
    backgroundColor: colors.stroke,
    marginHorizontal: 4,
    marginVertical: 2,
  },
  dotOn: { backgroundColor: colors.mint, width: 20 },
  dotHecho: { backgroundColor: colors.textMuted },
  card: {
    backgroundColor: colors.surface,
    borderRadius: radii.lg,
    padding: spacing.lg,
    borderWidth: 1,
    borderColor: colors.stroke,
    marginBottom: spacing.md,
  },
  logoHint: { alignSelf: 'center' },
  bulletRow: { flexDirection: 'row', alignItems: 'flex-start', marginBottom: spacing.md },
  bulletIcon: { marginRight: spacing.md, marginTop: 2, flexShrink: 0 },
  actions: { marginTop: spacing.lg },
  btnGhost: { marginBottom: 0 },
});
