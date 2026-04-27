import React, { useState, useCallback } from 'react';
import { View, Text, StyleSheet, ScrollView, useWindowDimensions } from 'react-native';
import { LinearGradient } from 'expo-linear-gradient';
import { Ionicons } from '@expo/vector-icons';
import { useSafeAreaInsets } from 'react-native-safe-area-context';
import { PrimaryButton, GhostButton } from '../components/Buttons';
import { useApp } from '../context/AppContext';
import { colors, spacing, radii, typography } from '../theme';

const TOTAL_PASOS = 8;

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
    'Las 4 pestañas',
    'Inicio y análisis',
    'Gastos y bolsillos',
    'Metas, ingresos y Más',
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
        {paso === 4 && <PasoInicioYAnalisis />}
        {paso === 5 && <PasoGastosYBolsillos />}
        {paso === 6 && <PasoMetasIngresosYMas />}
        {paso === 7 && <PasoListo />}

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
          <Text style={{ fontWeight: '600' }}>Inicio</Text> te muestra un resumen serio: patrimonio, ingresos y gastos
          del mes, <Text style={{ fontWeight: '600' }}>gráficos en forma de anillo</Text> con porcentajes (gastos por
          categoría e ingresos frente a gastos), <Text style={{ fontWeight: '600' }}>presupuesto mensual</Text> con barra
          de uso, saldo <Text style={{ fontWeight: '600' }}>por cuenta</Text> y, si aplica, recordatorios de tarjeta de
          crédito. En <Text style={{ fontWeight: '600' }}>Gastos</Text> registras movimientos, categorías,{' '}
          <Text style={{ fontWeight: '600' }}>pagos programados</Text> y usas <Text style={{ fontWeight: '600' }}>bolsillos
          </Text> (apartados con color). En <Text style={{ fontWeight: '600' }}>Más</Text> está ingresos, metas, reportes,{' '}
          extractos de tarjeta, notificaciones y administración.
        </Text>
        <Text style={[typography.body, { marginTop: spacing.md, lineHeight: 22, color: colors.textSecondary }]}>
          Los <Text style={{ fontWeight: '600', color: colors.text }}>montos de dinero no se escriben en este
          recorrido</Text>. Al final te llevaremos a la pestaña <Text style={{ fontWeight: '600' }}>Saldo</Text> para que
          indiques allí moneda, cuentas, bolsillos y lo demás.
        </Text>
      </View>
      <View style={[styles.logoHint, { maxWidth: box + 32 }]}>
        <Text style={typography.small}>
          Siguiente: dónde y qué cargarás en Saldo (incluido bolsillos y tarjeta).
        </Text>
      </View>
    </View>
  );
}

function PasoSaldoDondeYQue() {
  return (
    <View style={styles.card}>
      <Ionicons name="create-outline" size={48} color={colors.mint} style={{ marginBottom: spacing.md }} />
      <Text style={typography.title}>Dónde ingresarás saldos y reglas</Text>
      <Text style={[typography.body, { marginTop: spacing.md, lineHeight: 24 }]}>
        No escribes montos en este tutorial. Al cerrar el recorrido, la app abrirá <Text
        style={{ fontWeight: '600' }}>Saldo</Text>: allí eliges <Text style={{ fontWeight: '600' }}>moneda
        </Text>, <Text style={{ fontWeight: '600' }}>efectivo</Text>, <Text style={{ fontWeight: '600' }}>bancos
        </Text>, billeteras y, si aplica, <Text style={{ fontWeight: '600' }}>tarjeta de crédito
        </Text> (cupo, deuda, varias tarjetas) y <Text style={{ fontWeight: '600' }}>presupuesto mensual
        </Text> (tope de gasto del mes) y nota. También podrás crear <Text style={{ fontWeight: '600' }}>bolsillos
        </Text>: apartados con nombre y <Text style={{ fontWeight: '600' }}>color</Text>; el total en bolsillos se
        muestra en Inicio, aparte del patrimonio general.
      </Text>
      <Text style={[typography.small, { marginTop: spacing.lg, lineHeight: 20, color: colors.textFaint }]}>
        Esa base alimenta gráficos, presupuesto, alertas y metas. El resto del tutorial solo describe la app; los datos
        van en Saldo (y en los formularios de cada sección).
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
        El patrimonio y los totales de <Text style={{ fontWeight: '600' }}>Inicio</Text>, el reparto de las donas, la
        barra de <Text style={{ fontWeight: '600' }}>presupuesto</Text> y el bloque de tarjeta se basan en moneda, cuentas
        y tope que definas. Si un número no cuadra, ajústalo en <Text style={{ fontWeight: '600' }}>Saldo
        </Text> cuando quieras.
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
      <RowIcon
        icon="home-outline"
        text="Inicio: resumen con patrimonio, ingresos y gastos del mes, análisis (donas con %), presupuesto, por cuenta, recordatorios de tarjeta y metas al vuelo."
      />
      <RowIcon
        icon="card-outline"
        text="Gastos: registro de salidas con categoría, cuenta, bolsillo si aplica, fecha, y bloque de pagos programados con sus recordatorios."
      />
      <RowIcon
        icon="wallet-outline"
        text="Saldo: moneda, cuentas, bolsillos (color y saldo), plataformas, tarjeta, presupuesto tope y nota."
      />
      <RowIcon
        icon="apps-outline"
        text="Más: ingresos, categorías, metas, pagos programados, movimientos, reportes, extractos de tarjeta, notificaciones y administrar."
      />
    </View>
  );
}

function PasoInicioYAnalisis() {
  return (
    <View style={styles.card}>
      <Ionicons name="pie-chart-outline" size={48} color={colors.mint} style={{ marginBottom: spacing.md }} />
      <Text style={typography.title}>Inicio: resumen y análisis</Text>
      <Text style={[typography.subtitle, { marginTop: spacing.sm, marginBottom: spacing.md, lineHeight: 20 }]}>
        Vistas pensadas para leer de un vistazo, con porcentajes y montos.
      </Text>
      <RowIcon
        icon="analytics-outline"
        text="Bloque superior: patrimonio estimado, ingresos y gastos del mes en columnas, y bolsillos informativos si los usas."
      />
      <RowIcon
        icon="ellipse-outline"
        text="Sección Análisis: una dona de gastos por categoría (leyenda con %; centro con total de gastos del mes) y otra de reparto ingreso vs gasto (centro con ambos %)."
      />
      <RowIcon
        icon="speedometer-outline"
        text="Presupuesto mensual: tope, ingresos, gastos, flujo, disponible y barra de avance. Guía de términos integrada en la tarjeta."
      />
      <RowIcon
        icon="card-outline"
        text="Recordatorios de tarjeta (corte y pago) e información de cupo, si configuraste crédito en Saldo."
      />
    </View>
  );
}

function PasoGastosYBolsillos() {
  return (
    <View style={styles.card}>
      <Ionicons name="pricetags-outline" size={48} color={colors.accentBright} style={{ marginBottom: spacing.md }} />
      <Text style={typography.title}>Gastos y bolsillos</Text>
      <Text style={[typography.subtitle, { marginTop: spacing.sm, marginBottom: spacing.md, lineHeight: 20 }]}>
        Cada salida con contexto: categoría (con icono), fuente o destino, y bolsillos.
      </Text>
      <RowIcon
        icon="add-circle-outline"
        text="Nuevo gasto: cantidad, categoría, cuenta, fecha; opcional anotación y bolsillo—usa colores y nombres que definiste en Saldo o en Mis bolsillos."
      />
      <RowIcon
        icon="calendar-outline"
        text="Pagos programados: cargos o recordatorios recurrentes con avisos; puedes completarlos o gestionar recordatorios desde listas auxiliares."
      />
      <RowIcon
        icon="color-palette-outline"
        text="Categorías con icono semántico: en Más → Categorías ajustas colores, iconos y límites; se reflejan en listas e Inicio."
      />
    </View>
  );
}

function PasoMetasIngresosYMas() {
  return (
    <View style={styles.card}>
      <Text style={typography.title}>Metas, ingresos, reportes y Más</Text>
      <Text style={[typography.subtitle, { marginTop: spacing.sm, marginBottom: spacing.md, lineHeight: 20 }]}>
        Todo el detalle y los informes: pestaña Más o accesos desde otras pantallas.
      </Text>
      <RowIcon icon="trending-up-outline" text="Ingresos: entradas de dinero que afectan saldo y el resumen de Inicio; usa la moneda y cuentas que definiste en Saldo." />
      <RowIcon icon="flag-outline" text="Metas: objetivos y aportes desde cuentas; resumen y progreso en Inicio y en la sección de metas." />
      <RowIcon icon="swap-vertical-outline" text="Movimientos: listado unificado de ingresos, gastos, tarjeta y aportes a metas." />
      <RowIcon icon="bar-chart-outline" text="Reportes: cortes y vistas por periodo, según lo que exponga la app." />
      <RowIcon icon="receipt-outline" text="Extractos de tarjeta: desde Más, para alinear corte, pago y deuda con lo que cargaste en Saldo." />
      <RowIcon
        icon="notifications-outline"
        text="Notificaciones: avisos de pagos, recordatorios y otras alertas que puedes revisar o ajustar según el menú de la app."
      />
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
        se abre <Text style={{ fontWeight: '600' }}>Saldo</Text> para que indiques allí moneda, cuentas, bolsillos,
        crédito y presupuesto. Este recorrido no se volverá a mostrar salvo que restablezcas todo en Administrar. Luego
        explora <Text style={{ fontWeight: '600' }}>Inicio</Text> (gráficos y análisis) y <Text
        style={{ fontWeight: '600' }}>Gastos</Text> para registrar con categorías, bolsillos y programados a tu ritmo.
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
    marginHorizontal: 2,
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
