import React, { useState, useCallback } from 'react';
import { View, Text, StyleSheet, ScrollView, useWindowDimensions } from 'react-native';
import { LinearGradient } from 'expo-linear-gradient';
import { Ionicons } from '@expo/vector-icons';
import { useSafeAreaInsets } from 'react-native-safe-area-context';
import { PrimaryButton, GhostButton } from '../components/Buttons';
import { useApp } from '../context/AppContext';
import { colors, spacing, radii, typography } from '../theme';

const TOTAL_PASOS = 9;

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
    'Gastos y recibos',
    'Más: informes y asistente',
    'Botón + y campanas',
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
        {paso === 7 && <PasoFabYCampanas />}
        {paso === 8 && <PasoListo />}

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
          <Text style={{ fontWeight: '600' }}>Inicio</Text> concentra el mes: patrimonio estimado, ingresos y gastos,{' '}
          <Text style={{ fontWeight: '600' }}>gráficos en anillo</Text> (gastos por categoría, ingreso vs gasto), una
          tarjeta de <Text style={{ fontWeight: '600' }}>ahorro</Text> que une bolsillos y metas frente al ingreso del mes
          (con una guía amable ~30 %), <Text style={{ fontWeight: '600' }}>presupuesto mensual</Text> con barra, bloque{' '}
          <Text style={{ fontWeight: '600' }}>por cuenta</Text> (y tarjeta con cupo aparte si lo configuraste), recordatorios
          de crédito y, si tienes pendientes, un acceso a la <Text style={{ fontWeight: '600' }}>lista súper</Text> del
          asistente. En <Text style={{ fontWeight: '600' }}>Gastos</Text> anotas salidas a mano o con{' '}
          <Text style={{ fontWeight: '600' }}>escáner de recibo</Text> (OCR que tú revisas antes de guardar), categorías,{' '}
          <Text style={{ fontWeight: '600' }}>pagos programados</Text> y <Text style={{ fontWeight: '600' }}>bolsillos
          </Text>. En <Text style={{ fontWeight: '600' }}>Más</Text>: ingresos, metas,{' '}
          <Text style={{ fontWeight: '600' }}>reportes por mes</Text> (informe con gráficos y detalle), extractos de
          tarjeta, <Text style={{ fontWeight: '600' }}>asistente de compras</Text> (lista, deseos, cupo), movimientos,
          notificaciones y administrar.
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
        text="Inicio: patrimonio e ingresos/gastos del mes, donas de análisis, tarjeta de ahorro (bolsillos + metas vs ingreso), presupuesto, por cuenta, recordatorios de tarjeta, contribuciones a metas en el pie de la card por cuenta y atajo a la lista súper si hay pendientes."
      />
      <RowIcon
        icon="card-outline"
        text="Gastos: registro manual o escaneando recibo (cámara y OCR sugerido; siempre confirmas antes de guardar), categoría, cuenta, bolsillo, fecha, nota; pagos programados con recordatorios."
      />
      <RowIcon
        icon="wallet-outline"
        text="Saldo: moneda, cuentas, bolsillos (color y saldo), plataformas, tarjeta, presupuesto tope y nota."
      />
      <RowIcon
        icon="apps-outline"
        text="Más: ingresos, categorías, mis bolsillos, metas, pagos programados, asistente de compras (lista súper y deseos), movimientos, Reportes (informe mensual con gráficos y detalle por mes), extractos de tarjeta y administrar."
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
        text="Bloque superior: patrimonio estimado (sin contar cupo de tarjeta como patrimonio propio), ingresos y gastos del mes, aviso si el cupo va aparte, y total en bolsillos cuando aplica (ese dinero no suma al patrimonio mostrado)."
      />
      <RowIcon
        icon="wallet-outline"
        text="Por cuenta: participación de cada cuenta en el patrimonio y, si configuraste TC, un bloque aparte con cupo disponible; al pie, resumen del mes e importe aportado a metas si hubo movimientos."
      />
      <RowIcon
        icon="ellipse-outline"
        text="Análisis: dona de gastos por categoría (leyenda con % y total al centro) y dona ingreso vs gasto; debajo, mensajes rotativos que explican bolsillos, metas y la guía del ~30 % del ingreso (referencia, no castigo)."
      />
      <RowIcon
        icon="pie-chart-outline"
        text="Ahorro: bolsillos · metas — dona que muestra cuánto llevas apartado frente al ingreso del mes y frases de ánimo; complementa las metas que ves en Más y el bloque de bolsillos en Saldo."
      />
      <RowIcon
        icon="speedometer-outline"
        text="Presupuesto mensual: tope, ingresos, gastos, flujo, disponible y barra de avance; en pantallas con encabezado completo hay acceso al medidor/alertas junto a la campana."
      />
      <RowIcon
        icon="card-outline"
        text="Recordatorios de tarjeta (corte y pago) e información de cupo cuando cargaste crédito en Saldo."
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
        Manual o con recibo: la app sugiere datos, pero tú confirmas categoría, cuenta y monto.
      </Text>
      <RowIcon
        icon="camera-outline"
        text="Escáner de recibo: desde Gastos, captura el ticket; se intenta leer total, fecha, comercio e ítems (según el texto). Revisa siempre antes de guardar; en algunos entornos el OCR nativo requiere la app instalada (no solo el navegador)."
      />
      <RowIcon
        icon="add-circle-outline"
        text="Nuevo gasto: cantidad, categoría, cuenta, fecha; nota opcional y bolsillo—colores y nombres desde Saldo o Mis bolsillos."
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
      <Text style={typography.title}>Más: ingresos, metas, informes y asistente</Text>
      <Text style={[typography.subtitle, { marginTop: spacing.sm, marginBottom: spacing.md, lineHeight: 20 }]}>
        Desde la rejilla Más entras al detalle; varias pantallas comparten iconos de aviso arriba a la derecha.
      </Text>
      <RowIcon
        icon="trending-up-outline"
        text="Ingresos: entradas que actualizan saldos y el resumen de Inicio; respeta moneda y cuentas de Saldo."
      />
      <RowIcon
        icon="flag-outline"
        text="Metas: objetivos, aportes y seguimiento; las contribuciones del mes pueden verse en Inicio (card por cuenta y análisis de ahorro)."
      />
      <RowIcon
        icon="wallet-outline"
        text="Mis bolsillos: apartados con color; mueven dinero fuera de la caja visible del patrimonio en Inicio hasta que los devuelves."
      />
      <RowIcon
        icon="basket-outline"
        text="Asistente de compras: lista súper con checklist, deseos y cupo mensual de referencia antes de comprar; Inicio puede mostrar un acceso rápido si hay ítems pendientes en la súper."
      />
      <RowIcon
        icon="pie-chart-outline"
        text="Reportes (informe mensual): elige mes, ve totales, dona por categorías, barras de gasto, comparativas y movimientos del periodo en un solo informe."
      />
      <RowIcon
        icon="swap-vertical-outline"
        text="Movimientos: historial unificado de ingresos, gastos, tarjeta y aportes a metas."
      />
      <RowIcon
        icon="receipt-outline"
        text="Extractos de tarjeta: concilia corte, pago y cargos con lo definido en Saldo."
      />
      <RowIcon
        icon="settings-outline"
        text="Administrar: estado del proyecto, reset de saldos y gastos, o borrado completo (incluye moneda y categorías; al arrancar de nuevo verás este tutorial). Importar/exportar Excel está en la versión web."
      />
    </View>
  );
}

function PasoFabYCampanas() {
  return (
    <View style={styles.card}>
      <Ionicons name="flash-outline" size={48} color={colors.accentBright} style={{ marginBottom: spacing.md }} />
      <Text style={typography.title}>Botón + y campanas</Text>
      <Text style={[typography.subtitle, { marginTop: spacing.sm, marginBottom: spacing.md, lineHeight: 20 }]}>
        Atajos globales en Inicio, Saldo y la pantalla principal de Más (no en todas las subpantallas).
      </Text>
      <RowIcon
        icon="add-circle-outline"
        text="Botón flotante +: abre acciones rápidas — ir a registrar gasto o abrir el Asistente de compras sin buscar en el menú."
      />
      <RowIcon
        icon="notifications-outline"
        text="Campana de notificaciones: pagos y cortes de tarjeta, recordatorios de pagos programados, avisos de presupuesto o categoría, y recordatorios del asistente (p. ej. lista súper) según tu actividad."
      />
      <RowIcon
        icon="speedometer-outline"
        text="Icono de presupuesto (cuando aparece): acceso al estado del tope del mes y alertas relacionadas."
      />
      <RowIcon
        icon="heart-outline"
        text="Lista de deseos / asistente (icono cuando aplica): atajo a pendientes o deseos del asistente sin salir de la pantalla."
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
        Al tocar <Text style={{ fontWeight: '600' }}>«Ir a Saldos iniciales»</Text> se cierra el tutorial y se abre{' '}
        <Text style={{ fontWeight: '600' }}>Saldo</Text> para moneda, cuentas, bolsillos, crédito y presupuesto. Este
        recorrido no se repetirá salvo que lo indiques en Administrar. Después revisa <Text style={{ fontWeight: '600' }}
        >Inicio</Text> (ahorro, donas y por cuenta), usa el <Text style={{ fontWeight: '600' }}>+</Text> o{' '}
        <Text style={{ fontWeight: '600' }}>Gastos</Text> para anotar salidas (incluido recibo), y entra a{' '}
        <Text style={{ fontWeight: '600' }}>Más → Reportes</Text> cuando quieras el informe cerrado de un mes.
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
