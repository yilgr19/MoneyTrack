import { Alert } from 'react-native';
import {
  obtenerCuentasOrigenGastoElegible,
  calcularSaldosPorCuenta,
  formatearNumero,
} from './finance';
import { puedeRegistrarCompraPorRegla48h } from './asistenteComprasLogic';

function pad(n) {
  return String(n).padStart(2, '0');
}

/**
 * Registra gasto y quita la intención. Si hay varias cuentas, pide origen por Alert.
 * @param {{ onRemoved?: () => void }} opts
 */
export function registrarGastoDesdeIntencionConUi({ state, intencion, origenValor, replaceState, onRemoved }) {
  if (!puedeRegistrarCompraPorRegla48h(intencion, Date.now())) {
    Alert.alert('Regla 48 h', 'Aún debes esperar antes de registrar esta compra.');
    return;
  }
  const precio = intencion.precioEstimado;
  const opts = obtenerCuentasOrigenGastoElegible(state || {}, precio, precio, {});
  if (opts.length === 0) {
    Alert.alert('Saldo', 'No hay ninguna cuenta con saldo suficiente para este monto.');
    return;
  }
  if (!origenValor && opts.length > 1) {
    Alert.alert(
      'Cuenta',
      '¿Desde qué caja pagas?',
      [
        ...opts.map((o) => ({
          text: o.label.slice(0, 60),
          onPress: () =>
            registrarGastoDesdeIntencionConUi({
              state,
              intencion,
              origenValor: o.value,
              replaceState,
              onRemoved,
            }),
        })),
        { text: 'Cancelar', style: 'cancel' },
      ],
      { cancelable: true }
    );
    return;
  }
  const origen = origenValor || opts[0].value;
  const disponible = opts.find((o) => o.value === origen);
  if (!disponible || precio > (disponible.saldo || 0)) {
    Alert.alert('Saldo', 'No hay suficiente saldo en la cuenta elegida. Revisa en Saldo.');
    return;
  }
  const d = new Date();
  const fechaStr = `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(
    d.getMinutes()
  )}:00`;
  const nuevo = {
    nombre: String(intencion.nombre || '').trim(),
    cantidad: precio,
    fecha: fechaStr,
    categoria: String(intencion.nombreCategoria || '').trim(),
    origen,
    nota: 'Registrado desde Asistente de compras',
    cuotas: 1,
    cuotaMensual: precio,
  };
  replaceState((s) => ({
    ...s,
    gastos: [...(s.gastos || []), nuevo],
    intencionesCompra: (s.intencionesCompra || []).filter((x) => x.id !== intencion.id),
  }));
  onRemoved?.();
  Alert.alert('Listo', 'Compra registrada en tu historial.');
}

export function yaNoLoQuieroIntencionConUi({ state, intencion, moneda, replaceState, onRemoved }) {
  const m = intencion.precioEstimado;
  const primera = (state?.metas || [])[0];
  const nombreMeta = primera ? String(primera.nombre || '').trim() : '';
  function quitar() {
    replaceState((s) => ({
      ...s,
      intencionesCompra: (s.intencionesCompra || []).filter((x) => x.id !== intencion.id),
    }));
    onRemoved?.();
  }
  const msg = primera
    ? `Acabas de ahorrar ${formatearNumero(m)} ${moneda}. Si tienes ese efectivo disponible, puedes aportar a tu meta «${nombreMeta}».`
    : `Acabas de ahorrar ${formatearNumero(m)} ${moneda}. Ese gasto imaginario ya no pesa en tu decisión; puedes reasignarlo cuando quieras.`;
  Alert.alert('¡Felicidades!', msg, [
    ...(primera
      ? [
          {
            text: `Aportar a «${nombreMeta.slice(0, 22)}${nombreMeta.length > 22 ? '…' : ''}»`,
            onPress: () => {
              const saldos = calcularSaldosPorCuenta(state || {});
              if ((saldos.efectivo || 0) < m) {
                Alert.alert(
                  'Saldo insuficiente en efectivo',
                  'Cuando puedas, aporta manualmente desde Metas.'
                );
                quitar();
                return;
              }
              const d = new Date();
              const fechaStr = `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
              replaceState((s) => ({
                ...s,
                intencionesCompra: (s.intencionesCompra || []).filter((x) => x.id !== intencion.id),
                contribucionesMetas: [
                  ...(s.contribucionesMetas || []),
                  { metaId: primera.id, cantidad: m, fecha: fechaStr, origen: 'efectivo' },
                ],
              }));
              onRemoved?.();
              Alert.alert('Listo', `Aporte de ${formatearNumero(m)} desde efectivo hacia «${nombreMeta}».`);
            },
          },
        ]
      : []),
    {
      text: primera ? 'Solo cerrar' : 'Gracias',
      style: 'cancel',
      onPress: quitar,
    },
  ]);
}
