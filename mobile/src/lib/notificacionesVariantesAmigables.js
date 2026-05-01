/**
 * Textos relajados y variados (20 por categoría) para notificaciones locales.
 * Se elige uno al azar al programar cada aviso.
 */

function pick(arr) {
  return arr[Math.floor(Math.random() * arr.length)];
}

function applyPair(pair, ctx) {
  const titulo = typeof pair.titulo === 'function' ? pair.titulo(ctx) : pair.titulo;
  const cuerpo = typeof pair.cuerpo === 'function' ? pair.cuerpo(ctx) : pair.cuerpo;
  return { title: titulo, body: cuerpo };
}

/** —— Pagos programados: vence hoy (d === 0) —— */
const PAGO_PROG_HOY = [
  { titulo: '🔔 Hoy toca, sin drama', cuerpo: (c) => `«${c.concepto}» cae hoy 😊 Son ${c.montoLine}. Cuando lo pagues, anótalo en Gastos y listo ✨` },
  { titulo: '📅 Oye, esto vence hoy', cuerpo: (c) => `«${c.concepto}» te saluda 👋 ${c.montoLine}. Nada de correr: un ratito y queda al día.` },
  { titulo: '☕ Café + pago pendiente', cuerpo: (c) => `Entre sorbo y sorbo: «${c.concepto}» hoy 🗓️ ${c.montoLine}. Tú puedes.` },
  { titulo: '🎯 Misión suave del día', cuerpo: (c) => `Liquidar «${c.concepto}» (${c.montoLine}). Luego anótalo en Gastos y a otra cosa mariposa 🦋` },
  { titulo: '💫 Hoy es el día', cuerpo: (c) => `«${c.concepto}» no espera más 😅 ${c.montoLine}. Cuando caiga, celebra el mini logro.` },
  { titulo: '🌤️ Aviso con buena onda', cuerpo: (c) => `To-do: «${c.concepto}» vence hoy 📝 ${c.montoLine}. Sin presión de jefe, solo un recordatorio amigo.` },
  { titulo: '🧘 Respira y págalo', cuerpo: (c) => `«${c.concepto}» · hoy · ${c.montoLine}. Ya verás cómo se siente bien tacharlo.` },
  { titulo: '🎒 Mochila mental', cuerpo: (c) => `Llevas «${c.concepto}» pendiente hoy 🎒 ${c.montoLine}. Un clic en el banco y chau.` },
  { titulo: '🌈 Sin estrés formal', cuerpo: (c) => `Ojo bonito: «${c.concepto}» vence hoy ✨ ${c.montoLine}. Luego regístralo en Gastos cuando quieras.` },
  { titulo: '🐢 Modo tortuga ganadora', cuerpo: (c) => `Lento pero seguro: «${c.concepto}» hoy 🐢 ${c.montoLine}. Tú mandas el ritmo.` },
  { titulo: '🍀 Pequeña suerte organizada', cuerpo: (c) => `«${c.concepto}» toca hoy 🍀 ${c.montoLine}. El dinero fluye y tú fluyes.` },
  { titulo: '🎧 Playlist + trámite', cuerpo: (c) => `Mientras suena tu rolita: «${c.concepto}» 🎵 ${c.montoLine}. Hoy es buen día.` },
  { titulo: '🧃 Hidrátate y págalo', cuerpo: (c) => `«${c.concepto}» vence hoy 🧃 ${c.montoLine}. Dos minutos y queda en paz.` },
  { titulo: '🌙 Aún es de día para hacerlo', cuerpo: (c) => `«${c.concepto}» · vence hoy 🌤️ ${c.montoLine}. Sin culpa, con calma.` },
  { titulo: '🎪 Circo de adultos', cuerpo: (c) => `Hoy en el programa: pagar «${c.concepto}» 🎪 ${c.montoLine}. El público (tú) aplaude al final.` },
  { titulo: '🦩 Flamenco de la responsabilidad', cuerpo: (c) => `«${c.concepto}» baila hoy 🦩 ${c.montoLine}. Un paso y listo.` },
  { titulo: '🍕 Después del snack…', cuerpo: (c) => `…«${c.concepto}» 🍕 ${c.montoLine}. Hoy toca; mañana más pizza si quieres.` },
  { titulo: '🚀 Modo cohete suave', cuerpo: (c) => `«${c.concepto}» despega hoy 🚀 ${c.montoLine}. Sin turbulencia, prometido.` },
  { titulo: '🧸 Abrazo y recordatorio', cuerpo: (c) => `«${c.concepto}» te recuerda con cariño que es hoy 🧸 ${c.montoLine}.` },
  { titulo: '✅ Tick mental', cuerpo: (c) => `«${c.concepto}» → hoy ✅ ${c.montoLine}. Gastos te espera para el tick real.` },
];

const PAGO_PROG_D1 = [
  { titulo: '⏳ Mañana toca ese pago', cuerpo: (c) => `«${c.concepto}» · te queda 1 día 🙌 ${c.montoLine}. Hoy puedes adelantarlo si quieres.` },
  { titulo: '📆 Falta solo un día', cuerpo: (c) => `«${c.concepto}» vence mañana 📆 ${c.montoLine}. Respira y organízalo con tiempo.` },
  { titulo: '🌙 Antes de que llegue mañana', cuerpo: (c) => `«${c.concepto}» en 24 h aprox 🌙 ${c.montoLine}. Va quedando poquito.` },
  { titulo: '🎯 Casi en la meta', cuerpo: (c) => `Un día para «${c.concepto}» 🎯 ${c.montoLine}. Tú llevas el control.` },
  { titulo: '☕ Un día de colchón', cuerpo: (c) => `«${c.concepto}» cae mañana ☕ ${c.montoLine}. Colchón perfecto para no correr.` },
  { titulo: '🧩 Pieza que encaja mañana', cuerpo: (c) => `«${c.concepto}» 🧩 vence mañana · ${c.montoLine}.` },
  { titulo: '🎒 Prepárate con calma', cuerpo: (c) => `Mañana: «${c.concepto}» 🎒 ${c.montoLine}. Hoy ya puedes dejarlo listo si te anima.` },
  { titulo: '🌈 Arcoíris de plazos', cuerpo: (c) => `Queda 1 día para «${c.concepto}» 🌈 ${c.montoLine}. Todo fluye.` },
  { titulo: '🐢 Ritmo tranquilo', cuerpo: (c) => `«${c.concepto}» mañana 🐢 ${c.montoLine}. Sin apuro absurdo.` },
  { titulo: '🎧 Mañana con música', cuerpo: (c) => `«${c.concepto}» vence mañana 🎧 ${c.montoLine}. Playlist lista, pago listo.` },
  { titulo: '🍀 Un día de suerte', cuerpo: (c) => `Mañana toca «${c.concepto}» 🍀 ${c.montoLine}.` },
  { titulo: '🦩 Baila hacia el plazo', cuerpo: (c) => `1 día para «${c.concepto}» 🦩 ${c.montoLine}.` },
  { titulo: '🧃 Fresquito el recordatorio', cuerpo: (c) => `«${c.concepto}» mañana 🧃 ${c.montoLine}.` },
  { titulo: '🌤️ Clarito y con tiempo', cuerpo: (c) => `Mañana vence «${c.concepto}» 🌤️ ${c.montoLine}.` },
  { titulo: '🎪 Mañana función', cuerpo: (c) => `«${c.concepto}» entra mañana 🎪 ${c.montoLine}.` },
  { titulo: '🚀 Countdown suave', cuerpo: (c) => `T–1 día: «${c.concepto}» 🚀 ${c.montoLine}.` },
  { titulo: '🧸 Cariño y calendario', cuerpo: (c) => `Mañana «${c.concepto}» 🧸 ${c.montoLine}.` },
  { titulo: '✅ Casi tachado', cuerpo: (c) => `1 día para «${c.concepto}» ✅ ${c.montoLine}.` },
  { titulo: '🍕 Pizza de paciencia', cuerpo: (c) => `«${c.concepto}» mañana 🍕 ${c.montoLine}.` },
  { titulo: '🌙 Sueña tranquilo', cuerpo: (c) => `Mañana pagas «${c.concepto}» 🌙 ${c.montoLine}. Hoy descansa la cabeza.` },
];

const PAGO_PROG_D2 = [
  { titulo: '🗓️ Quedan 2 días', cuerpo: (c) => `«${c.concepto}» · 2 días de colchón 📅 ${c.montoLine}. Vas sobrado de tiempo.` },
  { titulo: '🎯 A la vuelta de la esquina', cuerpo: (c) => `«${c.concepto}» en 2 días 🎯 ${c.montoLine}.` },
  { titulo: '☕ Café para planear', cuerpo: (c) => `2 días para «${c.concepto}» ☕ ${c.montoLine}.` },
  { titulo: '🌈 Espacio para respirar', cuerpo: (c) => `«${c.concepto}» vence en 2 días 🌈 ${c.montoLine}.` },
  { titulo: '🐢 Modo preventivo', cuerpo: (c) => `2 días → «${c.concepto}» 🐢 ${c.montoLine}.` },
  { titulo: '🎧 Playlist larga', cuerpo: (c) => `«${c.concepto}» en 2 días 🎧 ${c.montoLine}.` },
  { titulo: '🍀 Tiempo de sobra', cuerpo: (c) => `2 días para «${c.concepto}» 🍀 ${c.montoLine}.` },
  { titulo: '🦩 Flamenco relajado', cuerpo: (c) => `«${c.concepto}» · 2 días 🦩 ${c.montoLine}.` },
  { titulo: '🧃 Todo bajo control', cuerpo: (c) => `2 días: «${c.concepto}» 🧃 ${c.montoLine}.` },
  { titulo: '🌤️ Buen margen', cuerpo: (c) => `«${c.concepto}» en 2 días 🌤️ ${c.montoLine}.` },
  { titulo: '🎪 Antesala del pago', cuerpo: (c) => `2 días para «${c.concepto}» 🎪 ${c.montoLine}.` },
  { titulo: '🚀 Countdown 2', cuerpo: (c) => `«${c.concepto}» T-2 🚀 ${c.montoLine}.` },
  { titulo: '🧸 Sin apuro', cuerpo: (c) => `2 días · «${c.concepto}» 🧸 ${c.montoLine}.` },
  { titulo: '✅ Calma administrativa', cuerpo: (c) => `«${c.concepto}» en 2 días ✅ ${c.montoLine}.` },
  { titulo: '🍕 Pizza y plan', cuerpo: (c) => `2 días para «${c.concepto}» 🍕 ${c.montoLine}.` },
  { titulo: '🌙 Luna de plazos', cuerpo: (c) => `«${c.concepto}» · 2 días 🌙 ${c.montoLine}.` },
  { titulo: '🧩 Encaja perfecto', cuerpo: (c) => `2 días → «${c.concepto}» 🧩 ${c.montoLine}.` },
  { titulo: '🎒 Mochila de tiempo', cuerpo: (c) => `«${c.concepto}» en 2 días 🎒 ${c.montoLine}.` },
  { titulo: '💫 Suave aviso', cuerpo: (c) => `2 días para «${c.concepto}» 💫 ${c.montoLine}.` },
  { titulo: '🔔 Timbre lejano', cuerpo: (c) => `«${c.concepto}» vence en 2 días 🔔 ${c.montoLine}.` },
];

const PAGO_PROG_D3 = [
  { titulo: '📅 3 días para organizarte', cuerpo: (c) => `«${c.concepto}» · 3 días de aire 📅 ${c.montoLine}.` },
  { titulo: '🌈 Margen cómodo', cuerpo: (c) => `3 días para «${c.concepto}» 🌈 ${c.montoLine}.` },
  { titulo: '☕ Café y calendario', cuerpo: (c) => `«${c.concepto}» en 3 días ☕ ${c.montoLine}.` },
  { titulo: '🎯 Vista larga', cuerpo: (c) => `3 días → «${c.concepto}» 🎯 ${c.montoLine}.` },
  { titulo: '🐢 Tortuga feliz', cuerpo: (c) => `«${c.concepto}» · 3 días 🐢 ${c.montoLine}.` },
  { titulo: '🎧 Sin estrés', cuerpo: (c) => `3 días para «${c.concepto}» 🎧 ${c.montoLine}.` },
  { titulo: '🍀 Buen colchón', cuerpo: (c) => `«${c.concepto}» en 3 días 🍀 ${c.montoLine}.` },
  { titulo: '🦩 Baila tranquilo', cuerpo: (c) => `3 días: «${c.concepto}» 🦩 ${c.montoLine}.` },
  { titulo: '🧃 Fresquito', cuerpo: (c) => `«${c.concepto}» en 3 días 🧃 ${c.montoLine}.` },
  { titulo: '🌤️ Cielo despejado', cuerpo: (c) => `3 días para «${c.concepto}» 🌤️ ${c.montoLine}.` },
  { titulo: '🎪 Preventivo', cuerpo: (c) => `«${c.concepto}» · 3 días 🎪 ${c.montoLine}.` },
  { titulo: '🚀 T-3', cuerpo: (c) => `3 días → «${c.concepto}» 🚀 ${c.montoLine}.` },
  { titulo: '🧸 Abrazo', cuerpo: (c) => `«${c.concepto}» en 3 días 🧸 ${c.montoLine}.` },
  { titulo: '✅ Todo bien', cuerpo: (c) => `3 días: «${c.concepto}» ✅ ${c.montoLine}.` },
  { titulo: '🍕 Modo chill', cuerpo: (c) => `«${c.concepto}» en 3 días 🍕 ${c.montoLine}.` },
  { titulo: '🌙 Onda relaj', cuerpo: (c) => `3 días para «${c.concepto}» 🌙 ${c.montoLine}.` },
  { titulo: '🧩 Pieza lejana', cuerpo: (c) => `«${c.concepto}» · 3 días 🧩 ${c.montoLine}.` },
  { titulo: '🎒 Tiempo extra', cuerpo: (c) => `3 días → «${c.concepto}» 🎒 ${c.montoLine}.` },
  { titulo: '💫 Aviso suave', cuerpo: (c) => `«${c.concepto}» en 3 días 💫 ${c.montoLine}.` },
  { titulo: '🔔 Campanilla lejana', cuerpo: (c) => `3 días para «${c.concepto}» 🔔 ${c.montoLine}.` },
];

/** Prueba: vencimiento lejano (solo dev) */
const PAGO_PROG_PRUEBA_LEJANO = [
  { titulo: '🧪 Prueba · canal OK', cuerpo: (c) => `«${c.concepto}» vence en ${c.dHoy} días 🔭 ${c.montoLine}. Los avisos reales saltan 0–3 días antes. Cierra la app y espera un poco ✨` },
  { titulo: '🎧 Modo laboratorio', cuerpo: (c) => `Test: «${c.concepto}» a ${c.dHoy} días 🎧 ${c.montoLine}. Si ves esto, el canal vive.` },
  { titulo: '🚀 Ping amistoso', cuerpo: (c) => `«${c.concepto}» · ${c.dHoy} días 🚀 ${c.montoLine}. Prueba de notificación.` },
  { titulo: '🌈 Todo en orden', cuerpo: (c) => `Aviso test · «${c.concepto}» 🌈 ${c.montoLine}. Vence en ${c.dHoy} días (lejos).` },
  { titulo: '🐢 Sin prisa real', cuerpo: (c) => `«${c.concepto}» aún lejos 🐢 ${c.montoLine}. ${c.dHoy} días. Solo probamos.` },
  { titulo: '☕ Café y test', cuerpo: (c) => `Prueba local ☕ «${c.concepto}» · ${c.dHoy} d · ${c.montoLine}` },
  { titulo: '🎯 Diana de prueba', cuerpo: (c) => `«${c.concepto}» 🎯 ${c.dHoy} días ${c.montoLine}` },
  { titulo: '🦩 Flamenco beta', cuerpo: (c) => `Test 🦩 «${c.concepto}» ${c.montoLine} · ${c.dHoy} d` },
  { titulo: '🧃 Beta fresca', cuerpo: (c) => `Notif prueba 🧃 «${c.concepto}» ${c.montoLine}` },
  { titulo: '🌤️ Cielo de test', cuerpo: (c) => `«${c.concepto}» 🌤️ vence en ${c.dHoy} d · ${c.montoLine}` },
  { titulo: '🎪 Circo de prueba', cuerpo: (c) => `Show test 🎪 «${c.concepto}» ${c.montoLine}` },
  { titulo: '🧸 Abrazo beta', cuerpo: (c) => `«${c.concepto}» 🧸 ${c.dHoy} d ${c.montoLine}` },
  { titulo: '✅ Check test', cuerpo: (c) => `Prueba ✅ «${c.concepto}» ${c.montoLine}` },
  { titulo: '🍕 Pizza test', cuerpo: (c) => `«${c.concepto}» 🍕 ${c.dHoy} d ${c.montoLine}` },
  { titulo: '🌙 Luna test', cuerpo: (c) => `Test nocturno 🌙 «${c.concepto}» ${c.montoLine}` },
  { titulo: '🧩 Puzzle test', cuerpo: (c) => `Pieza 🧩 «${c.concepto}» ${c.montoLine}` },
  { titulo: '🎒 Mochila test', cuerpo: (c) => `«${c.concepto}» 🎒 ${c.dHoy} d ${c.montoLine}` },
  { titulo: '💫 Spark test', cuerpo: (c) => `Spark 💫 «${c.concepto}» ${c.montoLine}` },
  { titulo: '🔔 Campana test', cuerpo: (c) => `Ping 🔔 «${c.concepto}» ${c.montoLine}` },
  { titulo: '🍀 Lucky test', cuerpo: (c) => `Suerte 🍀 notif «${c.concepto}» ${c.montoLine}` },
];

/** Notificación manual “probar en 20s” */
const PRUEBA_SISTEMA = [
  { titulo: '✨ MoneyTrack te saluda', cuerpo: 'Si ves esto con la app cerrada, las notificaciones locales van fino 🎉' },
  { titulo: '🎉 ¡Funciona!', cuerpo: 'Las alarmas del sistema y MoneyTrack están en la misma onda ✨' },
  { titulo: '🚀 Ping recibido', cuerpo: 'Notificación de prueba OK. Puedes celebrar un segundo 🍀' },
  { titulo: '🔔 Timbrazo amistoso', cuerpo: 'Todo en orden: el canal llegó hasta aquí sin drama 😊' },
  { titulo: '☕ Café virtual', cuerpo: 'Prueba superada. Tu app avisa aunque esté minimizada ☕' },
  { titulo: '🌈 Arcoíris técnico', cuerpo: 'Notif local = check. Sigue así 🌈' },
  { titulo: '🐢 Tortuga feliz', cuerpo: 'Lento pero seguro: el aviso llegó 🐢' },
  { titulo: '🎯 Diana centrada', cuerpo: 'Prueba en el blanco. Notificaciones ON 🎯' },
  { titulo: '🦩 Flamenco aprobado', cuerpo: 'El sistema y tú: equipo ganador 🦩' },
  { titulo: '🧃 Refresco de datos', cuerpo: 'Canal fresco, notif lista 🧃' },
  { titulo: '🌤️ Día despejado', cuerpo: 'Sin nubes en la notificación 🌤️' },
  { titulo: '🎪 Función cumplida', cuerpo: 'El show de prueba terminó bien 🎪' },
  { titulo: '🧸 Abrazo digital', cuerpo: 'MoneyTrack te avisó con cariño 🧸' },
  { titulo: '✅ Tick verde', cuerpo: 'Permisos y canal: todo OK ✅' },
  { titulo: '🍕 Rebanada de éxito', cuerpo: 'Prueba sabrosa completada 🍕' },
  { titulo: '🌙 Brillo nocturno', cuerpo: 'Aviso nocturno/diurno funcionando 🌙' },
  { titulo: '🧩 Pieza encajada', cuerpo: 'Notificación encaja perfecto 🧩' },
  { titulo: '🎒 Mochila lista', cuerpo: 'Sistema de avisos en la mochila 🎒' },
  { titulo: '💫 Magia menor', cuerpo: 'Un poquito de magia en la barra 💫' },
  { titulo: '🍀 Hoja de suerte', cuerpo: 'Prueba con buena vibra 🍀' },
];

function lineaMonto(montoStr, moneda) {
  return moneda ? `${montoStr} ${moneda}` : `${montoStr}`;
}

export function varianteNotifPagoProgramado(d, { concepto, montoStr, moneda }) {
  const montoLine = lineaMonto(montoStr, moneda);
  const ctx = { concepto, montoStr, moneda, montoLine };
  const pool = d === 0 ? PAGO_PROG_HOY : d === 1 ? PAGO_PROG_D1 : d === 2 ? PAGO_PROG_D2 : PAGO_PROG_D3;
  return applyPair(pick(pool), ctx);
}

export function varianteNotifPagoPruebaLejano({ concepto, montoStr, moneda, dHoy }) {
  const montoLine = lineaMonto(montoStr, moneda);
  return applyPair(pick(PAGO_PROG_PRUEBA_LEJANO), { concepto, montoStr, moneda, montoLine, dHoy });
}

export function varianteNotifPruebaSistema() {
  return applyPair(pick(PRUEBA_SISTEMA), {});
}

/** —— TC: corte —— */
const TC_CORTE_HOY = [
  { titulo: (c) => `📅 Hoy es corte · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `Mira el extracto con calma 👀 Guía ~${c.pagoSug}. Luego Gastos si aplica ✨` : `Revisa Saldo y el banco sin estrés 🏦 Todo fluye.`) },
  { titulo: (c) => `🗓️ Corte hoy · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `Tip relaj: ~${c.pagoSug} 💳 Confirma en tu app del banco.` : `Un vistazo a movimientos y listo 👀`) },
  { titulo: (c) => `☕ Café + corte · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `Referencia suave ~${c.pagoSug} ☕ Sin drama.` : `Día de cierre amable para ${c.nom} ☕`) },
  { titulo: (c) => `🎯 Cierre del ciclo · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} como guía 🎯 Tú mandas.` : `Corte hoy · revisa cuando puedas 🎯`) },
  { titulo: (c) => `🌈 Hoy corta ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `Guía ~${c.pagoSug} 🌈` : `Saldo y banco en modo chill 🌈`) },
  { titulo: (c) => `🐢 Corte tortuga · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🐢` : `Sin correr 🐢`) },
  { titulo: (c) => `🎧 Playlist y corte · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🎧` : `Revisa con música 🎧`) },
  { titulo: (c) => `🍀 Corte con suerte · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🍀` : `Todo bien 🍀`) },
  { titulo: (c) => `🦩 Flamenco corte · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🦩` : `Hoy corta 🦩`) },
  { titulo: (c) => `🧃 Corte fresco · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🧃` : `Chill 🧃`) },
  { titulo: (c) => `🌤️ Cielo de corte · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🌤️` : `Despejado 🌤️`) },
  { titulo: (c) => `🎪 Show de corte · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🎪` : `Función hoy 🎪`) },
  { titulo: (c) => `🚀 Corte launch · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🚀` : `Despega 🚀`) },
  { titulo: (c) => `🧸 Corte con abrazo · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🧸` : `Cariño 🧸`) },
  { titulo: (c) => `✅ Corte checklist · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} ✅` : `Tick mental ✅`) },
  { titulo: (c) => `🍕 Corte y pizza · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🍕` : `Premio después 🍕`) },
  { titulo: (c) => `🌙 Corte lunar · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🌙` : `Noche de revisar 🌙`) },
  { titulo: (c) => `🧩 Corte puzzle · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🧩` : `Encaja 🧩`) },
  { titulo: (c) => `🎒 Corte mochila · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🎒` : `Lleva el control 🎒`) },
  { titulo: (c) => `💫 Corte sparkle · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 💫` : `Brillo 💫`) },
];

const TC_CORTE_MANANA = [
  { titulo: (c) => `⏳ Mañana corte · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `Prepárate con calma ~${c.pagoSug} 🙌` : `Mañana revisas ${c.nom} sin apuro 📆`) },
  { titulo: (c) => `📆 Mañana cierra ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `Guía ~${c.pagoSug} 📆` : `Un día de colchón 📆`) },
  { titulo: (c) => `☕ Corte mañana · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} ☕` : `Café y plan ☕`) },
  { titulo: (c) => `🎯 T–1 corte · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🎯` : `Casi 🎯`) },
  { titulo: (c) => `🌈 Arcoíris mañana · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🌈` : `Mañana 🌈`) },
  { titulo: (c) => `🐢 Mañana tortuga · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🐢` : `🐢`) },
  { titulo: (c) => `🎧 Mañana música · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🎧` : `🎧`) },
  { titulo: (c) => `🍀 Mañana suerte · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🍀` : `🍀`) },
  { titulo: (c) => `🦩 Mañana flamenco · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🦩` : `🦩`) },
  { titulo: (c) => `🧃 Mañana fresco · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🧃` : `🧃`) },
  { titulo: (c) => `🌤️ Mañana claro · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🌤️` : `🌤️`) },
  { titulo: (c) => `🎪 Mañana show · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🎪` : `🎪`) },
  { titulo: (c) => `🚀 Mañana launch · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🚀` : `🚀`) },
  { titulo: (c) => `🧸 Mañana abrazo · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🧸` : `🧸`) },
  { titulo: (c) => `✅ Mañana tick · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} ✅` : `✅`) },
  { titulo: (c) => `🍕 Mañana pizza · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🍕` : `🍕`) },
  { titulo: (c) => `🌙 Mañana luna · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🌙` : `🌙`) },
  { titulo: (c) => `🧩 Mañana puzzle · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🧩` : `🧩`) },
  { titulo: (c) => `🎒 Mañana mochila · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🎒` : `🎒`) },
  { titulo: (c) => `💫 Mañana sparkle · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 💫` : `💫`) },
];

const TC_CORTE_2D = [
  { titulo: (c) => `🗓️ En 2 días · corte ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `Aire para organizarte ~${c.pagoSug} 😌` : `Buen momento para ver Saldo 🗓️`) },
  { titulo: (c) => `📅 2 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 📅` : `📅`) },
  { titulo: (c) => `☕ 2 días café · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} ☕` : `☕`) },
  { titulo: (c) => `🎯 2 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🎯` : `🎯`) },
  { titulo: (c) => `🌈 2 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🌈` : `🌈`) },
  { titulo: (c) => `🐢 2 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🐢` : `🐢`) },
  { titulo: (c) => `🎧 2 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🎧` : `🎧`) },
  { titulo: (c) => `🍀 2 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🍀` : `🍀`) },
  { titulo: (c) => `🦩 2 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🦩` : `🦩`) },
  { titulo: (c) => `🧃 2 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🧃` : `🧃`) },
  { titulo: (c) => `🌤️ 2 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🌤️` : `🌤️`) },
  { titulo: (c) => `🎪 2 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🎪` : `🎪`) },
  { titulo: (c) => `🚀 2 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🚀` : `🚀`) },
  { titulo: (c) => `🧸 2 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🧸` : `🧸`) },
  { titulo: (c) => `✅ 2 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} ✅` : `✅`) },
  { titulo: (c) => `🍕 2 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🍕` : `🍕`) },
  { titulo: (c) => `🌙 2 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🌙` : `🌙`) },
  { titulo: (c) => `🧩 2 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🧩` : `🧩`) },
  { titulo: (c) => `🎒 2 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🎒` : `🎒`) },
  { titulo: (c) => `💫 2 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 💫` : `💫`) },
];

function applyTcPair(pair, ctx) {
  const titulo = typeof pair.titulo === 'function' ? pair.titulo(ctx) : pair.titulo;
  const cuerpo = typeof pair.cuerpo === 'function' ? pair.cuerpo(ctx) : pair.cuerpo;
  return { title: titulo, body: cuerpo };
}

export function varianteNotifTcCorteHoy(ctx) {
  return applyTcPair(pick(TC_CORTE_HOY), ctx);
}
export function varianteNotifTcCorteManana(ctx) {
  return applyTcPair(pick(TC_CORTE_MANANA), ctx);
}
export function varianteNotifTcCorte2d(ctx) {
  return applyTcPair(pick(TC_CORTE_2D), ctx);
}

/** TC: pago límite dp 0..3 */
const TC_PAGO_0 = [
  { titulo: (c) => `💳 Hoy pago · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `Guía relaj ~${c.pagoSug} ✨ Banco + calma.` : `Revisa límite en Saldo 💳 Sin drama.`) },
  { titulo: (c) => `📅 Vence hoy · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 📅` : `Hoy mira el banco 📅`) },
  { titulo: (c) => `☕ Café y pago · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} ☕` : `☕`) },
  { titulo: (c) => `🎯 Hoy · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🎯` : `🎯`) },
  { titulo: (c) => `🌈 Hoy límite · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🌈` : `🌈`) },
  { titulo: (c) => `🐢 Hoy tortuga · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🐢` : `🐢`) },
  { titulo: (c) => `🎧 Hoy música · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🎧` : `🎧`) },
  { titulo: (c) => `🍀 Hoy suerte · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🍀` : `🍀`) },
  { titulo: (c) => `🦩 Hoy flamenco · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🦩` : `🦩`) },
  { titulo: (c) => `🧃 Hoy fresco · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🧃` : `🧃`) },
  { titulo: (c) => `🌤️ Hoy claro · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🌤️` : `🌤️`) },
  { titulo: (c) => `🎪 Hoy show · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🎪` : `🎪`) },
  { titulo: (c) => `🚀 Hoy launch · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🚀` : `🚀`) },
  { titulo: (c) => `🧸 Hoy abrazo · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🧸` : `🧸`) },
  { titulo: (c) => `✅ Hoy tick · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} ✅` : `✅`) },
  { titulo: (c) => `🍕 Hoy pizza · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🍕` : `🍕`) },
  { titulo: (c) => `🌙 Hoy luna · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🌙` : `🌙`) },
  { titulo: (c) => `🧩 Hoy puzzle · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🧩` : `🧩`) },
  { titulo: (c) => `🎒 Hoy mochila · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🎒` : `🎒`) },
  { titulo: (c) => `💫 Hoy sparkle · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 💫` : `💫`) },
];

const TC_PAGO_1 = [
  { titulo: (c) => `⏳ Mañana pago · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🙌 Mañana con calma.` : `Mañana miras el banco 📆`) },
  { titulo: (c) => `📆 Mañana límite · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 📆` : `📆`) },
  { titulo: (c) => `☕ Mañana café · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} ☕` : `☕`) },
  { titulo: (c) => `🎯 Mañana · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🎯` : `🎯`) },
  { titulo: (c) => `🌈 Mañana · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🌈` : `🌈`) },
  { titulo: (c) => `🐢 Mañana · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🐢` : `🐢`) },
  { titulo: (c) => `🎧 Mañana · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🎧` : `🎧`) },
  { titulo: (c) => `🍀 Mañana · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🍀` : `🍀`) },
  { titulo: (c) => `🦩 Mañana · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🦩` : `🦩`) },
  { titulo: (c) => `🧃 Mañana · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🧃` : `🧃`) },
  { titulo: (c) => `🌤️ Mañana · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🌤️` : `🌤️`) },
  { titulo: (c) => `🎪 Mañana · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🎪` : `🎪`) },
  { titulo: (c) => `🚀 Mañana · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🚀` : `🚀`) },
  { titulo: (c) => `🧸 Mañana · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🧸` : `🧸`) },
  { titulo: (c) => `✅ Mañana · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} ✅` : `✅`) },
  { titulo: (c) => `🍕 Mañana · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🍕` : `🍕`) },
  { titulo: (c) => `🌙 Mañana · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🌙` : `🌙`) },
  { titulo: (c) => `🧩 Mañana · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🧩` : `🧩`) },
  { titulo: (c) => `🎒 Mañana · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🎒` : `🎒`) },
  { titulo: (c) => `💫 Mañana · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 💫` : `💫`) },
];

const TC_PAGO_2 = [
  { titulo: (c) => `📅 En 2 días · pago ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 😌` : `Margen para ${c.nom} 📅`) },
  { titulo: (c) => `🗓️ 2 días pago · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🗓️` : `🗓️`) },
  { titulo: (c) => `☕ 2 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} ☕` : `☕`) },
  { titulo: (c) => `🎯 2 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🎯` : `🎯`) },
  { titulo: (c) => `🌈 2 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🌈` : `🌈`) },
  { titulo: (c) => `🐢 2 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🐢` : `🐢`) },
  { titulo: (c) => `🎧 2 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🎧` : `🎧`) },
  { titulo: (c) => `🍀 2 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🍀` : `🍀`) },
  { titulo: (c) => `🦩 2 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🦩` : `🦩`) },
  { titulo: (c) => `🧃 2 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🧃` : `🧃`) },
  { titulo: (c) => `🌤️ 2 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🌤️` : `🌤️`) },
  { titulo: (c) => `🎪 2 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🎪` : `🎪`) },
  { titulo: (c) => `🚀 2 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🚀` : `🚀`) },
  { titulo: (c) => `🧸 2 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🧸` : `🧸`) },
  { titulo: (c) => `✅ 2 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} ✅` : `✅`) },
  { titulo: (c) => `🍕 2 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🍕` : `🍕`) },
  { titulo: (c) => `🌙 2 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🌙` : `🌙`) },
  { titulo: (c) => `🧩 2 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🧩` : `🧩`) },
  { titulo: (c) => `🎒 2 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🎒` : `🎒`) },
  { titulo: (c) => `💫 2 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 💫` : `💫`) },
];

const TC_PAGO_3 = [
  { titulo: (c) => `📆 En 3 días · pago ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} ✨` : `Buen margen para ${c.nom} 📆`) },
  { titulo: (c) => `🗓️ 3 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🗓️` : `🗓️`) },
  { titulo: (c) => `☕ 3 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} ☕` : `☕`) },
  { titulo: (c) => `🎯 3 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🎯` : `🎯`) },
  { titulo: (c) => `🌈 3 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🌈` : `🌈`) },
  { titulo: (c) => `🐢 3 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🐢` : `🐢`) },
  { titulo: (c) => `🎧 3 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🎧` : `🎧`) },
  { titulo: (c) => `🍀 3 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🍀` : `🍀`) },
  { titulo: (c) => `🦩 3 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🦩` : `🦩`) },
  { titulo: (c) => `🧃 3 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🧃` : `🧃`) },
  { titulo: (c) => `🌤️ 3 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🌤️` : `🌤️`) },
  { titulo: (c) => `🎪 3 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🎪` : `🎪`) },
  { titulo: (c) => `🚀 3 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🚀` : `🚀`) },
  { titulo: (c) => `🧸 3 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🧸` : `🧸`) },
  { titulo: (c) => `✅ 3 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} ✅` : `✅`) },
  { titulo: (c) => `🍕 3 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🍕` : `🍕`) },
  { titulo: (c) => `🌙 3 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🌙` : `🌙`) },
  { titulo: (c) => `🧩 3 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🧩` : `🧩`) },
  { titulo: (c) => `🎒 3 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 🎒` : `🎒`) },
  { titulo: (c) => `💫 3 días · ${c.nom}`, cuerpo: (c) => (c.pagoSug ? `~${c.pagoSug} 💫` : `💫`) },
];

export function varianteNotifTcPago(dp, ctx) {
  const pool = dp === 0 ? TC_PAGO_0 : dp === 1 ? TC_PAGO_1 : dp === 2 ? TC_PAGO_2 : TC_PAGO_3;
  return applyTcPair(pick(pool), ctx);
}

/** Campana in-app: gasto editado (20 variantes, tono cercano como pagos/lista). */
const GASTO_EDITADO_CAMPANA = [
  { titulo: '✏️ Listo, quedó actualizado', cuerpo: (c) => `«${c.nombre}» ya va con tus cambios · ${c.montoLine} ✨` },
  { titulo: '📝 Ajuste guardado con cariño', cuerpo: (c) => `Retocaste «${c.nombre}» y la app lo celebra · ${c.montoLine} 💛` },
  { titulo: '🎯 Gasto al día', cuerpo: (c) => `«${c.nombre}» quedó fino en el historial · ${c.montoLine} 😊` },
  { titulo: '☕ Cambio suave', cuerpo: (c) => `Como un cafecito: «${c.nombre}» actualizado · ${c.montoLine} ☕` },
  { titulo: '🌈 Todo cuadra mejor', cuerpo: (c) => `«${c.nombre}» refleja lo real ahora · ${c.montoLine} 🌈` },
  { titulo: '🐢 Sin prisa, bien hecho', cuerpo: (c) => `Editaste «${c.nombre}» con calma · ${c.montoLine} 🐢` },
  { titulo: '🎧 Mini victoria', cuerpo: (c) => `«${c.nombre}» ya está como tú querías · ${c.montoLine} 🎧` },
  { titulo: '🍀 Buen ojo', cuerpo: (c) => `Corregir «${c.nombre}» también es cuidarte · ${c.montoLine} 🍀` },
  { titulo: '🦩 Flamenco orgulloso', cuerpo: (c) => `«${c.nombre}» baila con datos nuevos · ${c.montoLine} 🦩` },
  { titulo: '🧃 Fresquito el cambio', cuerpo: (c) => `«${c.nombre}» refrescado en movimientos · ${c.montoLine} 🧃` },
  { titulo: '🌤️ Más claro que antes', cuerpo: (c) => `«${c.nombre}» ya luce alineado · ${c.montoLine} 🌤️` },
  { titulo: '🎪 Función: edición', cuerpo: (c) => `«${c.nombre}» salió del telón renovado · ${c.montoLine} 🎪` },
  { titulo: '🚀 Pequeño cohete de orden', cuerpo: (c) => `«${c.nombre}» actualizado · ${c.montoLine} 🚀` },
  { titulo: '🧸 Abrazo al detalle', cuerpo: (c) => `«${c.nombre}» con tu toque · ${c.montoLine} 🧸` },
  { titulo: '✅ Tick de claridad', cuerpo: (c) => `«${c.nombre}» editado; resumen más honesto · ${c.montoLine} ✅` },
  { titulo: '🍕 Pizza de buena decisión', cuerpo: (c) => `«${c.nombre}» quedó como corresponde · ${c.montoLine} 🍕` },
  { titulo: '🌙 Noche tranquila', cuerpo: (c) => `«${c.nombre}» ya no te va a rondar mal en la app · ${c.montoLine} 🌙` },
  { titulo: '🧩 Pieza encajada', cuerpo: (c) => `«${c.nombre}» encaja mejor ahora · ${c.montoLine} 🧩` },
  { titulo: '🎒 Mochila mental más liviana', cuerpo: (c) => `«${c.nombre}» corregido · ${c.montoLine} 🎒` },
  { titulo: '💫 Brillo de orden', cuerpo: (c) => `«${c.nombre}» brilla con tu edición · ${c.montoLine} 💫` },
];

/** Campana in-app: gasto quitado del historial (20 variantes). */
const GASTO_ELIMINADO_CAMPANA = [
  { titulo: '🗑️ Fuera del historial', cuerpo: (c) => `«${c.nombre}» ya no cuenta en movimientos · era ${c.montoLine} · sin drama ✨` },
  { titulo: '🌿 Registro limpio', cuerpo: (c) => `Quitaste «${c.nombre}» (${c.montoLine}) · tu libreta respira 🌿` },
  { titulo: '🧹 Mini limpieza', cuerpo: (c) => `«${c.nombre}» salió (${c.montoLine}) · a veces hay que borrar para aclarar 🧹` },
  { titulo: '☕ Como borrar un typo', cuerpo: (c) => `«${c.nombre}» eliminado · ${c.montoLine} · nadie te juzga ☕` },
  { titulo: '🌈 Espacio nuevo', cuerpo: (c) => `«${c.nombre}» ya no está (${c.montoLine}) · el resumen se ajusta solo 🌈` },
  { titulo: '🐢 Paso atrás sin culpa', cuerpo: (c) => `«${c.nombre}» fuera (${c.montoLine}) · la tortuga también corrige 🐢` },
  { titulo: '🎧 Silencio en esa línea', cuerpo: (c) => `«${c.nombre}» quitado · ${c.montoLine} · playlist más clara 🎧` },
  { titulo: '🍀 Segunda oportunidad al resumen', cuerpo: (c) => `«${c.nombre}» eliminado (${c.montoLine}) · los números vuelven a cuadrar 🍀` },
  { titulo: '🦩 Baila sin esa fila', cuerpo: (c) => `«${c.nombre}» ya no baila en tu historial · ${c.montoLine} 🦩` },
  { titulo: '🧃 Fresquito el cambio', cuerpo: (c) => `«${c.nombre}» retirado (${c.montoLine}) · bebida mental fresca 🧃` },
  { titulo: '🌤️ Cielo más despejado', cuerpo: (c) => `«${c.nombre}» fuera · ${c.montoLine} · menos ruido en la app 🌤️` },
  { titulo: '🎪 Cortina en ese gasto', cuerpo: (c) => `«${c.nombre}» cerró función (${c.montoLine}) 🎪` },
  { titulo: '🚀 Despegó del listado', cuerpo: (c) => `«${c.nombre}» eliminado · ${c.montoLine} 🚀` },
  { titulo: '🧸 Con cariño, se fue', cuerpo: (c) => `«${c.nombre}» ya no está (${c.montoLine}) · abrazo al que decide 🧸` },
  { titulo: '✅ Tachado del mundo digital', cuerpo: (c) => `«${c.nombre}» quitado · ${c.montoLine} ✅` },
  { titulo: '🍕 Rebanada menos', cuerpo: (c) => `«${c.nombre}» salió del pastel (${c.montoLine}) 🍕` },
  { titulo: '🌙 Luna de calma', cuerpo: (c) => `«${c.nombre}» borrado (${c.montoLine}) · noche más tranquila 🌙` },
  { titulo: '🧩 Pieza que sobraba', cuerpo: (c) => `«${c.nombre}» fuera del puzzle · ${c.montoLine} 🧩` },
  { titulo: '🎒 Mochila más liviana', cuerpo: (c) => `«${c.nombre}» eliminado (${c.montoLine}) 🎒` },
  { titulo: '💫 Polvo de estrellas', cuerpo: (c) => `«${c.nombre}» se desvaneció del historial · ${c.montoLine} 💫` },
];

export function varianteGastoEditadoCampana(ctx) {
  return applyPair(pick(GASTO_EDITADO_CAMPANA), ctx);
}

export function varianteGastoEliminadoCampana(ctx) {
  return applyPair(pick(GASTO_ELIMINADO_CAMPANA), ctx);
}

/** Lista súper urgente: 20 pares con emoji (export para usar desde notificacionesLocalesListaSuper) */
export const AVISOS_LISTA_SUPER_URGENTE_CON_EMOJI = [
  { titulo: '🏠 ¿Miramos la despensa?', cuerpo: (r) => `Falta lo urgente: ${r} 🛒 Un detalle que alegra el hogar ✨` },
  { titulo: '📝 Recadito hogareño', cuerpo: (r) => `${r} sigue en rojo en la lista 😊 Cuando puedas, pa fuera del pendiente.` },
  { titulo: '🛒 ¿Quién va al súper?', cuerpo: (r) => `Toque urgentito: ${r} 🧊 La nevera manda besos.` },
  { titulo: '👋 Hola desde tu lista', cuerpo: (r) => `No olvides: ${r} ⭐ Prioridad casita.` },
  { titulo: '⏰ Antes de que vuele', cuerpo: (r) => `${r} está en urgente 🎯 Buen día para tacharlo.` },
  { titulo: '🔮 Tu yo de mañana aplaude', cuerpo: (r) => `Si hoy cae ${r} 🙌 el futuro tú sonríe.` },
  { titulo: '🏡 Modo casa ON', cuerpo: (r) => `Pasa por ${r} cuando salgas 🚶 Pequeño esfuerzo, gran paz.` },
  { titulo: '😌 Sin drama', cuerpo: (r) => `Pendiente urgente: ${r} 📌 Sin jefe respirando en la nuca.` },
  { titulo: '🛍️ La lista sonríe', cuerpo: (r) => `${r} te espera en el mercado 🥬 Con buena onda.` },
  { titulo: '☕ Café y recado', cuerpo: (r) => `Entre taza y taza: ${r} ☕ sigue urgente.` },
  { titulo: '🤗 Empujoncito', cuerpo: (r) => `${r} — para que no falte en casa 🤗` },
  { titulo: '🏡 Hogar 100', cuerpo: (r) => `Con ${r} completo el hogar rima mejor 🎵` },
  { titulo: '🎯 Misión súper', cuerpo: (r) => `Objetivo: ${r} 🎯 Nivel: urgente con calma.` },
  { titulo: '🧊 La nevera, en off', cuerpo: (r) => `Susurra que falta ${r} 🧊 (o era hambre, igual ayuda).` },
  { titulo: '✨ Casi magia', cuerpo: (r) => `Tachar ${r} es magia menor ✨ ¿Hoy?` },
  { titulo: '🚶 Sin prisa absurda', cuerpo: (r) => `Urgente: ${r} 🚶 Cuando salgas, ya está en la mochila mental.` },
  { titulo: '📋 Mini memo', cuerpo: (r) => `${r} · urgente · lista MoneyTrack 📋` },
  { titulo: '💪 Tú puedes', cuerpo: (r) => `Ir por ${r} es victoria pequeña 💪` },
  { titulo: '👨‍👩‍👧 Equipo casa', cuerpo: (r) => `A veces falta ${r} 💛 Aviso con cariño.` },
  { titulo: '🎉 Cuando lo compres…', cuerpo: (r) => `…${r} sale de urgente y mereces mini fiesta 🎉` },
];

export function varianteListaSuperUrgente(resumen) {
  const pair = pick(AVISOS_LISTA_SUPER_URGENTE_CON_EMOJI);
  return {
    title: pair.titulo,
    body: pair.cuerpo(resumen),
  };
}
