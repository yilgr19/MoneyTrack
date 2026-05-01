/**
 * Ejecuta `adb reverse` en cada dispositivo en estado `device` (evita
 * "more than one device/emulator" cuando hay Wi‑Fi + TLS duplicados).
 *
 * Uso:
 *   node scripts/adb-reverse.cjs           → falla con código 1 si no hay dispositivos
 *   node scripts/adb-reverse.cjs --soft    → sin dispositivo / sin adb: sale 0 (solo aviso)
 */
const { execSync } = require('child_process');

const soft = process.argv.includes('--soft');

function listDeviceSerials() {
  const out = execSync('adb devices', { encoding: 'utf8', stdio: ['pipe', 'pipe', 'pipe'] });
  const serials = [];
  for (const line of out.split(/\r?\n/)) {
    const m = line.match(/^(\S+)\s+device\s*$/);
    if (m) serials.push(m[1]);
  }
  return serials;
}

let serials;
try {
  serials = listDeviceSerials();
} catch (e) {
  if (soft) {
    console.error('[MoneyTrack] adb reverse omitido: adb no disponible o error al listar dispositivos.');
    process.exit(0);
  }
  throw e;
}

if (serials.length === 0) {
  if (soft) {
    console.error(
      '[MoneyTrack] adb reverse omitido: ningún dispositivo en estado "device". Conecta USB o Wi‑Fi (adb connect IP:PUERTO).',
    );
    process.exit(0);
  }
  console.error('[MoneyTrack] No hay ningún dispositivo adb en estado "device". Ejecuta adb devices.');
  process.exit(1);
}

for (const serial of serials) {
  console.error(`[MoneyTrack] adb reverse → ${serial}`);
  execSync(`adb -s "${serial}" reverse tcp:8081 tcp:8081`, { stdio: 'inherit', shell: true });
  execSync(`adb -s "${serial}" reverse tcp:8082 tcp:8082`, { stdio: 'inherit', shell: true });
}
