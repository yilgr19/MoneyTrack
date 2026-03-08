/**
 * Importa datos desde un archivo Excel (plantilla MoneyTrack).
 * Reconoce la estructura exportada y rellena el sistema automáticamente.
 */
function importarDesdeExcel(archivo) {
    if (typeof XLSX === 'undefined') {
        alert('Error: La librería Excel no está cargada. Recarga la página.');
        return;
    }
    if (!archivo || !archivo.name) {
        alert('Selecciona un archivo Excel.');
        return;
    }
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const MESES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];

            const parseNum = (v) => {
                if (v == null || v === '') return 0;
                if (typeof v === 'number' && !isNaN(v)) return v;
                const s = String(v).replace(/\s+[A-Z]{3}$/i, '').replace(/\./g, '').replace(',', '.');
                return parseFloat(s) || 0;
            };
            const parseFecha = (v) => {
                if (!v) return null;
                if (typeof v === 'number') {
                    const d = XLSX.SSF.parse_date_code ? XLSX.SSF.parse_date_code(v) : null;
                    if (d) return new Date(d.y, d.m - 1, d.d, d.H || 0, d.M || 0);
                    return new Date((v - 25569) * 86400 * 1000);
                }
                const m = String(v).match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})(?:\s+(\d{1,2}):(\d{2}))?/);
                if (m) {
                    const d = parseInt(m[3], 10), mes = parseInt(m[2], 10) - 1, dia = parseInt(m[1], 10);
                    const h = m[4] ? parseInt(m[4], 10) : 12, min = m[5] ? parseInt(m[5], 10) : 0;
                    return new Date(d, mes, dia, h, min);
                }
                const mesMatch = MESES.findIndex(m => String(v).includes(m));
                if (mesMatch >= 0) {
                    const año = parseInt(String(v).match(/\d{4}/)?.[0] || new Date().getFullYear(), 10);
                    return new Date(año, mesMatch, 1, 12, 0);
                }
                return new Date(v);
            };
            const fechaToStr = (d) => {
                if (!d || !(d instanceof Date) || isNaN(d)) return null;
                const p = n => String(n).padStart(2, '0');
                return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}:00`;
            };
            const cuentaNombreToId = (nombre) => {
                const n = (nombre || '').trim().toLowerCase();
                const map = { 'efectivo':'efectivo','banco':'banco','tarjeta de crédito':'tarjetaCredito','nequi':'nequi','daviplata':'daviplata' };
                return map[n] || CUENTAS.find(c => c.nombre.toLowerCase() === n)?.id || 'efectivo';
            };

            let moneda = localStorage.getItem('moneda') || '';

            const getSheet = (nombre) => {
                const ws = workbook.Sheets[workbook.SheetNames.find(n => n === nombre)];
                return ws ? XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }) : [];
            };
            const getCell = (arr, r, c) => (arr[r] && arr[r][c] != null) ? arr[r][c] : '';

            const datos = getSheet('Resumen');
            for (let r = 0; r < (datos.length || 0); r++) {
                const a = String(getCell(datos, r, 0) || '').trim();
                const b = getCell(datos, r, 1);
                if (a === 'Moneda' && b) moneda = String(b).trim().split(/\s/)[0] || moneda;
                if (a === 'Presupuesto mensual' && b && String(b) !== 'No definido') {
                    const num = parseNum(b);
                    if (num > 0) localStorage.setItem('presupuestoMensual', num.toString());
                }
                if (a === 'Límite tarjeta de crédito' && b && String(b) !== 'No definido') {
                    const num = parseNum(b);
                    if (num > 0) localStorage.setItem('limiteTarjetaCredito', num.toString());
                }
                if (a === 'SALDOS INICIALES POR CUENTA') {
                    for (let i = r + 1; i < datos.length; i++) {
                        const nom = String(getCell(datos, i, 0) || '').trim();
                        const val = parseNum(getCell(datos, i, 1));
                        if (nom && !nom.startsWith('SALDO') && !nom.startsWith('TOTAL')) {
                            const id = cuentaNombreToId(nom);
                            const saldos = JSON.parse(localStorage.getItem('saldosCuentas') || '{}');
                            saldos[id] = val;
                            localStorage.setItem('saldosCuentas', JSON.stringify(saldos));
                        }
                        if (nom === 'SALDO ACTUAL' || nom === 'Saldo actual') break;
                    }
                    break;
                }
            }
            if (moneda) localStorage.setItem('moneda', moneda);

            const rowsIng = getSheet('Ingresos');
            const ingresos = [];
            for (let r = 1; r < (rowsIng.length || 0); r++) {
                const fechaVal = getCell(rowsIng, r, 1);
                const concepto = getCell(rowsIng, r, 2);
                const monto = parseNum(getCell(rowsIng, r, 3));
                const cuenta = getCell(rowsIng, r, 4);
                if (String(concepto || '').includes('═══') || String(concepto || '').includes('Sin ingresos')) continue;
                if (!monto && !concepto) continue;
                const d = parseFecha(fechaVal) || parseFecha(getCell(rowsIng, r, 0));
                if (d) {
                    ingresos.push({
                        cantidad: monto || 0,
                        fecha: fechaToStr(d),
                        origen: cuentaNombreToId(cuenta),
                        nota: (concepto || 'Ingreso').toString().trim()
                    });
                }
            }
            if (ingresos.length) localStorage.setItem('ingresos', JSON.stringify(ingresos));

            const rowsGas = getSheet('Gastos');
            const gastos = [];
            for (let r = 1; r < (rowsGas.length || 0); r++) {
                const fechaVal = getCell(rowsGas, r, 1);
                const nombre = getCell(rowsGas, r, 2);
                const categoria = getCell(rowsGas, r, 3);
                const monto = parseNum(getCell(rowsGas, r, 4));
                const cuenta = getCell(rowsGas, r, 5);
                const nota = getCell(rowsGas, r, 6);
                const cuotasStr = getCell(rowsGas, r, 7);
                if (String(nombre || '').includes('═══') || String(nombre || '').includes('Sin gastos')) continue;
                if (!monto && !nombre) continue;
                const d = parseFecha(fechaVal) || parseFecha(getCell(rowsGas, r, 0));
                if (d) {
                    const cuotas = parseInt(String(cuotasStr || '1').replace(/\D/g, '') || '1', 10) || 1;
                    gastos.push({
                        nombre: (nombre || '').toString().trim(),
                        cantidad: monto || 0,
                        fecha: fechaToStr(d),
                        categoria: (categoria || '').toString().trim(),
                        origen: cuentaNombreToId(cuenta),
                        nota: (nota || '').toString().trim() || null,
                        cuotas: cuotas,
                        cuotaMensual: cuotas > 1 ? (monto || 0) / cuotas : (monto || 0)
                    });
                }
            }
            if (gastos.length) localStorage.setItem('gastos', JSON.stringify(gastos));

            const rowsCat = getSheet('Categorías');
            const categorias = [];
            for (let r = 1; r < (rowsCat.length || 0); r++) {
                const nom = getCell(rowsCat, r, 0);
                const limite = parseNum(getCell(rowsCat, r, 1));
                const color = getCell(rowsCat, r, 2);
                if (String(nom || '').includes('Sin categorías')) continue;
                if (nom) categorias.push({ nombre: String(nom).trim(), limite: limite || null, color: (color || '').toString().trim() || undefined });
            }
            if (categorias.length) localStorage.setItem('categorias', JSON.stringify(categorias));

            const rowsMet = getSheet('Metas');
            const metas = [];
            const metaNombreToId = {};
            for (let r = 1; r < (rowsMet.length || 0); r++) {
                const nom = getCell(rowsMet, r, 0);
                const obj = parseNum(getCell(rowsMet, r, 1));
                const plazo = getCell(rowsMet, r, 4);
                if (String(nom || '').includes('Sin metas')) continue;
                if (nom) {
                    const id = 'meta_' + Date.now() + '_' + r;
                    metaNombreToId[String(nom).trim()] = id;
                    metas.push({ id, nombre: String(nom).trim(), objetivo: obj || 0, plazo: (plazo || '').toString().trim() || null });
                }
            }
            if (metas.length) localStorage.setItem('metas', JSON.stringify(metas));

            const rowsAport = getSheet('Aportes a metas');
            const contribuciones = [];
            for (let r = 1; r < (rowsAport.length || 0); r++) {
                const fechaVal = getCell(rowsAport, r, 1);
                const metaNom = getCell(rowsAport, r, 2);
                const monto = parseNum(getCell(rowsAport, r, 3));
                const cuenta = getCell(rowsAport, r, 4);
                if (String(metaNom || '').includes('═══') || String(metaNom || '').includes('Sin aportes')) continue;
                if (!monto && !metaNom) continue;
                const metaId = metaNombreToId[String(metaNom).trim()] || Object.values(metaNombreToId)[0];
                const d = parseFecha(fechaVal) || parseFecha(getCell(rowsAport, r, 0));
                if (d && metaId) {
                    contribuciones.push({ metaId, cantidad: monto, fecha: fechaToStr(d), origen: cuentaNombreToId(cuenta) });
                }
            }
            if (contribuciones.length) localStorage.setItem('contribucionesMetas', JSON.stringify(contribuciones));

            const rowsPagos = getSheet('Pagos programados');
            const pagos = [];
            for (let r = 1; r < (rowsPagos.length || 0); r++) {
                const concepto = getCell(rowsPagos, r, 0);
                const monto = parseNum(getCell(rowsPagos, r, 1));
                const frecuencia = getCell(rowsPagos, r, 2);
                const diaPago = getCell(rowsPagos, r, 3);
                const cuenta = getCell(rowsPagos, r, 4);
                const categoria = getCell(rowsPagos, r, 5);
                const fechaInicio = getCell(rowsPagos, r, 6);
                const activo = getCell(rowsPagos, r, 7);
                if (String(concepto || '').includes('Sin pagos')) continue;
                if (!concepto && !monto) continue;
                const d = parseFecha(fechaInicio);
                pagos.push({
                    id: 'pago_' + Date.now() + '_' + r,
                    concepto: (concepto || '').toString().trim(),
                    monto: monto || 0,
                    frecuencia: (frecuencia || 'mensual').toString().trim().toLowerCase(),
                    diaPago: parseInt(diaPago, 10) || 1,
                    cuenta: cuentaNombreToId(cuenta),
                    categoria: (categoria || '').toString().trim(),
                    fechaInicio: d ? fechaToStr(d).slice(0, 10) : null,
                    activo: String(activo || '').toLowerCase() !== 'no'
                });
            }
            if (pagos.length) localStorage.setItem('pagosProgramados', JSON.stringify(pagos));

            alert('¡Importación completada! Se han cargado: ' +
                (moneda ? 'Configuración, ' : '') +
                (ingresos.length ? ingresos.length + ' ingresos, ' : '') +
                (gastos.length ? gastos.length + ' gastos, ' : '') +
                (categorias.length ? categorias.length + ' categorías, ' : '') +
                (metas.length ? metas.length + ' metas, ' : '') +
                (contribuciones.length ? contribuciones.length + ' aportes, ' : '') +
                (pagos.length ? pagos.length + ' pagos programados.' : ''));
            location.reload();
        } catch (err) {
            console.error(err);
            alert('Error al importar: ' + (err.message || 'Formato de archivo no reconocido. Asegúrate de usar una plantilla exportada desde MoneyTrack.'));
        }
    };
    reader.readAsArrayBuffer(archivo);
}
