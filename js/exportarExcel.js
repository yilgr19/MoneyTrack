/**
 * Exporta toda la información de MoneyTrack a un archivo Excel (.xlsx)
 * Con tablas estilizadas, títulos resaltados en colores y diseño estético.
 */
function exportarAExcel() {
    if (typeof XLSX === 'undefined') {
        alert('Error: La librería Excel no está cargada. Recarga la página e intenta de nuevo.');
        return;
    }

    const moneda = localStorage.getItem('moneda') || '';
    const MESES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];

    // Estilos reutilizables
    const estilos = {
        tituloPrincipal: { font: { bold: true, sz: 16, color: { rgb: "2E5090" } }, fill: { fgColor: { rgb: "E2EFDA" }, patternType: "solid" }, alignment: { horizontal: "center" } },
        tituloSeccion: { font: { bold: true, sz: 12, color: { rgb: "FFFFFF" } }, fill: { fgColor: { rgb: "4472C4" }, patternType: "solid" } },
        headerTabla: { font: { bold: true, sz: 11, color: { rgb: "FFFFFF" } }, fill: { fgColor: { rgb: "5B9BD5" }, patternType: "solid" }, alignment: { horizontal: "center", wrapText: true } },
        mesHeader: { font: { bold: true, sz: 11, color: { rgb: "2E5090" } }, fill: { fgColor: { rgb: "D6DCE4" }, patternType: "solid" } },
        filaPar: { fill: { fgColor: { rgb: "F2F2F2" }, patternType: "solid" } },
        filaImpar: { fill: { fgColor: { rgb: "FFFFFF" }, patternType: "solid" } },
        filaTotal: { font: { bold: true }, fill: { fgColor: { rgb: "FFF2CC" }, patternType: "solid" } },
        borde: { border: { top: { style: "thin", color: { rgb: "B4B4B4" } }, bottom: { style: "thin", color: { rgb: "B4B4B4" } }, left: { style: "thin", color: { rgb: "B4B4B4" } }, right: { style: "thin", color: { rgb: "B4B4B4" } } } }
    };

    const celda = (v, tipo, estilo) => {
        const c = { v, t: tipo || (typeof v === 'number' ? 'n' : 's') };
        if (estilo) c.s = { ...estilo, ...(estilos.borde || {}) };
        return c;
    };
    const h = (v) => celda(v, 's', estilos.headerTabla);
    const sec = (v) => celda(v, 's', estilos.tituloSeccion);
    const mes = (v) => celda(v, 's', estilos.mesHeader);
    const dato = (v, tipo, par) => celda(v, tipo, par ? estilos.filaPar : estilos.filaImpar);
    const total = (v, tipo) => celda(v, tipo, estilos.filaTotal);

    const saldosIni = obtenerSaldosIniciales();
    const saldosActuales = calcularSaldosPorCuenta();
    const ingresos = JSON.parse(localStorage.getItem('ingresos') || '[]');
    const gastos = JSON.parse(localStorage.getItem('gastos') || '[]');
    const categorias = JSON.parse(localStorage.getItem('categorias') || '[]');
    const metas = JSON.parse(localStorage.getItem('metas') || '[]');
    const contribuciones = JSON.parse(localStorage.getItem('contribucionesMetas') || '[]');
    const pagosProgramados = JSON.parse(localStorage.getItem('pagosProgramados') || '[]');
    const presupuestoMensual = parseFloat(localStorage.getItem('presupuestoMensual')) || 0;
    const limiteTc = parseFloat(localStorage.getItem('limiteTarjetaCredito')) || 0;

    const totalIngresos = ingresos.reduce((s,i)=>s+i.cantidad,0);
    const totalGastos = gastos.reduce((s,g)=>s+montoGastoAfectaSaldo(g),0);
    const totalAportadoMetas = contribuciones.reduce((s,c)=>s+c.cantidad,0);

    const normalizarCat = c => typeof c === 'string' ? { nombre: c } : { nombre: c.nombre || c };
    const nombreCuenta = id => CUENTAS.find(c => c.id === id)?.nombre || id;
    const nombreMeta = id => metas.find(m => m.id === id)?.nombre || id;

    const workbook = XLSX.utils.book_new();

    // ========== HOJA 1: Resumen ==========
    const resumenData = [
        [celda('RESUMEN MONEYTRACK', 's', estilos.tituloPrincipal), celda('', 's')],
        [celda('Fecha de exportación', 's'), celda(new Date().toLocaleString('es'), 's')],
        [celda('', 's'), celda('', 's')],
        [sec('CONFIGURACIÓN'), sec('')],
        [celda('Moneda', 's'), celda(moneda, 's')],
        [celda('Presupuesto mensual', 's'), celda(presupuestoMensual > 0 ? formatearNumero(presupuestoMensual) + ' ' + moneda : 'No definido', 's')],
        [celda('Límite tarjeta de crédito', 's'), celda(limiteTc > 0 ? formatearNumero(limiteTc) + ' ' + moneda : 'No definido', 's')],
        [celda('', 's'), celda('', 's')],
        [sec('SALDOS INICIALES POR CUENTA'), sec('')],
        ...CUENTAS.map((c, i) => [celda(c.nombre, 's'), celda(formatearNumero(saldosIni[c.id] || 0) + ' ' + moneda, 's')]),
        [celda('', 's'), celda('', 's')],
        [sec('SALDO ACTUAL POR CUENTA'), sec('')],
        ...CUENTAS.filter(c => (saldosActuales[c.id] || 0) !== 0).map(c => [celda(c.nombre, 's'), celda(formatearNumero(saldosActuales[c.id]) + ' ' + moneda, 's')]),
        [total('Saldo total disponible', 's'), total(formatearNumero(saldosActuales.total || 0) + ' ' + moneda, 's')],
        [celda('', 's'), celda('', 's')],
        [sec('TOTALES'), sec('')],
        [celda('Total ingresos', 's'), celda(formatearNumero(totalIngresos) + ' ' + moneda, 's')],
        [celda('Total gastos', 's'), celda(formatearNumero(totalGastos) + ' ' + moneda, 's')],
        [total('Total aportado a metas', 's'), total(formatearNumero(totalAportadoMetas) + ' ' + moneda, 's')],
    ];
    const wsResumen = XLSX.utils.aoa_to_sheet(resumenData);
    wsResumen['!cols'] = [{ wch: 28 }, { wch: 22 }];
    wsResumen['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }]; // Título centrado en 2 columnas
    XLSX.utils.book_append_sheet(workbook, wsResumen, 'Resumen');

    // ========== HOJA 2: Ingresos ==========
    const ingresosOrdenados = [...ingresos].sort((a,b) => new Date(a.fecha) - new Date(b.fecha));
    const ingresosRows = [[h('MES'), h('FECHA'), h('CONCEPTO / NOTA'), h('MONTO (' + moneda + ')'), h('CUENTA')]];
    let mesActualIng = null;
    let idxIng = 0;
    ingresosOrdenados.forEach(i => {
        const d = new Date(i.fecha && i.fecha.includes('T') ? i.fecha : i.fecha + 'T12:00:00');
        const mesKey = d.getFullYear() + '-' + d.getMonth();
        if (mesActualIng !== mesKey) {
            mesActualIng = mesKey;
            ingresosRows.push([
                mes(''), mes(''), mes('═══ ' + MESES[d.getMonth()] + ' ' + d.getFullYear() + ' ═══'), mes(''), mes('')
            ]);
        }
        idxIng++;
        const par = idxIng % 2 === 0;
        ingresosRows.push([
            dato(MESES[d.getMonth()] + ' ' + d.getFullYear(), 's', par),
            dato(d.toLocaleDateString('es') + (i.fecha && i.fecha.includes('T') ? ' ' + d.toLocaleTimeString('es', { hour: '2-digit', minute: '2-digit' }) : ''), 's', par),
            dato(i.nota || 'Ingreso', 's', par),
            dato(i.cantidad, 'n', par),
            dato(nombreCuenta(i.origen), 's', par)
        ]);
    });
    if (ingresosRows.length === 1) ingresosRows.push([dato('Sin ingresos registrados', 's', false), dato('', 's'), dato('', 's'), dato('', 's'), dato('', 's')]);
    const wsIngresos = XLSX.utils.aoa_to_sheet(ingresosRows);
    wsIngresos['!cols'] = [{ wch: 14 }, { wch: 20 }, { wch: 32 }, { wch: 14 }, { wch: 18 }];
    XLSX.utils.book_append_sheet(workbook, wsIngresos, 'Ingresos');

    // ========== HOJA 3: Gastos ==========
    const gastosOrdenados = [...gastos].sort((a,b) => new Date(a.fecha) - new Date(b.fecha));
    const gastosRows = [[h('MES'), h('FECHA'), h('CONCEPTO'), h('CATEGORÍA'), h('MONTO (' + moneda + ')'), h('CUENTA'), h('NOTA'), h('CUOTAS')]];
    let mesActualGas = null;
    let idxGas = 0;
    gastosOrdenados.forEach(g => {
        const d = new Date(g.fecha && g.fecha.includes('T') ? g.fecha : g.fecha + 'T12:00:00');
        const mesKey = d.getFullYear() + '-' + d.getMonth();
        if (mesActualGas !== mesKey) {
            mesActualGas = mesKey;
            gastosRows.push([
                mes(''), mes(''), mes('═══ ' + MESES[d.getMonth()] + ' ' + d.getFullYear() + ' ═══'), mes(''), mes(''), mes(''), mes(''), mes('')
            ]);
        }
        idxGas++;
        const par = idxGas % 2 === 0;
        gastosRows.push([
            dato(MESES[d.getMonth()] + ' ' + d.getFullYear(), 's', par),
            dato(d.toLocaleDateString('es') + (g.fecha && g.fecha.includes('T') ? ' ' + d.toLocaleTimeString('es', { hour: '2-digit', minute: '2-digit' }) : ''), 's', par),
            dato(g.nombre || '', 's', par),
            dato(g.categoria || '', 's', par),
            dato(g.cantidad, 'n', par),
            dato(nombreCuenta(g.origen), 's', par),
            dato(g.nota || '', 's', par),
            dato(g.cuotas > 1 ? g.cuotas + ' cuotas' : '1', 's', par)
        ]);
    });
    if (gastosRows.length === 1) gastosRows.push([dato('Sin gastos registrados', 's', false), dato('', 's'), dato('', 's'), dato('', 's'), dato('', 's'), dato('', 's'), dato('', 's'), dato('', 's')]);
    const wsGastos = XLSX.utils.aoa_to_sheet(gastosRows);
    wsGastos['!cols'] = [{ wch: 14 }, { wch: 20 }, { wch: 24 }, { wch: 16 }, { wch: 12 }, { wch: 18 }, { wch: 22 }, { wch: 10 }];
    XLSX.utils.book_append_sheet(workbook, wsGastos, 'Gastos');

    // ========== HOJA 4: Categorías ==========
    const catRows = [[h('CATEGORÍA'), h('LÍMITE MENSUAL (' + moneda + ')'), h('COLOR')]];
    categorias.map(normalizarCat).forEach((c, i) => {
        const catCompleta = categorias.find(x => (typeof x === 'string' ? x : x.nombre) === c.nombre);
        const limite = catCompleta && catCompleta.limite ? catCompleta.limite : '';
        const color = catCompleta && catCompleta.color ? catCompleta.color : '';
        const par = (i + 1) % 2 === 0;
        catRows.push([dato(c.nombre, 's', par), dato(limite, 'n', par), dato(color, 's', par)]);
    });
    if (catRows.length === 1) catRows.push([dato('Sin categorías creadas', 's', false), dato('', 's'), dato('', 's')]);
    const wsCategorias = XLSX.utils.aoa_to_sheet(catRows);
    wsCategorias['!cols'] = [{ wch: 22 }, { wch: 20 }, { wch: 12 }];
    XLSX.utils.book_append_sheet(workbook, wsCategorias, 'Categorías');

    // ========== HOJA 5: Metas ==========
    const metasRows = [[h('META'), h('OBJETIVO (' + moneda + ')'), h('AHORRADO'), h('% LOGRADO'), h('PLAZO'), h('Nº APORTES')]];
    metas.forEach((m, i) => {
        const aportado = contribuciones.filter(c => c.metaId === m.id).reduce((s, c) => s + c.cantidad, 0);
        const obj = m.objetivo || 0;
        const pct = obj > 0 ? Math.min(100, Math.round((aportado / obj) * 100)) : (aportado > 0 ? 100 : 0);
        const par = (i + 1) % 2 === 0;
        metasRows.push([
            dato(m.nombre || '', 's', par),
            dato(obj, 'n', par),
            dato(aportado, 'n', par),
            dato(pct + '%', 's', par),
            dato(m.plazo || '—', 's', par),
            dato(contribuciones.filter(c => c.metaId === m.id).length, 'n', par)
        ]);
    });
    if (metasRows.length === 1) metasRows.push([dato('Sin metas creadas', 's', false), dato('', 's'), dato('', 's'), dato('', 's'), dato('', 's'), dato('', 's')]);
    const wsMetas = XLSX.utils.aoa_to_sheet(metasRows);
    wsMetas['!cols'] = [{ wch: 24 }, { wch: 14 }, { wch: 12 }, { wch: 10 }, { wch: 14 }, { wch: 10 }];
    XLSX.utils.book_append_sheet(workbook, wsMetas, 'Metas');

    // ========== HOJA 6: Aportes a metas ==========
    const aportesOrdenados = [...contribuciones].sort((a,b) => new Date(a.fecha) - new Date(b.fecha));
    const aportesRows = [[h('MES'), h('FECHA'), h('META'), h('MONTO (' + moneda + ')'), h('CUENTA')]];
    let mesActualAport = null;
    let idxAport = 0;
    aportesOrdenados.forEach(c => {
        const d = new Date(c.fecha && c.fecha.includes('T') ? c.fecha : c.fecha + 'T12:00:00');
        const mesKey = d.getFullYear() + '-' + d.getMonth();
        if (mesActualAport !== mesKey) {
            mesActualAport = mesKey;
            aportesRows.push([mes(''), mes(''), mes('═══ ' + MESES[d.getMonth()] + ' ' + d.getFullYear() + ' ═══'), mes(''), mes('')]);
        }
        idxAport++;
        const par = idxAport % 2 === 0;
        aportesRows.push([
            dato(MESES[d.getMonth()] + ' ' + d.getFullYear(), 's', par),
            dato(d.toLocaleDateString('es'), 's', par),
            dato(nombreMeta(c.metaId), 's', par),
            dato(c.cantidad, 'n', par),
            dato(nombreCuenta(c.origen), 's', par)
        ]);
    });
    if (aportesRows.length === 1) aportesRows.push([dato('Sin aportes registrados', 's', false), dato('', 's'), dato('', 's'), dato('', 's'), dato('', 's')]);
    const wsAportes = XLSX.utils.aoa_to_sheet(aportesRows);
    wsAportes['!cols'] = [{ wch: 14 }, { wch: 12 }, { wch: 24 }, { wch: 14 }, { wch: 18 }];
    XLSX.utils.book_append_sheet(workbook, wsAportes, 'Aportes a metas');

    // ========== HOJA 7: Pagos programados ==========
    const pagosRows = [[h('CONCEPTO'), h('MONTO (' + moneda + ')'), h('FRECUENCIA'), h('DÍA PAGO'), h('CUENTA'), h('CATEGORÍA'), h('FECHA INICIO'), h('ACTIVO')]];
    pagosProgramados.forEach((p, i) => {
        const par = (i + 1) % 2 === 0;
        pagosRows.push([
            dato(p.concepto || '', 's', par),
            dato(p.monto || 0, 'n', par),
            dato(p.frecuencia || '', 's', par),
            dato(p.diaPago || '', 's', par),
            dato(nombreCuenta(p.cuenta), 's', par),
            dato(p.categoria || '', 's', par),
            dato(p.fechaInicio ? p.fechaInicio.slice(0, 10) : '', 's', par),
            dato(p.activo !== false ? 'Sí' : 'No', 's', par)
        ]);
    });
    if (pagosRows.length === 1) pagosRows.push([dato('Sin pagos programados', 's', false), dato('', 's'), dato('', 's'), dato('', 's'), dato('', 's'), dato('', 's'), dato('', 's'), dato('', 's')]);
    const wsPagos = XLSX.utils.aoa_to_sheet(pagosRows);
    wsPagos['!cols'] = [{ wch: 22 }, { wch: 12 }, { wch: 12 }, { wch: 10 }, { wch: 18 }, { wch: 16 }, { wch: 12 }, { wch: 8 }];
    XLSX.utils.book_append_sheet(workbook, wsPagos, 'Pagos programados');

    const nombreArchivo = 'MoneyTrack_' + new Date().toISOString().slice(0, 10) + '.xlsx';
    XLSX.writeFile(workbook, nombreArchivo);
}

/** Convierte nombre de cuenta a id (Efectivo -> efectivo, Tarjeta de crédito -> tarjetaCredito) */
function cuentaNombreToId(nombre) {
    if (!nombre || typeof nombre !== 'string') return '';
    const n = nombre.trim().toLowerCase();
    const map = { 'efectivo': 'efectivo', 'banco': 'banco', 'tarjeta de crédito': 'tarjetaCredito', 'nequi': 'nequi', 'daviplata': 'daviplata' };
    return map[n] || CUENTAS.find(c => c.nombre.toLowerCase() === n)?.id || '';
}

/** Parsea número desde Excel (puede ser número o "1.000,00 COP") */
function parseNumExcel(val) {
    if (val == null || val === '') return 0;
    if (typeof val === 'number' && !isNaN(val)) return val;
    const s = String(val).replace(/\s+[A-Z]{3}$/i, '').trim().replace(/\./g, '').replace(',', '.');
    return parseFloat(s) || 0;
}

/** Parsea fecha desde Excel (serial o string dd/mm/yyyy) */
function parseFechaExcel(val) {
    if (!val) return null;
    if (val instanceof Date) return val;
    if (typeof val === 'number') {
        const d = new Date((val - 25569) * 86400 * 1000);
        return isNaN(d.getTime()) ? null : d;
    }
    const s = String(val).trim();
    const m = s.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:\s+(\d{1,2}):(\d{2}))?/);
    if (m) {
        const d = new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1]), parseInt(m[4]) || 12, parseInt(m[5]) || 0);
        return isNaN(d.getTime()) ? null : d;
    }
    const d2 = new Date(s);
    return isNaN(d2.getTime()) ? null : d2;
}

/** Obtiene valor de celda desde hoja (por fila/col) */
function celdaVal(ws, r, c) {
    const col = String.fromCharCode(65 + c);
    const ref = col + (r + 1);
    const cell = ws[ref];
    return cell ? (cell.v !== undefined ? cell.v : cell.w) : '';
}

/** Importa datos desde archivo Excel (plantilla MoneyTrack) */
function importarDesdeExcel(archivo) {
    if (typeof XLSX === 'undefined') {
        alert('Error: La librería Excel no está cargada.');
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

            let moneda = '';
            let presupuestoMensual = 0;
            let limiteTc = 0;
            const saldosIni = { efectivo: 0, banco: 0, tarjetaCredito: 0, nequi: 0, daviplata: 0 };

            if (workbook.SheetNames.includes('Resumen')) {
                const ws = workbook.Sheets['Resumen'];
                const filas = XLSX.utils.sheet_to_json(ws, { header: 1 });
                let enSeccionSaldosIniciales = false;
                filas.forEach((fila, i) => {
                    const a = String(fila[0] || '').trim();
                    const b = fila[1];
                    if (a === 'Moneda' && b) moneda = String(b).trim().split(/\s/)[0] || '';
                    if (a === 'Presupuesto mensual' && b) presupuestoMensual = parseNumExcel(b);
                    if (a === 'Límite tarjeta de crédito' && b) limiteTc = parseNumExcel(b);
                    if (a === 'SALDOS INICIALES POR CUENTA') enSeccionSaldosIniciales = true;
                    if (a === 'SALDO ACTUAL POR CUENTA' || a === 'Saldo total disponible') enSeccionSaldosIniciales = false;
                    if (enSeccionSaldosIniciales && ['Efectivo','Banco','Tarjeta de crédito','Nequi','Daviplata'].includes(a)) {
                        const id = cuentaNombreToId(a);
                        if (id) saldosIni[id] = parseNumExcel(b);
                    }
                });
            }

            let ingresos = [];
            if (workbook.SheetNames.includes('Ingresos')) {
                const ws = workbook.Sheets['Ingresos'];
                const filas = XLSX.utils.sheet_to_json(ws, { header: 1 });
                for (let r = 1; r < filas.length; r++) {
                    const f = filas[r];
                    const concepto = String(f[2] || '').trim();
                    const montoVal = f[3];
                    const cuentaNom = String(f[4] || '').trim();
                    if (!concepto || concepto.startsWith('═══') || concepto === 'Sin ingresos registrados') continue;
                    const monto = parseNumExcel(montoVal);
                    if (monto <= 0) continue;
                    const fecha = parseFechaExcel(f[1]);
                    if (!fecha) continue;
                    const pad = n => String(n).padStart(2, '0');
                    const fechaStr = `${fecha.getFullYear()}-${pad(fecha.getMonth()+1)}-${pad(fecha.getDate())}T${pad(fecha.getHours())}:${pad(fecha.getMinutes())}`;
                    const cuentaId = cuentaNombreToId(cuentaNom) || 'efectivo';
                    ingresos.push({ cantidad: monto, fecha: fechaStr, origen: cuentaId, nota: concepto || 'Ingreso' });
                }
            }

            let gastos = [];
            if (workbook.SheetNames.includes('Gastos')) {
                const ws = workbook.Sheets['Gastos'];
                const filas = XLSX.utils.sheet_to_json(ws, { header: 1 });
                for (let r = 1; r < filas.length; r++) {
                    const f = filas[r];
                    const concepto = String(f[2] || '').trim();
                    if (!concepto || concepto.startsWith('═══') || concepto === 'Sin gastos registrados') continue;
                    const monto = parseNumExcel(f[4]);
                    if (monto <= 0) continue;
                    const fecha = parseFechaExcel(f[1]);
                    if (!fecha) continue;
                    const pad = n => String(n).padStart(2, '0');
                    const fechaStr = `${fecha.getFullYear()}-${pad(fecha.getMonth()+1)}-${pad(fecha.getDate())}T${pad(fecha.getHours())}:${pad(fecha.getMinutes())}`;
                    const cuotasStr = String(f[7] || '1');
                    const cuotas = parseInt(cuotasStr) || 1;
                    const cuotaMensual = cuotas > 1 ? monto / cuotas : monto;
                    gastos.push({
                        nombre: concepto, cantidad: monto, fecha: fechaStr, categoria: String(f[3] || '').trim(),
                        origen: cuentaNombreToId(String(f[5] || '')) || 'efectivo', nota: String(f[6] || '') || null,
                        cuotas: cuotas, cuotaMensual
                    });
                }
            }

            let categorias = [];
            if (workbook.SheetNames.includes('Categorías')) {
                const ws = workbook.Sheets['Categorías'];
                const filas = XLSX.utils.sheet_to_json(ws, { header: 1 });
                for (let r = 1; r < filas.length; r++) {
                    const f = filas[r];
                    const nom = String(f[0] || '').trim();
                    if (!nom || nom === 'Sin categorías creadas') continue;
                    const limite = parseNumExcel(f[1]);
                    const color = String(f[2] || '').trim();
                    categorias.push({ nombre: nom, color: color || '#6b7280', limite: limite > 0 ? limite : null });
                }
            }

            const metasImportadas = [];
            let metas = [];
            if (workbook.SheetNames.includes('Metas')) {
                const ws = workbook.Sheets['Metas'];
                const filas = XLSX.utils.sheet_to_json(ws, { header: 1 });
                for (let r = 1; r < filas.length; r++) {
                    const f = filas[r];
                    const nom = String(f[0] || '').trim();
                    if (!nom || nom === 'Sin metas creadas') continue;
                    const id = 'meta_' + Date.now() + '_' + r;
                    const obj = parseNumExcel(f[1]);
                    metas.push({ id, nombre: nom, objetivo: obj, plazo: String(f[4] || '').trim() || null });
                    metasImportadas.push({ nombre: nom, id });
                }
            }

            let contribuciones = [];
            if (workbook.SheetNames.includes('Aportes a metas')) {
                const ws = workbook.Sheets['Aportes a metas'];
                const filas = XLSX.utils.sheet_to_json(ws, { header: 1 });
                for (let r = 1; r < filas.length; r++) {
                    const f = filas[r];
                    const metaNom = String(f[2] || '').trim();
                    if (!metaNom || metaNom.startsWith('═══') || metaNom === 'Sin aportes registrados') continue;
                    const monto = parseNumExcel(f[3]);
                    if (monto <= 0) continue;
                    const fecha = parseFechaExcel(f[1]);
                    if (!fecha) continue;
                    const metaObj = metasImportadas.find(m => m.nombre === metaNom);
                    if (!metaObj) continue;
                    const pad = n => String(n).padStart(2, '0');
                    const fechaStr = `${fecha.getFullYear()}-${pad(fecha.getMonth()+1)}-${pad(fecha.getDate())}T12:00:00`;
                    contribuciones.push({ metaId: metaObj.id, cantidad: monto, fecha: fechaStr, origen: cuentaNombreToId(String(f[4] || '')) || 'efectivo' });
                }
            }

            let pagosProgramados = [];
            if (workbook.SheetNames.includes('Pagos programados')) {
                const ws = workbook.Sheets['Pagos programados'];
                const filas = XLSX.utils.sheet_to_json(ws, { header: 1 });
                for (let r = 1; r < filas.length; r++) {
                    const f = filas[r];
                    const concepto = String(f[0] || '').trim();
                    if (!concepto || concepto === 'Sin pagos programados') continue;
                    const monto = parseNumExcel(f[1]);
                    if (monto <= 0) continue;
                    const fechaIni = f[6] ? (parseFechaExcel(f[6]) ? parseFechaExcel(f[6]).toISOString().slice(0, 10) : '') : '';
                    pagosProgramados.push({
                        id: 'pago_' + Date.now() + '_' + r,
                        concepto, monto, frecuencia: String(f[2] || 'mensual').toLowerCase(),
                        diaPago: parseInt(f[3]) || 1, cuenta: cuentaNombreToId(String(f[4] || '')) || 'efectivo',
                        categoria: String(f[5] || '').trim(), fechaInicio: fechaIni, activo: String(f[7] || 'Sí').toLowerCase() !== 'no', nota: ''
                    });
                }
            }

            if (moneda) localStorage.setItem('moneda', moneda);
            if (presupuestoMensual > 0) localStorage.setItem('presupuestoMensual', presupuestoMensual.toString());
            if (limiteTc > 0) localStorage.setItem('limiteTarjetaCredito', limiteTc.toString());
            localStorage.setItem('saldosCuentas', JSON.stringify(saldosIni));
            localStorage.setItem('ingresos', JSON.stringify(ingresos));
            localStorage.setItem('gastos', JSON.stringify(gastos));
            localStorage.setItem('categorias', JSON.stringify(categorias));
            localStorage.setItem('metas', JSON.stringify(metas));
            localStorage.setItem('contribucionesMetas', JSON.stringify(contribuciones));
            localStorage.setItem('pagosProgramados', JSON.stringify(pagosProgramados));

            alert('¡Importación completada!\n\n• ' + ingresos.length + ' ingresos\n• ' + gastos.length + ' gastos\n• ' + categorias.length + ' categorías\n• ' + metas.length + ' metas\n• ' + contribuciones.length + ' aportes\n• ' + pagosProgramados.length + ' pagos programados');
            location.reload();
        } catch (err) {
            console.error(err);
            alert('Error al importar: ' + (err.message || 'Formato de archivo no válido. Usa una plantilla exportada desde MoneyTrack.'));
        }
    };
    reader.readAsArrayBuffer(archivo);
}

/** Convierte nombre de cuenta a id (Efectivo->efectivo, Tarjeta de crédito->tarjetaCredito) */
function cuentaNombreToId(nombre) {
    if (!nombre || typeof nombre !== 'string') return '';
    const n = nombre.trim().toLowerCase();
    const map = { 'efectivo': 'efectivo', 'banco': 'banco', 'tarjeta de crédito': 'tarjetaCredito', 'nequi': 'nequi', 'daviplata': 'daviplata' };
    return map[n] || CUENTAS.find(c => c.nombre.toLowerCase() === nombre.trim())?.id || '';
}

/** Parsea número desde texto como "1.000,50 COP" o valor numérico */
function parseNumExcel(val) {
    if (val == null || val === '') return 0;
    if (typeof val === 'number' && !isNaN(val)) return val;
    const s = String(val).replace(/\s+[A-Z]{3}$/i, '').trim().replace(/\./g, '').replace(',', '.');
    const n = parseFloat(s);
    return isNaN(n) ? 0 : n;
}

/** Parsea fecha desde Excel (serial o string dd/mm/yyyy) */
function parseFechaExcel(val) {
    if (!val) return null;
    if (val instanceof Date) return val;
    if (typeof val === 'number') {
        const d = new Date((val - 25569) * 86400 * 1000);
        return isNaN(d.getTime()) ? null : d;
    }
    const s = String(val).trim();
    const m = s.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:\s+(\d{1,2}):(\d{2}))?/);
    if (m) {
        const d = new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1]), parseInt(m[4]) || 12, parseInt(m[5]) || 0);
        return isNaN(d.getTime()) ? null : d;
    }
    const d2 = new Date(s);
    return isNaN(d2.getTime()) ? null : d2;
}

/** Formatea fecha para localStorage (YYYY-MM-DDTHH:mm) */
function fechaToStr(d) {
    if (!d || !(d instanceof Date)) return '';
    const p = n => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}`;
}

/** Obtiene valor de celda desde hoja (por fila/col) */
function celdaVal(ws, r, c) {
    const col = String.fromCharCode(65 + c);
    const ref = col + (r + 1);
    const cell = ws[ref];
    return cell ? (cell.v !== undefined ? cell.v : cell.w) : '';
}

/** Importa datos desde archivo Excel (plantilla MoneyTrack) */
function importarDesdeExcel(archivo) {
    if (typeof XLSX === 'undefined') {
        alert('Error: La librería Excel no está cargada.');
        return;
    }
    if (!archivo || !archivo.name) return;
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const wb = XLSX.read(data, { type: 'array' });
            const MESES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];

            const getSheet = (name) => wb.SheetNames.includes(name) ? wb.Sheets[name] : null;
            const toArr = (ws) => ws ? XLSX.utils.sheet_to_json(ws, { header: 1 }) : [];

            let moneda = localStorage.getItem('moneda') || 'COP';
            const saldosIni = {};
            CUENTAS.forEach(c => { saldosIni[c.id] = 0; });

            const wsResumen = getSheet('Resumen');
            if (wsResumen) {
                const arr = toArr(wsResumen);
                for (let i = 0; i < arr.length; i++) {
                    const a = arr[i];
                    const k = String(a[0] || '').trim();
                    const v = a[1];
                    if (k === 'Moneda' && v) moneda = String(v).trim().toUpperCase();
                    if (k === 'Presupuesto mensual' && v && String(v) !== 'No definido') {
                        const n = parseNumExcel(v);
                        if (n > 0) localStorage.setItem('presupuestoMensual', n.toString());
                    }
                    if (k === 'Límite tarjeta de crédito' && v && String(v) !== 'No definido') {
                        const n = parseNumExcel(v);
                        if (n > 0) localStorage.setItem('limiteTarjetaCredito', n.toString());
                    }
                    if (CUENTAS.some(c => c.nombre === k)) {
                        const n = parseNumExcel(v);
                        const c = CUENTAS.find(x => x.nombre === k);
                        if (c) saldosIni[c.id] = n;
                    }
                }
            }
            localStorage.setItem('moneda', moneda);
            localStorage.setItem('saldosCuentas', JSON.stringify(saldosIni));

            const ingresos = [];
            const arrIng = toArr(getSheet('Ingresos'));
            for (let i = 1; i < arrIng.length; i++) {
                const r = arrIng[i];
                const fechaStr = r[1]; const concepto = r[2]; const monto = parseNumExcel(r[3]); const cuentaNom = r[4];
                if (!fechaStr || !concepto || monto <= 0) continue;
                if (String(concepto).includes('═══') || String(concepto).includes('Sin ingresos')) continue;
                const d = parseFechaExcel(fechaStr);
                if (!d) continue;
                const cuentaId = cuentaNombreToId(cuentaNom) || 'efectivo';
                ingresos.push({ cantidad: monto, fecha: fechaToStr(d), origen: cuentaId, nota: String(concepto || '').trim() || 'Ingreso' });
            }
            localStorage.setItem('ingresos', JSON.stringify(ingresos));

            const categorias = [];
            const arrCat = toArr(getSheet('Categorías'));
            for (let i = 1; i < arrCat.length; i++) {
                const r = arrCat[i];
                const nom = String(r[0] || '').trim();
                if (!nom || nom.includes('Sin categorías')) continue;
                const limite = parseNumExcel(r[1]);
                const color = String(r[2] || '').trim() || '#6b7280';
                categorias.push({ nombre: nom, color, limite: limite > 0 ? limite : null });
            }
            localStorage.setItem('categorias', JSON.stringify(categorias));

            const metas = [];
            const arrMetas = toArr(getSheet('Metas'));
            const metaNombreToId = {};
            for (let i = 1; i < arrMetas.length; i++) {
                const r = arrMetas[i];
                const nom = String(r[0] || '').trim();
                if (!nom || nom.includes('Sin metas')) continue;
                const id = 'meta_' + Date.now() + '_' + i;
                metaNombreToId[nom] = id;
                metas.push({ id, nombre: nom, objetivo: parseNumExcel(r[1]), plazo: r[4] && String(r[4]) !== '—' ? String(r[4]) : null });
            }
            localStorage.setItem('metas', JSON.stringify(metas));

            const contribuciones = [];
            const arrAport = toArr(getSheet('Aportes a metas'));
            for (let i = 1; i < arrAport.length; i++) {
                const r = arrAport[i];
                const fechaStr = r[1]; const metaNom = String(r[2] || '').trim(); const monto = parseNumExcel(r[3]); const cuentaNom = r[4];
                if (!fechaStr || !metaNom || monto <= 0) continue;
                if (metaNom.includes('═══') || metaNom.includes('Sin aportes')) continue;
                const metaId = metaNombreToId[metaNom] || metas[0]?.id;
                if (!metaId) continue;
                const d = parseFechaExcel(fechaStr);
                if (!d) continue;
                contribuciones.push({ metaId, cantidad: monto, fecha: fechaToStr(d), origen: cuentaNombreToId(cuentaNom) || 'efectivo' });
            }
            localStorage.setItem('contribucionesMetas', JSON.stringify(contribuciones));

            const gastos = [];
            const arrGastos = toArr(getSheet('Gastos'));
            for (let i = 1; i < arrGastos.length; i++) {
                const r = arrGastos[i];
                const fechaStr = r[1]; const concepto = r[2]; const cat = r[3]; const monto = parseNumExcel(r[4]);
                const cuentaNom = r[5]; const nota = r[6]; const cuotasStr = r[7];
                if (!fechaStr || !concepto || monto <= 0) continue;
                if (String(concepto).includes('═══') || String(concepto).includes('Sin gastos')) continue;
                const d = parseFechaExcel(fechaStr);
                if (!d) continue;
                let cuotas = 1;
                if (cuotasStr && String(cuotasStr).match(/\d+/)) cuotas = parseInt(String(cuotasStr).match(/\d+/)[0], 10) || 1;
                const cuotaMensual = cuotas > 1 ? monto / cuotas : monto;
                gastos.push({
                    nombre: String(concepto).trim(),
                    cantidad: monto,
                    fecha: fechaToStr(d),
                    categoria: String(cat || '').trim() || (categorias[0]?.nombre || 'Otros'),
                    origen: cuentaNombreToId(cuentaNom) || 'efectivo',
                    nota: nota ? String(nota).trim() : null,
                    cuotas: cuotas,
                    cuotaMensual
                });
            }
            localStorage.setItem('gastos', JSON.stringify(gastos));

            const pagosProgramados = [];
            const arrPagos = toArr(getSheet('Pagos programados'));
            for (let i = 1; i < arrPagos.length; i++) {
                const r = arrPagos[i];
                const concepto = String(r[0] || '').trim();
                if (!concepto || concepto.includes('Sin pagos')) continue;
                const monto = parseNumExcel(r[1]);
                if (monto <= 0) continue;
                const activo = String(r[7] || '').toLowerCase() !== 'no';
                pagosProgramados.push({
                    id: 'pago_' + Date.now() + '_' + i,
                    concepto,
                    monto,
                    frecuencia: String(r[2] || 'mensual').toLowerCase() || 'mensual',
                    diaPago: r[3] ? parseInt(r[3], 10) || 1 : 1,
                    cuenta: cuentaNombreToId(r[4]) || 'efectivo',
                    categoria: String(r[5] || '').trim() || (categorias[0]?.nombre || ''),
                    fechaInicio: r[6] ? String(r[6]).slice(0, 10) : new Date().toISOString().slice(0, 10),
                    activo,
                    nota: ''
                });
            }
            localStorage.setItem('pagosProgramados', JSON.stringify(pagosProgramados));

            alert('¡Importación completada!\n\nSe han cargado: ' + ingresos.length + ' ingresos, ' + gastos.length + ' gastos, ' + categorias.length + ' categorías, ' + metas.length + ' metas, ' + contribuciones.length + ' aportes y ' + pagosProgramados.length + ' pagos programados.');
            location.reload();
        } catch (err) {
            console.error(err);
            alert('Error al importar el archivo. Asegúrate de que sea una plantilla MoneyTrack válida (descargada desde esta aplicación).');
        }
    };
    reader.readAsArrayBuffer(archivo);
}

/** Convierte nombre de cuenta a id (Efectivo->efectivo, Tarjeta de crédito->tarjetaCredito) */
function cuentaNombreToId(nombre) {
    if (!nombre || typeof nombre !== 'string') return '';
    const n = String(nombre).trim().toLowerCase();
    const map = { 'efectivo': 'efectivo', 'banco': 'banco', 'tarjeta de crédito': 'tarjetaCredito', 'nequi': 'nequi', 'daviplata': 'daviplata' };
    return CUENTAS.find(c => c.nombre.toLowerCase() === nombre.trim())?.id || map[n] || '';
}

/** Parsea número desde Excel (puede ser número o "1.000,00 COP") */
function parseNumExcel(val) {
    if (val == null || val === '') return 0;
    if (typeof val === 'number' && !isNaN(val)) return val;
    const s = String(val).replace(/\s+[A-Z]{3}$/i, '').replace(/\./g, '').replace(',', '.');
    return parseFloat(s) || 0;
}

/** Parsea fecha desde Excel (serial o string dd/mm/yyyy) */
function parseFechaExcel(val) {
    if (!val) return null;
    if (val instanceof Date) return val;
    if (typeof val === 'number') {
        const d = new Date((val - 25569) * 86400 * 1000);
        return isNaN(d.getTime()) ? null : d;
    }
    const s = String(val).trim();
    const m = s.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:\s+(\d{1,2}):(\d{2}))?/);
    if (m) {
        const d = new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1]), parseInt(m[4]) || 12, parseInt(m[5]) || 0);
        return isNaN(d.getTime()) ? null : d;
    }
    const d2 = new Date(s);
    return isNaN(d2.getTime()) ? null : d2;
}

function fechaToISO(d) {
    if (!d) return null;
    const pad = n => String(n).padStart(2, '0');
    return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()) + 'T' + pad(d.getHours()) + ':' + pad(d.getMinutes()) + ':00';
}

/** Importa datos desde un archivo Excel (plantilla MoneyTrack) */
function importarDesdeExcel(archivo) {
    if (typeof XLSX === 'undefined') {
        alert('Error: La librería Excel no está cargada.');
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

            const getSheet = (nombre) => {
                const ws = workbook.Sheets[workbook.SheetNames.find(n => n === nombre)];
                return ws ? XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }) : [];
            };
            const celda = (row, col) => {
                if (!row || col >= row.length) return '';
                const v = row[col];
                return v != null ? (typeof v === 'object' && v.v !== undefined ? v.v : v) : '';
            };

            let moneda = localStorage.getItem('moneda') || '';
            const resumen = getSheet('Resumen');
            for (let r = 0; r < resumen.length; r++) {
                const a = String(celda(resumen[r], 0)).trim();
                const b = celda(resumen[r], 1);
                if (a === 'Moneda' && b) moneda = String(b).trim().split(/\s/)[0] || moneda;
                if (a === 'Presupuesto mensual' && b) {
                    const num = parseNumExcel(b);
                    if (num > 0) localStorage.setItem('presupuestoMensual', num.toString());
                }
                if (a === 'Límite tarjeta de crédito' && b) {
                    const num = parseNumExcel(b);
                    if (num > 0) localStorage.setItem('limiteTarjetaCredito', num.toString());
                }
                if (['Efectivo','Banco','Tarjeta de crédito','Nequi','Daviplata'].includes(a)) {
                    const num = parseNumExcel(b);
                    const id = cuentaNombreToId(a);
                    if (id) {
                        const saldos = JSON.parse(localStorage.getItem('saldosCuentas') || '{}');
                        saldos[id] = num;
                        localStorage.setItem('saldosCuentas', JSON.stringify(saldos));
                    }
                }
            }
            if (moneda) localStorage.setItem('moneda', moneda);

            const ingresos = [];
            const ingRows = getSheet('Ingresos');
            for (let r = 1; r < ingRows.length; r++) {
                const row = ingRows[r];
                const fechaStr = celda(row, 1);
                const concepto = celda(row, 2);
                const monto = parseNumExcel(celda(row, 3));
                const cuentaNom = celda(row, 4);
                if (!fechaStr || String(concepto).includes('═══') || String(concepto).includes('Sin ingresos')) continue;
                const d = parseFechaExcel(fechaStr);
                if (!d || monto <= 0) continue;
                const origen = cuentaNombreToId(cuentaNom) || 'efectivo';
                ingresos.push({ cantidad: monto, fecha: fechaToISO(d), origen, nota: concepto || 'Ingreso' });
            }
            localStorage.setItem('ingresos', JSON.stringify(ingresos));

            const gastos = [];
            const gasRows = getSheet('Gastos');
            for (let r = 1; r < gasRows.length; r++) {
                const row = gasRows[r];
                const fechaStr = celda(row, 1);
                const concepto = celda(row, 2);
                const categoria = celda(row, 3);
                const monto = parseNumExcel(celda(row, 4));
                const cuentaNom = celda(row, 5);
                const nota = celda(row, 6);
                const cuotasStr = celda(row, 7);
                if (!fechaStr || String(concepto).includes('═══') || String(concepto).includes('Sin gastos')) continue;
                const d = parseFechaExcel(fechaStr);
                if (!d || monto <= 0) continue;
                const origen = cuentaNombreToId(cuentaNom) || 'efectivo';
                const cuotas = parseInt(String(cuotasStr).replace(/\D/g, '')) || 1;
                const cuotaMensual = cuotas > 1 ? monto / cuotas : monto;
                gastos.push({ nombre: concepto, cantidad: monto, fecha: fechaToISO(d), categoria, origen, nota: nota || null, cuotas, cuotaMensual });
            }
            localStorage.setItem('gastos', JSON.stringify(gastos));

            const categorias = [];
            const catRows = getSheet('Categorías');
            for (let r = 1; r < catRows.length; r++) {
                const row = catRows[r];
                const nom = celda(row, 0);
                const limite = parseNumExcel(celda(row, 1));
                const color = celda(row, 2);
                if (!nom || String(nom).includes('Sin categorías')) continue;
                categorias.push({ nombre: nom, color: color || '#6b7280', limite: limite > 0 ? limite : null });
            }
            localStorage.setItem('categorias', JSON.stringify(categorias));

            const metas = [];
            const metaRows = getSheet('Metas');
            const metaIds = {};
            for (let r = 1; r < metaRows.length; r++) {
                const row = metaRows[r];
                const nom = celda(row, 0);
                const objetivo = parseNumExcel(celda(row, 1));
                const plazo = celda(row, 4);
                if (!nom || String(nom).includes('Sin metas')) continue;
                const id = 'meta_' + Date.now() + '_' + r;
                metaIds[nom] = id;
                metas.push({ id, nombre: nom, objetivo, plazo: plazo && plazo !== '—' ? plazo : null });
            }
            localStorage.setItem('metas', JSON.stringify(metas));

            const contribuciones = [];
            const aportRows = getSheet('Aportes a metas');
            for (let r = 1; r < aportRows.length; r++) {
                const row = aportRows[r];
                const fechaStr = celda(row, 1);
                const metaNom = celda(row, 2);
                const monto = parseNumExcel(celda(row, 3));
                const cuentaNom = celda(row, 4);
                if (!fechaStr || String(metaNom).includes('═══') || String(metaNom).includes('Sin aportes')) continue;
                const d = parseFechaExcel(fechaStr);
                const metaId = metaIds[metaNom];
                if (!d || monto <= 0 || !metaId) continue;
                contribuciones.push({ metaId, cantidad: monto, fecha: fechaToISO(d), origen: cuentaNombreToId(cuentaNom) || 'efectivo' });
            }
            localStorage.setItem('contribucionesMetas', JSON.stringify(contribuciones));

            const pagos = [];
            const pagRows = getSheet('Pagos programados');
            for (let r = 1; r < pagRows.length; r++) {
                const row = pagRows[r];
                const concepto = celda(row, 0);
                const monto = parseNumExcel(celda(row, 1));
                const frecuencia = celda(row, 2);
                const diaPago = celda(row, 3);
                const cuentaNom = celda(row, 4);
                const categoria = celda(row, 5);
                const fechaInicio = celda(row, 6);
                const activoStr = celda(row, 7);
                if (!concepto || String(concepto).includes('Sin pagos')) continue;
                pagos.push({
                    id: 'pago_' + Date.now() + '_' + r,
                    concepto,
                    monto,
                    frecuencia: frecuencia || 'mensual',
                    diaPago: parseInt(diaPago) || 1,
                    cuenta: cuentaNombreToId(cuentaNom) || 'efectivo',
                    categoria: categoria || '',
                    fechaInicio: fechaInicio ? String(fechaInicio).slice(0, 10) : new Date().toISOString().slice(0, 10),
                    activo: activoStr !== 'No',
                    nota: ''
                });
            }
            localStorage.setItem('pagosProgramados', JSON.stringify(pagos));

            alert('¡Importación completada! Se han cargado: ' + ingresos.length + ' ingresos, ' + gastos.length + ' gastos, ' + categorias.length + ' categorías, ' + metas.length + ' metas, ' + contribuciones.length + ' aportes, ' + pagos.length + ' pagos programados.');
            location.reload();
        } catch (err) {
            console.error(err);
            alert('Error al importar: ' + (err.message || 'Formato de archivo no válido. Asegúrate de usar una plantilla MoneyTrack exportada.'));
        }
    };
    reader.readAsArrayBuffer(archivo);
}

/**
 * Importa datos desde un archivo Excel (plantilla MoneyTrack).
 * Reconoce la estructura exportada y rellena el sistema automáticamente.
 */
function importarDesdeExcel(archivo) {
    if (typeof XLSX === 'undefined') {
        alert('Error: La librería Excel no está cargada. Recarga la página e intenta de nuevo.');
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

            const val = (ws, r, c) => {
                const cell = ws[XLSX.utils.encode_cell({ r, c })];
                return cell ? (cell.v !== undefined ? cell.v : '') : '';
            };
            const parseNum = (v) => {
                if (v === '' || v === null || v === undefined) return 0;
                if (typeof v === 'number' && !isNaN(v)) return v;
                const s = String(v).replace(/\s+[A-Z]{3}$/i, '').replace(/\./g, '').replace(',', '.');
                return parseFloat(s) || 0;
            };
            const parseFecha = (v) => {
                if (!v) return null;
                if (typeof v === 'number') {
                    const d = XLSX.SSF.parse_date_code ? XLSX.SSF.parse_date_code(v) : null;
                    if (d) return new Date(d.y, d.m - 1, d.d, d.H || 0, d.M || 0);
                    const epoch = (v - 25569) * 86400 * 1000;
                    return new Date(epoch);
                }
                const m = String(v).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2}))?/);
                if (m) return new Date(+m[3], +m[2]-1, +m[1], +(m[4]||0), +(m[5]||0));
                const mMes = MESES.findIndex(m => String(v).includes(m));
                if (mMes >= 0) {
                    const año = String(v).match(/\d{4}/);
                    return new Date(+(año&&año[0])||new Date().getFullYear(), mMes, 1);
                }
                return new Date(v) || null;
            };
            const fechaToStr = (d) => {
                if (!d || !(d instanceof Date) || isNaN(d)) return null;
                const p = n => String(n).padStart(2, '0');
                return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}:00`;
            };
            const cuentaNombreToId = (nombre) => {
                const n = (nombre||'').trim().toLowerCase();
                const map = { 'efectivo':'efectivo','banco':'banco','tarjeta de crédito':'tarjetaCredito','nequi':'nequi','daviplata':'daviplata' };
                return map[n] || CUENTAS.find(c => c.nombre.toLowerCase() === n)?.id || 'efectivo';
            };

            let moneda = localStorage.getItem('moneda') || '';
            const saldosCuentas = { efectivo:0, banco:0, tarjetaCredito:0, nequi:0, daviplata:0 };

            if (workbook.SheetNames.includes('Resumen')) {
                const ws = workbook.Sheets['Resumen'];
                const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
                for (let r = 0; r <= range.e.r; r++) {
                    const a = String(val(ws,r,0)).trim();
                    const b = val(ws,r,1);
                    if (a === 'Moneda' && b) moneda = String(b).trim();
                    if (a === 'Presupuesto mensual' && b) {
                        const n = parseNum(b);
                        if (n > 0) localStorage.setItem('presupuestoMensual', n.toString());
                    }
                    if (a === 'Límite tarjeta de crédito' && b) {
                        const n = parseNum(b);
                        if (n > 0) localStorage.setItem('limiteTarjetaCredito', n.toString());
                    }
                    CUENTAS.forEach(c => {
                        if (a === c.nombre && b) saldosCuentas[c.id] = parseNum(b);
                    });
                }
                localStorage.setItem('moneda', moneda);
                localStorage.setItem('saldosCuentas', JSON.stringify(saldosCuentas));
            }

            if (workbook.SheetNames.includes('Categorías')) {
                const ws = workbook.Sheets['Categorías'];
                const json = XLSX.utils.sheet_to_json(ws, { header: 1 });
                const cats = [];
                for (let i = 1; i < json.length; i++) {
                    const row = json[i];
                    const nom = String(row[0]||'').trim();
                    if (!nom || nom.startsWith('Sin ') || nom.startsWith('═══')) continue;
                    cats.push({ nombre: nom, limite: parseNum(row[1]) || null, color: String(row[2]||'').trim() || '#6b7280' });
                }
                localStorage.setItem('categorias', JSON.stringify(cats));
            }

            if (workbook.SheetNames.includes('Metas')) {
                const ws = workbook.Sheets['Metas'];
                const json = XLSX.utils.sheet_to_json(ws, { header: 1 });
                const metas = [];
                for (let i = 1; i < json.length; i++) {
                    const row = json[i];
                    const nom = String(row[0]||'').trim();
                    if (!nom || nom.startsWith('Sin ') || nom.startsWith('═══')) continue;
                    metas.push({ id: 'meta_' + Date.now() + '_' + i, nombre: nom, objetivo: parseNum(row[1]), plazo: String(row[4]||'').trim() || null });
                }
                localStorage.setItem('metas', JSON.stringify(metas));
            }

            const metasActuales = JSON.parse(localStorage.getItem('metas') || '[]');
            const metaNombreToId = (nombre) => metasActuales.find(m => (m.nombre||'').trim() === (nombre||'').trim())?.id || null;

            if (workbook.SheetNames.includes('Ingresos')) {
                const ws = workbook.Sheets['Ingresos'];
                const json = XLSX.utils.sheet_to_json(ws, { header: 1 });
                const ingresos = [];
                for (let i = 1; i < json.length; i++) {
                    const row = json[i];
                    const fechaVal = row[1] || row[0];
                    const monto = parseNum(row[3]);
                    const cuenta = cuentaNombreToId(row[4]);
                    if (!fechaVal || monto <= 0) continue;
                    const d = parseFecha(fechaVal);
                    if (!d) continue;
                    ingresos.push({ cantidad: monto, fecha: fechaToStr(d), origen: cuenta, nota: String(row[2]||'Ingreso').trim() });
                }
                localStorage.setItem('ingresos', JSON.stringify(ingresos));
            }

            if (workbook.SheetNames.includes('Gastos')) {
                const ws = workbook.Sheets['Gastos'];
                const json = XLSX.utils.sheet_to_json(ws, { header: 1 });
                const gastos = [];
                for (let i = 1; i < json.length; i++) {
                    const row = json[i];
                    const fechaVal = row[1] || row[0];
                    const monto = parseNum(row[4]);
                    const cuotasStr = String(row[7]||'1');
                    const cuotas = parseInt(cuotasStr) || 1;
                    if (!fechaVal || monto <= 0) continue;
                    const d = parseFecha(fechaVal);
                    if (!d) continue;
                    const cuotaMensual = cuotas > 1 ? monto / cuotas : monto;
                    gastos.push({
                        nombre: String(row[2]||'').trim(),
                        cantidad: monto,
                        fecha: fechaToStr(d),
                        categoria: String(row[3]||'').trim(),
                        origen: cuentaNombreToId(row[5]),
                        nota: String(row[6]||'').trim() || null,
                        cuotas: cuotas,
                        cuotaMensual: cuotaMensual
                    });
                }
                localStorage.setItem('gastos', JSON.stringify(gastos));
            }

            if (workbook.SheetNames.includes('Aportes a metas')) {
                const ws = workbook.Sheets['Aportes a metas'];
                const json = XLSX.utils.sheet_to_json(ws, { header: 1 });
                const contribuciones = [];
                for (let i = 1; i < json.length; i++) {
                    const row = json[i];
                    const fechaVal = row[1];
                    const metaNombre = String(row[2]||'').trim();
                    const monto = parseNum(row[3]);
                    const metaId = metaNombreToId(metaNombre);
                    if (!fechaVal || monto <= 0 || !metaId) continue;
                    const d = parseFecha(fechaVal);
                    if (!d) continue;
                    contribuciones.push({ metaId, cantidad: monto, fecha: fechaToStr(d), origen: cuentaNombreToId(row[4]) });
                }
                localStorage.setItem('contribucionesMetas', JSON.stringify(contribuciones));
            }

            if (workbook.SheetNames.includes('Pagos programados')) {
                const ws = workbook.Sheets['Pagos programados'];
                const json = XLSX.utils.sheet_to_json(ws, { header: 1 });
                const pagos = [];
                for (let i = 1; i < json.length; i++) {
                    const row = json[i];
                    const concepto = String(row[0]||'').trim();
                    const monto = parseNum(row[1]);
                    if (!concepto || concepto.startsWith('Sin ') || monto <= 0) continue;
                    const activo = String(row[7]||'Sí').toLowerCase() !== 'no';
                    const fechaIni = row[6] ? String(row[6]).slice(0,10) : new Date().toISOString().slice(0,10);
                    pagos.push({
                        id: 'pago_' + Date.now() + '_' + i,
                        concepto, monto,
                        frecuencia: String(row[2]||'mensual').trim() || 'mensual',
                        diaPago: parseInt(row[3]) || 1,
                        cuenta: cuentaNombreToId(row[4]),
                        categoria: String(row[5]||'').trim(),
                        fechaInicio: fechaIni,
                        activo, nota: ''
                    });
                }
                localStorage.setItem('pagosProgramados', JSON.stringify(pagos));
            }

            alert('¡Importación completada! Los datos del Excel se han cargado correctamente.');
            location.reload();
        } catch (err) {
            console.error(err);
            alert('Error al importar: ' + (err.message || 'Formato de archivo no reconocido. Asegúrate de usar una plantilla MoneyTrack exportada.'));
        }
    };
    reader.readAsArrayBuffer(archivo);
}

/** Convierte nombre de cuenta a id */
function cuentaNombreToId(nombre) {
    if (!nombre || typeof nombre !== 'string') return '';
    const n = nombre.trim().toLowerCase();
    const map = { 'efectivo': 'efectivo', 'banco': 'banco', 'tarjeta de crédito': 'tarjetaCredito', 'nequi': 'nequi', 'daviplata': 'daviplata' };
    for (const [k, v] of Object.entries(map)) if (n.includes(k)) return v;
    return CUENTAS.find(c => c.nombre.toLowerCase() === n)?.id || '';
}

/** Parsea número desde texto con formato español (1.000,50) o número */
function parseNumExcel(val) {
    if (val == null || val === '') return 0;
    if (typeof val === 'number' && !isNaN(val)) return val;
    const s = String(val).trim().replace(/\s+[A-Z]{3}$/i, '').replace(/\./g, '').replace(',', '.');
    return parseFloat(s) || 0;
}

/** Parsea fecha desde Excel (serial o string dd/mm/yyyy) */
function parseFechaExcel(val) {
    if (!val) return null;
    if (val instanceof Date) return val;
    if (typeof val === 'number') {
        const d = new Date((val - 25569) * 86400 * 1000);
        return isNaN(d.getTime()) ? null : d;
    }
    const s = String(val).trim();
    const m = s.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:\s+(\d{1,2}):(\d{2}))?/);
    if (m) {
        const d = new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1]), m[4] || 12, m[5] || 0);
        return isNaN(d.getTime()) ? null : d;
    }
    return null;
}

/** Formatea fecha a ISO para localStorage */
function fechaToISO(d) {
    if (!d) return null;
    const pad = n => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}:00`;
}

/** Importa datos desde un archivo Excel (plantilla MoneyTrack) */
function importarDesdeExcel(archivo) {
    if (typeof XLSX === 'undefined') {
        alert('Error: La librería Excel no está cargada.');
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

            const getSheet = (nombre) => workbook.Sheets[nombre] || workbook.Sheets[workbook.SheetNames.find(n => n.toLowerCase().includes(nombre.toLowerCase()))];
            const sheetToRows = (ws) => ws ? XLSX.utils.sheet_to_json(ws, { header: 1 }) : [];
            const celda = (row, col) => { const r = row[col]; return r != null ? (typeof r === 'object' && r.v !== undefined ? r.v : r) : ''; };

            let moneda = '';
            const saldosIni = {};
            let presupuesto = 0, limiteTc = 0;

            const wsResumen = getSheet('Resumen');
            if (wsResumen) {
                const rows = sheetToRows(wsResumen);
                for (let i = 0; i < rows.length; i++) {
                    const a = String(rows[i][0] || '').trim();
                    const b = rows[i][1];
                    if (a === 'Moneda') moneda = String(b || '').trim();
                    else if (a === 'Presupuesto mensual' && b) presupuesto = parseNumExcel(b);
                    else if (a === 'Límite tarjeta de crédito' && b) limiteTc = parseNumExcel(b);
                    else if (CUENTAS.some(c => c.nombre === a)) saldosIni[CUENTAS.find(c => c.nombre === a).id] = parseNumExcel(b);
                }
            }

            if (moneda) localStorage.setItem('moneda', moneda);
            if (presupuesto > 0) localStorage.setItem('presupuestoMensual', presupuesto.toString());
            if (limiteTc > 0) localStorage.setItem('limiteTarjetaCredito', limiteTc.toString());
            localStorage.setItem('saldosCuentas', JSON.stringify(CUENTAS.reduce((acc, c) => { acc[c.id] = saldosIni[c.id] || 0; return acc; }, {})));

            const ingresos = [];
            const wsIng = getSheet('Ingresos');
            if (wsIng) {
                const rows = sheetToRows(wsIng);
                for (let i = 1; i < rows.length; i++) {
                    const r = rows[i];
                    const mes = celda(r, 0);
                    const fechaStr = celda(r, 1);
                    const concepto = celda(r, 2);
                    const monto = parseNumExcel(celda(r, 3));
                    const cuentaNom = celda(r, 4);
                    if (!mes || String(mes).includes('═══') || !fechaStr || monto <= 0) continue;
                    const cuentaId = cuentaNombreToId(cuentaNom) || 'efectivo';
                    const d = parseFechaExcel(fechaStr);
                    if (!d) continue;
                    ingresos.push({ cantidad: monto, fecha: fechaToISO(d), origen: cuentaId, nota: concepto || 'Ingreso' });
                }
            }
            localStorage.setItem('ingresos', JSON.stringify(ingresos));

            const categorias = [];
            const wsCat = getSheet('Categorías');
            if (wsCat) {
                const rows = sheetToRows(wsCat);
                for (let i = 1; i < rows.length; i++) {
                    const r = rows[i];
                    const nom = celda(r, 0);
                    if (!nom || String(nom).includes('Sin categorías')) continue;
                    const limite = parseNumExcel(celda(r, 1));
                    const color = celda(r, 2) || '#6b7280';
                    categorias.push({ nombre: String(nom).trim(), limite: limite > 0 ? limite : null, color });
                }
            }
            localStorage.setItem('categorias', JSON.stringify(categorias));

            const metas = [];
            const metaNombreToId = {};
            const wsMet = getSheet('Metas');
            if (wsMet) {
                const rows = sheetToRows(wsMet);
                for (let i = 1; i < rows.length; i++) {
                    const r = rows[i];
                    const nom = celda(r, 0);
                    if (!nom || String(nom).includes('Sin metas')) continue;
                    const id = 'meta_' + Date.now() + '_' + i;
                    metas.push({ id, nombre: String(nom).trim(), objetivo: parseNumExcel(celda(r, 1)), plazo: celda(r, 4) || null });
                    metaNombreToId[String(nom).trim().toLowerCase()] = id;
                }
            }
            localStorage.setItem('metas', JSON.stringify(metas));

            const contribuciones = [];
            const wsAport = getSheet('Aportes a metas');
            if (wsAport) {
                const rows = sheetToRows(wsAport);
                for (let i = 1; i < rows.length; i++) {
                    const r = rows[i];
                    const fechaStr = celda(r, 1);
                    const metaNom = celda(r, 2);
                    const monto = parseNumExcel(celda(r, 3));
                    const cuentaNom = celda(r, 4);
                    if (!fechaStr || String(metaNom).includes('═══') || monto <= 0) continue;
                    const metaId = metaNombreToId[String(metaNom).trim().toLowerCase()] || metas[0]?.id;
                    if (!metaId) continue;
                    const d = parseFechaExcel(fechaStr);
                    if (!d) continue;
                    contribuciones.push({ metaId, cantidad: monto, fecha: fechaToISO(d), origen: cuentaNombreToId(cuentaNom) || 'efectivo' });
                }
            }
            localStorage.setItem('contribucionesMetas', JSON.stringify(contribuciones));

            const gastos = [];
            const wsGas = getSheet('Gastos');
            if (wsGas) {
                const rows = sheetToRows(wsGas);
                for (let i = 1; i < rows.length; i++) {
                    const r = rows[i];
                    const fechaStr = celda(r, 1);
                    const concepto = celda(r, 2);
                    const categoria = celda(r, 3);
                    const monto = parseNumExcel(celda(r, 4));
                    const cuentaNom = celda(r, 5);
                    const nota = celda(r, 6);
                    const cuotasStr = celda(r, 7);
                    if (!fechaStr || String(concepto).includes('═══') || monto <= 0) continue;
                    const d = parseFechaExcel(fechaStr);
                    if (!d) continue;
                    const cuotas = parseInt(String(cuotasStr).replace(/\D/g, '')) || 1;
                    const cuotaMensual = cuotas > 1 ? monto / cuotas : monto;
                    gastos.push({
                        nombre: String(concepto || '').trim(),
                        cantidad: monto,
                        fecha: fechaToISO(d),
                        categoria: String(categoria || '').trim() || categorias[0]?.nombre,
                        origen: cuentaNombreToId(cuentaNom) || 'efectivo',
                        nota: nota || null,
                        cuotas,
                        cuotaMensual
                    });
                }
            }
            localStorage.setItem('gastos', JSON.stringify(gastos));

            const pagos = [];
            const wsPag = getSheet('Pagos programados');
            if (wsPag) {
                const rows = sheetToRows(wsPag);
                for (let i = 1; i < rows.length; i++) {
                    const r = rows[i];
                    const concepto = celda(r, 0);
                    if (!concepto || String(concepto).includes('Sin pagos')) continue;
                    const monto = parseNumExcel(celda(r, 1));
                    const activoStr = String(celda(r, 7)).toLowerCase();
                    pagos.push({
                        id: 'pago_' + Date.now() + '_' + i,
                        concepto: String(concepto).trim(),
                        monto,
                        frecuencia: celda(r, 2) || 'mensual',
                        diaPago: parseInt(celda(r, 3)) || 1,
                        cuenta: cuentaNombreToId(celda(r, 4)) || 'efectivo',
                        categoria: celda(r, 5) || '',
                        fechaInicio: celda(r, 6) || new Date().toISOString().slice(0, 10),
                        activo: activoStr !== 'no' && activoStr !== 'false',
                        nota: ''
                    });
                }
            }
            localStorage.setItem('pagosProgramados', JSON.stringify(pagos));

            alert('¡Importación completada! Se han cargado: ' + ingresos.length + ' ingresos, ' + gastos.length + ' gastos, ' + categorias.length + ' categorías, ' + metas.length + ' metas.');
            location.reload();
        } catch (err) {
            console.error(err);
            alert('Error al importar: ' + (err.message || 'Formato de archivo no válido. Usa una plantilla exportada desde MoneyTrack.'));
        }
    };
    reader.readAsArrayBuffer(archivo);
}

/** Obtiene el valor de una celda del sheet (por fila/col 0-based) */
function obtenerCelda(ws, r, c) {
    const col = String.fromCharCode(65 + c);
    const ref = col + (r + 1);
    const cell = ws[ref];
    return cell ? (cell.v !== undefined ? cell.v : cell.w) : '';
}

/** Parsea un número desde texto como "1.000,00 COP" o número */
function parsearNumeroExcel(val) {
    if (val === '' || val === null || val === undefined) return 0;
    if (typeof val === 'number' && !isNaN(val)) return val;
    const s = String(val).trim().replace(/\s+[A-Z]{3}$/i, '').replace(/\./g, '').replace(',', '.');
    const n = parseFloat(s);
    return isNaN(n) ? 0 : n;
}

/** Parsea fecha: Excel serial, o string dd/mm/yyyy o dd/mm/yyyy hh:mm */
function parsearFechaExcel(val) {
    if (!val) return null;
    if (typeof val === 'number') {
        const d = new Date((val - 25569) * 86400 * 1000);
        return isNaN(d.getTime()) ? null : d;
    }
    const s = String(val).trim();
    const match = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2}))?/);
    if (match) {
        const [, d, m, y, h, min] = match;
        const fecha = new Date(y, parseInt(m, 10) - 1, parseInt(d, 10), parseInt(h || 0, 10), parseInt(min || 0, 10));
        return isNaN(fecha.getTime()) ? null : fecha;
    }
    const d = new Date(s);
    return isNaN(d.getTime()) ? null : d;
}

/** Convierte Date a string ISO para fecha (YYYY-MM-DD) o datetime (YYYY-MM-DDTHH:mm) */
function fechaAString(d, conHora) {
    if (!d || !(d instanceof Date)) return '';
    const p = n => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())}` + (conHora ? `T${p(d.getHours())}:${p(d.getMinutes())}` : 'T12:00:00');
}

/** Mapea nombre de cuenta a id */
function cuentaNombreToId(nombre) {
    if (!nombre) return 'efectivo';
    const n = String(nombre).trim().toLowerCase();
    const map = { 'efectivo': 'efectivo', 'banco': 'banco', 'tarjeta de crédito': 'tarjetaCredito', 'tarjeta de credito': 'tarjetaCredito', 'nequi': 'nequi', 'daviplata': 'daviplata' };
    return map[n] || CUENTAS.find(c => c.nombre.toLowerCase() === nombre.trim())?.id || 'efectivo';
}

/** Importa datos desde un archivo Excel (plantilla MoneyTrack) */
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

            let moneda = localStorage.getItem('moneda') || '';
            const saldosIni = {};
            CUENTAS.forEach(c => { saldosIni[c.id] = 0; });

            const wsResumen = workbook.Sheets['Resumen'] || workbook.Sheets[workbook.SheetNames[0]];
            if (wsResumen) {
                const range = XLSX.utils.decode_range(wsResumen['!ref'] || 'A1');
                for (let r = 0; r <= range.e.r; r++) {
                    const a = obtenerCelda(wsResumen, r, 0);
                    const b = obtenerCelda(wsResumen, r, 1);
                    if (String(a).trim() === 'Moneda' && b) moneda = String(b).trim().replace(/\s+\d+$/, '').trim();
                    if (String(a).trim() === 'Presupuesto mensual' && b && String(b) !== 'No definido') {
                        const num = parsearNumeroExcel(b);
                        if (num > 0) localStorage.setItem('presupuestoMensual', num.toString());
                    }
                    if (String(a).trim() === 'Límite tarjeta de crédito' && b && String(b) !== 'No definido') {
                        const num = parsearNumeroExcel(b);
                        if (num > 0) localStorage.setItem('limiteTarjetaCredito', num.toString());
                    }
                    const nomCuenta = String(a).trim();
                    const cuenta = CUENTAS.find(c => c.nombre === nomCuenta);
                    if (cuenta && b && !String(a).includes('SALDO') && !String(a).includes('CONFIGURACIÓN')) {
                        const num = parsearNumeroExcel(b);
                        saldosIni[cuenta.id] = num;
                    }
                }
                if (moneda) localStorage.setItem('moneda', moneda);
                localStorage.setItem('saldosCuentas', JSON.stringify(saldosIni));
            }

            const ingresos = [];
            const wsIng = workbook.Sheets['Ingresos'];
            if (wsIng) {
                const datos = XLSX.utils.sheet_to_json(wsIng, { header: 1 });
                for (let i = 1; i < datos.length; i++) {
                    const row = datos[i];
                    if (!row || row.length < 5) continue;
                    const fechaStr = row[1], concepto = row[2], monto = row[3], cuentaNom = row[4];
                    if (!fechaStr || String(fechaStr).includes('═══') || String(concepto).includes('Sin ingresos')) continue;
                    const fecha = parsearFechaExcel(fechaStr);
                    const num = parsearNumeroExcel(monto);
                    if (fecha && num > 0) {
                        ingresos.push({ cantidad: num, fecha: fechaAString(fecha, true), origen: cuentaNombreToId(cuentaNom), nota: concepto || 'Ingreso' });
                    }
                }
            }
            localStorage.setItem('ingresos', JSON.stringify(ingresos));

            const gastos = [];
            const wsGas = workbook.Sheets['Gastos'];
            if (wsGas) {
                const datos = XLSX.utils.sheet_to_json(wsGas, { header: 1 });
                for (let i = 1; i < datos.length; i++) {
                    const row = datos[i];
                    if (!row || row.length < 6) continue;
                    const fechaStr = row[1], nombre = row[2], categoria = row[3], monto = row[4], cuentaNom = row[5], nota = row[6] || '', cuotasStr = row[7];
                    if (!fechaStr || String(fechaStr).includes('═══') || String(nombre).includes('Sin gastos')) continue;
                    const fecha = parsearFechaExcel(fechaStr);
                    const num = parsearNumeroExcel(monto);
                    if (fecha && num > 0) {
                        const cuotas = (cuotasStr && String(cuotasStr).match(/(\d+)/)) ? parseInt(RegExp.$1, 10) : 1;
                        const origen = cuentaNombreToId(cuentaNom);
                        const cuotaMensual = (origen === 'tarjetaCredito' && cuotas > 1) ? num / cuotas : num;
                        gastos.push({ nombre, cantidad: num, fecha: fechaAString(fecha, true), categoria, origen, nota: nota || null, cuotas, cuotaMensual });
                    }
                }
            }
            localStorage.setItem('gastos', JSON.stringify(gastos));

            const categorias = [];
            const wsCat = workbook.Sheets['Categorías'];
            if (wsCat) {
                const datos = XLSX.utils.sheet_to_json(wsCat, { header: 1 });
                for (let i = 1; i < datos.length; i++) {
                    const row = datos[i];
                    if (!row || !row[0] || String(row[0]).includes('Sin categorías')) continue;
                    const limite = parsearNumeroExcel(row[1]);
                    categorias.push({ nombre: String(row[0]).trim(), color: row[2] || '#6b7280', limite: limite > 0 ? limite : null });
                }
            }
            localStorage.setItem('categorias', JSON.stringify(categorias));

            const metas = [];
            const metaNombreToId = {};
            const wsMet = workbook.Sheets['Metas'];
            if (wsMet) {
                const datos = XLSX.utils.sheet_to_json(wsMet, { header: 1 });
                for (let i = 1; i < datos.length; i++) {
                    const row = datos[i];
                    if (!row || !row[0] || String(row[0]).includes('Sin metas')) continue;
                    const id = 'meta_' + Date.now() + '_' + i;
                    metaNombreToId[String(row[0]).trim()] = id;
                    metas.push({ id, nombre: String(row[0]).trim(), objetivo: parsearNumeroExcel(row[1]), plazo: row[4] || null });
                }
            }
            localStorage.setItem('metas', JSON.stringify(metas));

            const contribuciones = [];
            const wsAport = workbook.Sheets['Aportes a metas'];
            if (wsAport) {
                const datos = XLSX.utils.sheet_to_json(wsAport, { header: 1 });
                for (let i = 1; i < datos.length; i++) {
                    const row = datos[i];
                    if (!row || row.length < 5) continue;
                    const fechaStr = row[1], metaNom = row[2], monto = row[3], cuentaNom = row[4];
                    if (!fechaStr || String(fechaStr).includes('═══') || String(metaNom).includes('Sin aportes')) continue;
                    const metaId = metaNombreToId[String(metaNom).trim()];
                    if (!metaId) continue;
                    const fecha = parsearFechaExcel(fechaStr);
                    const num = parsearNumeroExcel(monto);
                    if (fecha && num > 0) {
                        contribuciones.push({ metaId, cantidad: num, fecha: fechaAString(fecha, false), origen: cuentaNombreToId(cuentaNom) });
                    }
                }
            }
            localStorage.setItem('contribucionesMetas', JSON.stringify(contribuciones));

            const pagos = [];
            const wsPagos = workbook.Sheets['Pagos programados'];
            if (wsPagos) {
                const datos = XLSX.utils.sheet_to_json(wsPagos, { header: 1 });
                for (let i = 1; i < datos.length; i++) {
                    const row = datos[i];
                    if (!row || !row[0] || String(row[0]).includes('Sin pagos')) continue;
                    const activo = String(row[7] || 'Sí').toLowerCase();
                    pagos.push({
                        id: 'pago_' + Date.now() + '_' + i,
                        concepto: String(row[0]).trim(),
                        monto: parsearNumeroExcel(row[1]),
                        frecuencia: (row[2] || 'mensual').toLowerCase(),
                        diaPago: row[3] || 1,
                        cuenta: cuentaNombreToId(row[4]),
                        categoria: row[5] || '',
                        fechaInicio: row[6] ? String(row[6]).slice(0, 10) : new Date().toISOString().slice(0, 10),
                        activo: activo !== 'no' && activo !== '0',
                        nota: ''
                    });
                }
            }
            localStorage.setItem('pagosProgramados', JSON.stringify(pagos));

            alert('¡Importación completada! Se han cargado: ' + ingresos.length + ' ingresos, ' + gastos.length + ' gastos, ' + categorias.length + ' categorías, ' + metas.length + ' metas, ' + contribuciones.length + ' aportes, ' + pagos.length + ' pagos programados.');
            location.reload();
        } catch (err) {
            console.error(err);
            alert('Error al importar: ' + (err.message || 'Formato de archivo no válido. Usa una plantilla exportada desde MoneyTrack.'));
        }
    };
    reader.readAsArrayBuffer(archivo);
}

/** Convierte nombre de cuenta a id (Efectivo -> efectivo, Tarjeta de crédito -> tarjetaCredito) */
function cuentaNombreToId(nombre) {
    if (!nombre || typeof nombre !== 'string') return '';
    const n = nombre.trim();
    const map = { 'Efectivo': 'efectivo', 'Banco': 'banco', 'Tarjeta de crédito': 'tarjetaCredito', 'Nequi': 'nequi', 'Daviplata': 'daviplata' };
    return map[n] || CUENTAS.find(c => c.nombre.toLowerCase() === n.toLowerCase())?.id || '';
}

/** Parsea número desde texto "1.000,00 COP" o valor numérico */
function parseNumExcel(val) {
    if (val == null || val === '') return 0;
    if (typeof val === 'number' && !isNaN(val)) return val;
    const s = String(val).replace(/\s+[A-Z]{3}$/i, '').replace(/\./g, '').replace(',', '.');
    const n = parseFloat(s);
    return isNaN(n) ? 0 : n;
}

/** Parsea fecha desde Excel (serial o texto dd/mm/yyyy) */
function parseFechaExcel(val) {
    if (!val) return null;
    if (typeof val === 'number') {
        const d = new Date((val - 25569) * 86400 * 1000);
        return isNaN(d.getTime()) ? null : d;
    }
    const s = String(val).trim();
    const m = s.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:\s+(\d{1,2}):(\d{2}))?/);
    if (m) {
        const d = new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1]), parseInt(m[4]) || 12, parseInt(m[5]) || 0);
        return isNaN(d.getTime()) ? null : d;
    }
    const d = new Date(s);
    return isNaN(d.getTime()) ? null : d;
}

/** Importa datos desde archivo Excel (plantilla MoneyTrack) */
function importarDesdeExcel(archivo) {
    if (typeof XLSX === 'undefined') {
        alert('Error: La librería Excel no está cargada.');
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

            const getVal = (ws, r, c) => {
                const cell = ws[XLSX.utils.encode_cell({ r, c })];
                return cell ? (cell.v != null ? cell.v : '') : '';
            };
            const toArr = (ws) => XLSX.utils.sheet_to_json(ws, { header: 1 });

            let moneda = '';
            let presupuestoMensual = 0;
            let limiteTc = 0;
            const saldosIni = { efectivo: 0, banco: 0, tarjetaCredito: 0, nequi: 0, daviplata: 0 };

            if (workbook.SheetNames.includes('Resumen')) {
                const arr = toArr(workbook.Sheets['Resumen']);
                for (let i = 0; i < arr.length; i++) {
                    const a = arr[i];
                    const k = String(a[0] || '').trim();
                    const v = a[1];
                    if (k === 'Moneda' && v) moneda = String(v).trim();
                    if (k === 'Presupuesto mensual' && v) presupuestoMensual = parseNumExcel(v);
                    if (k === 'Límite tarjeta de crédito' && v) limiteTc = parseNumExcel(v);
                    if (CUENTAS.some(c => c.nombre === k)) saldosIni[cuentaNombreToId(k)] = parseNumExcel(v);
                }
            }
            if (moneda) localStorage.setItem('moneda', moneda);
            if (presupuestoMensual > 0) localStorage.setItem('presupuestoMensual', presupuestoMensual.toString());
            if (limiteTc > 0) localStorage.setItem('limiteTarjetaCredito', limiteTc.toString());
            localStorage.setItem('saldosCuentas', JSON.stringify(saldosIni));

            const ingresos = [];
            if (workbook.SheetNames.includes('Ingresos')) {
                const arr = toArr(workbook.Sheets['Ingresos']);
                for (let i = 1; i < arr.length; i++) {
                    const r = arr[i];
                    const fechaStr = r[1], concepto = r[2], monto = parseNumExcel(r[3]), cuenta = r[4];
                    if (!fechaStr && !concepto && !monto) continue;
                    if (String(concepto || '').includes('═══') || String(concepto || '').includes('Sin ingresos')) continue;
                    const d = parseFechaExcel(fechaStr);
                    if (!d || !cuentaNombreToId(cuenta)) continue;
                    const pad = n => String(n).padStart(2, '0');
                    ingresos.push({ cantidad: monto, fecha: `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`, origen: cuentaNombreToId(cuenta), nota: concepto || 'Ingreso' });
                }
            }
            localStorage.setItem('ingresos', JSON.stringify(ingresos));

            const gastos = [];
            if (workbook.SheetNames.includes('Gastos')) {
                const arr = toArr(workbook.Sheets['Gastos']);
                for (let i = 1; i < arr.length; i++) {
                    const r = arr[i];
                    const fechaStr = r[1], nombre = r[2], categoria = r[3], monto = parseNumExcel(r[4]), cuenta = r[5], nota = r[6], cuotasStr = r[7];
                    if (!fechaStr && !nombre && !monto) continue;
                    if (String(nombre || '').includes('═══') || String(nombre || '').includes('Sin gastos')) continue;
                    const d = parseFechaExcel(fechaStr);
                    if (!d || !cuentaNombreToId(cuenta)) continue;
                    const cuotas = parseInt(String(cuotasStr || '1').replace(/\D/g, '')) || 1;
                    const pad = n => String(n).padStart(2, '0');
                    gastos.push({ nombre, cantidad: monto, fecha: `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`, categoria, origen: cuentaNombreToId(cuenta), nota: nota || null, cuotas, cuotaMensual: monto / cuotas });
                }
            }
            localStorage.setItem('gastos', JSON.stringify(gastos));

            const categorias = [];
            if (workbook.SheetNames.includes('Categorías')) {
                const arr = toArr(workbook.Sheets['Categorías']);
                for (let i = 1; i < arr.length; i++) {
                    const r = arr[i];
                    const nom = String(r[0] || '').trim();
                    if (!nom || nom.includes('Sin categorías')) continue;
                    const limite = parseNumExcel(r[1]);
                    const color = String(r[2] || '').trim();
                    categorias.push({ nombre: nom, color: color || '#6b7280', limite: limite > 0 ? limite : null });
                }
            }
            localStorage.setItem('categorias', JSON.stringify(categorias));

            const metas = [];
            const metaNombreToId = {};
            if (workbook.SheetNames.includes('Metas')) {
                const arr = toArr(workbook.Sheets['Metas']);
                for (let i = 1; i < arr.length; i++) {
                    const r = arr[i];
                    const nom = String(r[0] || '').trim();
                    if (!nom || nom.includes('Sin metas')) continue;
                    const id = 'meta_' + Date.now() + '_' + i;
                    metaNombreToId[nom] = id;
                    metas.push({ id, nombre: nom, objetivo: parseNumExcel(r[1]), plazo: r[4] && r[4] !== '—' ? String(r[4]) : null });
                }
            }
            localStorage.setItem('metas', JSON.stringify(metas));

            const contribuciones = [];
            if (workbook.SheetNames.includes('Aportes a metas')) {
                const arr = toArr(workbook.Sheets['Aportes a metas']);
                for (let i = 1; i < arr.length; i++) {
                    const r = arr[i];
                    const fechaStr = r[1], metaNom = String(r[2] || '').trim(), monto = parseNumExcel(r[3]), cuenta = r[4];
                    if (!fechaStr || !metaNom || !monto) continue;
                    if (metaNom.includes('═══') || metaNom.includes('Sin aportes')) continue;
                    const metaId = metaNombreToId[metaNom] || Object.keys(metaNombreToId).find(k => k.toLowerCase() === metaNom.toLowerCase()) ? metaNombreToId[Object.keys(metaNombreToId).find(k => k.toLowerCase() === metaNom.toLowerCase())] : null;
                    if (!metaId) continue;
                    const d = parseFechaExcel(fechaStr);
                    if (!d) continue;
                    const pad = n => String(n).padStart(2, '0');
                    contribuciones.push({ metaId, cantidad: monto, fecha: `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}T12:00:00`, origen: cuentaNombreToId(cuenta) });
                }
            }
            localStorage.setItem('contribucionesMetas', JSON.stringify(contribuciones));

            const pagosProgramados = [];
            if (workbook.SheetNames.includes('Pagos programados')) {
                const arr = toArr(workbook.Sheets['Pagos programados']);
                for (let i = 1; i < arr.length; i++) {
                    const r = arr[i];
                    const concepto = String(r[0] || '').trim();
                    if (!concepto || concepto.includes('Sin pagos')) continue;
                    const monto = parseNumExcel(r[1]);
                    const frecuencia = String(r[2] || '').trim();
                    const diaPago = r[3] != null ? String(r[3]) : '';
                    const cuenta = cuentaNombreToId(r[4]);
                    const categoria = String(r[5] || '').trim();
                    const fechaInicio = r[6] ? String(r[6]).slice(0, 10) : '';
                    const activo = String(r[7] || 'Sí').toLowerCase() !== 'no';
                    pagosProgramados.push({ id: 'pago_' + Date.now() + '_' + i, concepto, monto, frecuencia, diaPago, cuenta, categoria, fechaInicio, activo });
                }
            }
            localStorage.setItem('pagosProgramados', JSON.stringify(pagosProgramados));

            alert('¡Importación completada! Se han cargado: ' + ingresos.length + ' ingresos, ' + gastos.length + ' gastos, ' + categorias.length + ' categorías, ' + metas.length + ' metas.');
            location.reload();
        } catch (err) {
            console.error(err);
            alert('Error al importar: ' + (err.message || 'Formato de archivo no válido. Asegúrate de usar una plantilla MoneyTrack.'));
        }
    };
    reader.readAsArrayBuffer(archivo);
}

/** Importa datos desde un archivo Excel (plantilla MoneyTrack) */
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

            const val = (ws, r, c) => {
                const cell = ws[XLSX.utils.encode_cell({ r, c })];
                return cell ? (cell.v !== undefined ? cell.v : '') : '';
            };
            const parseNum = (v) => {
                if (v === '' || v === null || v === undefined) return 0;
                if (typeof v === 'number' && !isNaN(v)) return v;
                const s = String(v).replace(/\s+/g, ' ').trim();
                const sinMoneda = s.replace(/\s*(USD|EUR|MXN|COP|ARS|CLP|PEN|GBP)\s*$/i, '').trim();
                const numStr = sinMoneda.replace(/\./g, '').replace(',', '.');
                const n = parseFloat(numStr);
                return isNaN(n) ? 0 : n;
            };
            const parseFecha = (v) => {
                if (!v) return null;
                if (v instanceof Date) return v;
                if (typeof v === 'number') {
                    const d = new Date((v - 25569) * 86400 * 1000);
                    return isNaN(d.getTime()) ? null : d;
                }
                const s = String(v).trim();
                const m = s.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})(?:\s+(\d{1,2}):(\d{2}))?/);
                if (m) {
                    const dia = parseInt(m[1],10), mes = parseInt(m[2],10)-1, año = parseInt(m[3],10) < 100 ? 2000 + parseInt(m[3],10) : parseInt(m[3],10);
                    const h = m[4] ? parseInt(m[4],10) : 12, min = m[5] ? parseInt(m[5],10) : 0;
                    return new Date(año, mes, dia, h, min);
                }
                const mMes = MESES.findIndex(m => s.includes(m));
                if (mMes >= 0) {
                    const añoMatch = s.match(/(\d{4})/);
                    const año = añoMatch ? parseInt(añoMatch[1],10) : new Date().getFullYear();
                    return new Date(año, mMes, 1, 12, 0);
                }
                return null;
            };
            const cuentaNombreToId = (nombre) => {
                const n = (nombre || '').toString().trim().toLowerCase();
                const map = { 'efectivo':'efectivo','banco':'banco','tarjeta de crédito':'tarjetaCredito','nequi':'nequi','daviplata':'daviplata' };
                for (const [k, id] of Object.entries(map)) if (n.includes(k)) return id;
                return CUENTAS.find(c => c.nombre.toLowerCase() === n)?.id || 'efectivo';
            };

            let moneda = localStorage.getItem('moneda') || '';
            const saldosCuentas = { efectivo:0, banco:0, tarjetaCredito:0, nequi:0, daviplata:0 };

            if (workbook.SheetNames.includes('Resumen')) {
                const ws = workbook.Sheets['Resumen'];
                const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
                for (let r = 0; r <= range.e.r; r++) {
                    const a = val(ws, r, 0);
                    const b = val(ws, r, 1);
                    if (String(a).trim() === 'Moneda' && b) moneda = String(b).trim().split(/\s/)[0] || moneda;
                    if (String(a).trim() === 'Presupuesto mensual' && b) {
                        const n = parseNum(b);
                        if (n > 0) localStorage.setItem('presupuestoMensual', n.toString());
                    }
                    if (String(a).trim() === 'Límite tarjeta de crédito' && b) {
                        const n = parseNum(b);
                        if (n > 0) localStorage.setItem('limiteTarjetaCredito', n.toString());
                    }
                    const nom = String(a).trim();
                    if (CUENTAS.some(c => c.nombre === nom) && b) {
                        const id = cuentaNombreToId(nom);
                        saldosCuentas[id] = parseNum(b);
                    }
                }
                localStorage.setItem('moneda', moneda);
                localStorage.setItem('saldosCuentas', JSON.stringify(saldosCuentas));
            }

            if (workbook.SheetNames.includes('Categorías')) {
                const ws = workbook.Sheets['Categorías'];
                const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
                const cats = [];
                for (let i = 1; i < data.length; i++) {
                    const row = data[i];
                    const nom = (row && row[0]) ? String(row[0]).trim() : '';
                    if (!nom || nom.includes('Sin categorías')) continue;
                    const limite = parseNum(row && row[1]);
                    const color = (row && row[2]) ? String(row[2]).trim() : '#6b7280';
                    cats.push({ nombre: nom, color, limite: limite > 0 ? limite : null });
                }
                localStorage.setItem('categorias', JSON.stringify(cats));
            }

            if (workbook.SheetNames.includes('Metas')) {
                const ws = workbook.Sheets['Metas'];
                const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
                const metas = [];
                for (let i = 1; i < data.length; i++) {
                    const row = data[i];
                    const nom = (row && row[0]) ? String(row[0]).trim() : '';
                    if (!nom || nom.includes('Sin metas')) continue;
                    const id = 'meta_' + Date.now() + '_' + i;
                    metas.push({ id, nombre: nom, objetivo: parseNum(row && row[1]), plazo: (row && row[4]) ? String(row[4]).trim() : null });
                }
                localStorage.setItem('metas', JSON.stringify(metas));
            }

            const mapaMetaNombre = {};
            const metasGuardadas = JSON.parse(localStorage.getItem('metas') || '[]');
            metasGuardadas.forEach(m => { mapaMetaNombre[m.nombre.toLowerCase().trim()] = m.id; });

            if (workbook.SheetNames.includes('Ingresos')) {
                const ws = workbook.Sheets['Ingresos'];
                const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
                const ingresos = [];
                for (let i = 1; i < data.length; i++) {
                    const row = data[i];
                    const fechaVal = row && (row[1] !== undefined) ? row[1] : (row && row[0]);
                    const montoVal = row && (row[3] !== undefined) ? row[3] : (row && row[2]);
                    const notaVal = row && (row[2] !== undefined) ? row[2] : '';
                    const cuentaVal = row && (row[4] !== undefined) ? row[4] : (row && row[3]);
                    if (String(fechaVal || '').includes('═══') || String(notaVal || '').includes('═══')) continue;
                    const d = parseFecha(fechaVal);
                    const monto = parseNum(montoVal);
                    if (!d || monto <= 0) continue;
                    const pad = n => String(n).padStart(2,'0');
                    ingresos.push({
                        cantidad: monto,
                        fecha: d.toISOString().slice(0,16),
                        origen: cuentaNombreToId(cuentaVal),
                        nota: String(notaVal || 'Ingreso').trim()
                    });
                }
                localStorage.setItem('ingresos', JSON.stringify(ingresos));
            }

            if (workbook.SheetNames.includes('Gastos')) {
                const ws = workbook.Sheets['Gastos'];
                const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
                const gastos = [];
                for (let i = 1; i < data.length; i++) {
                    const row = data[i];
                    const fechaVal = row && (row[1] !== undefined) ? row[1] : (row && row[0]);
                    const nombreVal = row && (row[2] !== undefined) ? row[2] : '';
                    const catVal = row && (row[3] !== undefined) ? row[3] : '';
                    const montoVal = row && (row[4] !== undefined) ? row[4] : (row && row[3]);
                    const cuentaVal = row && (row[5] !== undefined) ? row[5] : '';
                    const notaVal = row && (row[6] !== undefined) ? row[6] : '';
                    const cuotasVal = row && (row[7] !== undefined) ? row[7] : '1';
                    if (String(fechaVal || '').includes('═══')) continue;
                    const d = parseFecha(fechaVal);
                    const monto = parseNum(montoVal);
                    if (!d || monto <= 0) continue;
                    const cuotas = parseInt(String(cuotasVal).replace(/\D/g, '') || '1', 10) || 1;
                    const pad = n => String(n).padStart(2,'0');
                    gastos.push({
                        nombre: String(nombreVal || '').trim(),
                        cantidad: monto,
                        fecha: d.toISOString().slice(0,16),
                        categoria: String(catVal || '').trim(),
                        origen: cuentaNombreToId(cuentaVal),
                        nota: String(notaVal || '').trim() || null,
                        cuotas: cuotas,
                        cuotaMensual: monto / cuotas
                    });
                }
                localStorage.setItem('gastos', JSON.stringify(gastos));
            }

            if (workbook.SheetNames.includes('Aportes a metas')) {
                const ws = workbook.Sheets['Aportes a metas'];
                const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
                const contribuciones = [];
                for (let i = 1; i < data.length; i++) {
                    const row = data[i];
                    const fechaVal = row && (row[1] !== undefined) ? row[1] : (row && row[0]);
                    const metaNombre = (row && (row[2] !== undefined) ? row[2] : '').toString().trim().toLowerCase();
                    const montoVal = row && (row[3] !== undefined) ? row[3] : (row && row[2]);
                    const cuentaVal = row && (row[4] !== undefined) ? row[4] : '';
                    if (String(fechaVal || '').includes('═══') || !metaNombre) continue;
                    const metaId = mapaMetaNombre[metaNombre];
                    if (!metaId) continue;
                    const d = parseFecha(fechaVal);
                    const monto = parseNum(montoVal);
                    if (!d || monto <= 0) continue;
                    contribuciones.push({ metaId, cantidad: monto, fecha: d.toISOString().slice(0,10) + 'T12:00:00', origen: cuentaNombreToId(cuentaVal) });
                }
                localStorage.setItem('contribucionesMetas', JSON.stringify(contribuciones));
            }

            if (workbook.SheetNames.includes('Pagos programados')) {
                const ws = workbook.Sheets['Pagos programados'];
                const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
                const pagos = [];
                for (let i = 1; i < data.length; i++) {
                    const row = data[i];
                    const concepto = (row && row[0]) ? String(row[0]).trim() : '';
                    if (!concepto || concepto.includes('Sin pagos')) continue;
                    const monto = parseNum(row && row[1]);
                    const frecuencia = (row && row[2]) ? String(row[2]).trim() : 'mensual';
                    const diaPago = (row && row[3]) ? String(row[3]).trim() : '1';
                    const cuenta = cuentaNombreToId(row && row[4]);
                    const categoria = (row && row[5]) ? String(row[5]).trim() : '';
                    const fechaInicio = (row && row[6]) ? String(row[6]).trim() : new Date().toISOString().slice(0,10);
                    const activo = (row && row[7]) ? !String(row[7]).toLowerCase().includes('no') : true;
                    pagos.push({
                        id: 'pago_' + Date.now() + '_' + i,
                        concepto, monto, frecuencia, diaPago: parseInt(diaPago,10) || 1, cuenta, categoria,
                        fechaInicio: fechaInicio || new Date().toISOString().slice(0,10), activo
                    });
                }
                localStorage.setItem('pagosProgramados', JSON.stringify(pagos));
            }

            alert('¡Importación completada! Se han cargado los datos del Excel. La página se recargará.');
            location.reload();
        } catch (err) {
            console.error(err);
            alert('Error al importar: ' + (err.message || 'Formato de archivo no válido. Usa una plantilla exportada desde MoneyTrack.'));
        }
    };
    reader.readAsArrayBuffer(archivo);
}

/**
 * Importa datos desde un archivo Excel (plantilla MoneyTrack).
 * Reconoce la estructura exportada y rellena el sistema automáticamente.
 */
function importarDesdeExcel(archivo) {
    if (typeof XLSX === 'undefined') {
        alert('Error: La librería Excel no está cargada. Recarga la página e intenta de nuevo.');
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

            const val = (ws, r, c) => {
                const cell = ws[XLSX.utils.encode_cell({ r, c })];
                return cell ? (cell.v !== undefined ? cell.v : '') : '';
            };
            const parseNum = (v) => {
                if (v === '' || v === null || v === undefined) return 0;
                if (typeof v === 'number' && !isNaN(v)) return v;
                const s = String(v).replace(/\s+/g, '').replace(/\./g, '').replace(',', '.');
                const n = parseFloat(s);
                return isNaN(n) ? 0 : n;
            };
            const parseFecha = (v) => {
                if (!v) return null;
                if (v instanceof Date) return v;
                if (typeof v === 'number') {
                    const d = new Date((v - 25569) * 86400 * 1000);
                    return isNaN(d.getTime()) ? null : d;
                }
                const s = String(v).trim();
                const m = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2}))?/);
                if (m) {
                    const d = new Date(parseInt(m[3]), parseInt(m[2])-1, parseInt(m[1]), parseInt(m[4])||0, parseInt(m[5])||0);
                    return isNaN(d.getTime()) ? null : d;
                }
                const d = new Date(s);
                return isNaN(d.getTime()) ? null : d;
            };
            const fechaToStr = (d) => {
                if (!d) return null;
                const p = n => String(n).padStart(2, '0');
                return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}:00`;
            };
            const cuentaNombreToId = (nombre) => {
                const n = (nombre || '').toString().trim().toLowerCase();
                const map = { 'efectivo':'efectivo','banco':'banco','tarjeta de crédito':'tarjetaCredito','nequi':'nequi','daviplata':'daviplata' };
                for (const [k, id] of Object.entries(map)) if (n.includes(k)) return id;
                return CUENTAS.find(c => c.nombre.toLowerCase() === n)?.id || 'efectivo';
            };

            let moneda = '';
            const saldosIni = {};
            let presupuesto = 0, limiteTc = 0;

            if (workbook.SheetNames.includes('Resumen')) {
                const ws = workbook.Sheets['Resumen'];
                for (let r = 0; r < 30; r++) {
                    const a = String(val(ws, r, 0)).trim();
                    const b = val(ws, r, 1);
                    if (a === 'Moneda') moneda = String(b).trim() || moneda;
                    if (a === 'Presupuesto mensual') presupuesto = parseNum(b);
                    if (a === 'Límite tarjeta de crédito') limiteTc = parseNum(b);
                    if (a === 'SALDOS INICIALES POR CUENTA') {
                        for (let i = 1; i <= 5; i++) {
                            const nom = String(val(ws, r + i, 0)).trim();
                            const id = cuentaNombreToId(nom);
                            if (id && nom) saldosIni[id] = parseNum(val(ws, r + i, 1));
                        }
                        break;
                    }
                }
            }

            const ingresos = [];
            if (workbook.SheetNames.includes('Ingresos')) {
                const ws = workbook.Sheets['Ingresos'];
                const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
                for (let r = 1; r <= range.e.r; r++) {
                    const mes = String(val(ws, r, 0));
                    const fecha = val(ws, r, 1);
                    const concepto = String(val(ws, r, 2)).trim();
                    const monto = parseNum(val(ws, r, 3));
                    const cuenta = cuentaNombreToId(val(ws, r, 4));
                    if (mes.includes('═══') || !concepto && monto === 0) continue;
                    const d = parseFecha(fecha) || parseFecha(mes);
                    if (d && monto > 0) ingresos.push({ cantidad: monto, fecha: fechaToStr(d), origen: cuenta, nota: concepto || 'Importado' });
                }
            }

            const gastos = [];
            if (workbook.SheetNames.includes('Gastos')) {
                const ws = workbook.Sheets['Gastos'];
                const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
                for (let r = 1; r <= range.e.r; r++) {
                    const mes = String(val(ws, r, 0));
                    const fecha = val(ws, r, 1);
                    const concepto = String(val(ws, r, 2)).trim();
                    const categoria = String(val(ws, r, 3)).trim();
                    const monto = parseNum(val(ws, r, 4));
                    const cuenta = cuentaNombreToId(val(ws, r, 5));
                    const nota = String(val(ws, r, 6)).trim();
                    const cuotasStr = String(val(ws, r, 7));
                    if (mes.includes('═══') || !concepto && monto === 0) continue;
                    const d = parseFecha(fecha) || parseFecha(mes);
                    if (d && monto > 0) {
                        const cuotas = parseInt(cuotasStr) || 1;
                        gastos.push({
                            nombre: concepto, cantidad: monto, fecha: fechaToStr(d), categoria: categoria || 'Otros',
                            origen: cuenta, nota: nota || null, cuotas: cuotas, cuotaMensual: monto / cuotas
                        });
                    }
                }
            }

            const categorias = [];
            if (workbook.SheetNames.includes('Categorías')) {
                const ws = workbook.Sheets['Categorías'];
                const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
                for (let r = 1; r <= range.e.r; r++) {
                    const nom = String(val(ws, r, 0)).trim();
                    const limite = parseNum(val(ws, r, 1));
                    const color = String(val(ws, r, 2)).trim();
                    if (nom && !nom.toLowerCase().includes('sin categorías')) {
                        categorias.push({ nombre: nom, color: color || '#6b7280', limite: limite > 0 ? limite : null });
                    }
                }
            }

            const metas = [];
            const metaNombreToId = {};
            if (workbook.SheetNames.includes('Metas')) {
                const ws = workbook.Sheets['Metas'];
                const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
                for (let r = 1; r <= range.e.r; r++) {
                    const nom = String(val(ws, r, 0)).trim();
                    const obj = parseNum(val(ws, r, 1));
                    const plazo = String(val(ws, r, 4)).trim();
                    if (nom && !nom.toLowerCase().includes('sin metas')) {
                        const id = 'meta_' + Date.now() + '_' + r;
                        metaNombreToId[nom] = id;
                        metas.push({ id, nombre: nom, objetivo: obj, plazo: plazo || null });
                    }
                }
            }

            const contribuciones = [];
            if (workbook.SheetNames.includes('Aportes a metas')) {
                const ws = workbook.Sheets['Aportes a metas'];
                const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
                for (let r = 1; r <= range.e.r; r++) {
                    const mes = String(val(ws, r, 0));
                    const fecha = val(ws, r, 1);
                    const metaNom = String(val(ws, r, 2)).trim();
                    const monto = parseNum(val(ws, r, 3));
                    const cuenta = cuentaNombreToId(val(ws, r, 4));
                    if (mes.includes('═══') || !metaNom || monto <= 0) continue;
                    const metaId = metaNombreToId[metaNom] || Object.values(metaNombreToId)[0];
                    const d = parseFecha(fecha) || parseFecha(mes);
                    if (d && metaId) contribuciones.push({ metaId, cantidad: monto, fecha: fechaToStr(d), origen: cuenta });
                }
            }

            const pagosProgramados = [];
            if (workbook.SheetNames.includes('Pagos programados')) {
                const ws = workbook.Sheets['Pagos programados'];
                const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
                for (let r = 1; r <= range.e.r; r++) {
                    const concepto = String(val(ws, r, 0)).trim();
                    const monto = parseNum(val(ws, r, 1));
                    const frecuencia = String(val(ws, r, 2)).trim();
                    const diaPago = val(ws, r, 3);
                    const cuenta = cuentaNombreToId(val(ws, r, 4));
                    const categoria = String(val(ws, r, 5)).trim();
                    const fechaInicio = val(ws, r, 6);
                    const activo = String(val(ws, r, 7)).toLowerCase();
                    if (!concepto || concepto.toLowerCase().includes('sin pagos')) continue;
                    const d = parseFecha(fechaInicio);
                    pagosProgramados.push({
                        id: 'pago_' + Date.now() + '_' + r,
                        concepto, monto, frecuencia: frecuencia || 'mensual',
                        diaPago: parseInt(diaPago) || 1, cuenta, categoria,
                        fechaInicio: d ? d.toISOString().slice(0, 10) : new Date().toISOString().slice(0, 10),
                        activo: activo !== 'no', nota: ''
                    });
                }
            }

            if (moneda) localStorage.setItem('moneda', moneda);
            if (presupuesto > 0) localStorage.setItem('presupuestoMensual', presupuesto.toString());
            if (limiteTc > 0) localStorage.setItem('limiteTarjetaCredito', limiteTc.toString());
            if (Object.keys(saldosIni).length > 0) localStorage.setItem('saldosCuentas', JSON.stringify(saldosIni));
            localStorage.setItem('ingresos', JSON.stringify(ingresos));
            localStorage.setItem('gastos', JSON.stringify(gastos));
            localStorage.setItem('categorias', JSON.stringify(categorias));
            localStorage.setItem('metas', JSON.stringify(metas));
            localStorage.setItem('contribucionesMetas', JSON.stringify(contribuciones));
            localStorage.setItem('pagosProgramados', JSON.stringify(pagosProgramados));

            alert('¡Importación completada!\n\n' +
                '• Ingresos: ' + ingresos.length + '\n' +
                '• Gastos: ' + gastos.length + '\n' +
                '• Categorías: ' + categorias.length + '\n' +
                '• Metas: ' + metas.length + '\n' +
                '• Aportes: ' + contribuciones.length + '\n' +
                '• Pagos programados: ' + pagosProgramados.length);
            location.reload();
        } catch (err) {
            console.error(err);
            alert('Error al importar: ' + (err.message || 'Formato de archivo no reconocido. Asegúrate de usar una plantilla exportada de MoneyTrack.'));
        }
    };
    reader.readAsArrayBuffer(archivo);
}

/**
 * Importa datos desde un archivo Excel (plantilla MoneyTrack).
 * Reconoce la estructura exportada y rellena automáticamente el sistema.
 */
function importarDesdeExcel(archivo) {
    if (typeof XLSX === 'undefined') {
        alert('Error: La librería Excel no está cargada. Recarga la página e intenta de nuevo.');
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

            const val = (ws, r, c) => {
                const cell = ws[XLSX.utils.encode_cell({ r, c })];
                return cell ? (cell.v !== undefined ? cell.v : '') : '';
            };
            const parseNum = (v) => {
                if (v === '' || v === null || v === undefined) return 0;
                if (typeof v === 'number' && !isNaN(v)) return v;
                const s = String(v).replace(/\s+[A-Z]{3}$/i, '').replace(/\./g, '').replace(',', '.');
                return parseFloat(s) || 0;
            };
            const parseFecha = (v) => {
                if (!v) return null;
                if (typeof v === 'number') {
                    const d = XLSX.SSF.parse_date_code ? XLSX.SSF.parse_date_code(v) : null;
                    if (d) return `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}T12:00:00`;
                }
                const s = String(v).trim();
                const m = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
                if (m) return `${m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}T12:00:00`;
                const m2 = s.match(/(\d{4})-(\d{2})-(\d{2})/);
                if (m2) return `${m2[1]}-${m2[2]}-${m2[3]}T12:00:00`;
                return null;
            };
            const cuentaNombreToId = (nombre) => {
                const n = (nombre || '').toString().trim().toLowerCase();
                const map = { 'efectivo':'efectivo','banco':'banco','tarjeta de crédito':'tarjetaCredito','nequi':'nequi','daviplata':'daviplata' };
                for (const [k, id] of Object.entries(map)) if (n.includes(k)) return id;
                return CUENTAS.find(c => c.nombre.toLowerCase() === n)?.id || 'efectivo';
            };
            const esFilaDatos = (txt) => txt && !String(txt).includes('═══') && !String(txt).match(/^(MES|FECHA|CONCEPTO|CATEGORÍA|META|MONTO|CUENTA|OBJETIVO|AHORRADO|FRECUENCIA|DÍA|ACTIVO)$/i);

            let moneda = localStorage.getItem('moneda') || '';
            const saldosIni = {};
            CUENTAS.forEach(c => { saldosIni[c.id] = 0; });

            if (workbook.SheetNames.includes('Resumen')) {
                const ws = workbook.Sheets['Resumen'];
                for (let r = 0; r < 50; r++) {
                    const a = val(ws, r, 0), b = val(ws, r, 1);
                    if (String(a).trim() === 'Moneda' && b) moneda = String(b).trim().split(/\s/)[0] || moneda;
                    if (String(a).trim() === 'Presupuesto mensual' && b) {
                        const num = parseNum(b);
                        if (num > 0) localStorage.setItem('presupuestoMensual', num.toString());
                    }
                    if (String(a).trim() === 'Límite tarjeta de crédito' && b) {
                        const num = parseNum(b);
                        if (num > 0) localStorage.setItem('limiteTarjetaCredito', num.toString());
                    }
                    if (String(a).trim() === 'SALDOS INICIALES POR CUENTA') {
                        for (let j = 1; j <= 10; j++) {
                            const nom = val(ws, r + j, 0), monto = val(ws, r + j, 1);
                            if (!nom || String(nom).includes('SALDO') || String(nom).includes('CONFIG')) break;
                            const id = cuentaNombreToId(nom);
                            saldosIni[id] = parseNum(monto);
                        }
                        break;
                    }
                }
                localStorage.setItem('moneda', moneda);
                localStorage.setItem('saldosCuentas', JSON.stringify(saldosIni));
            }

            if (workbook.SheetNames.includes('Ingresos')) {
                const ws = workbook.Sheets['Ingresos'];
                const ingresos = [];
                const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
                for (let r = 1; r <= range.e.r; r++) {
                    const fechaTxt = val(ws, r, 1), concepto = val(ws, r, 2), monto = val(ws, r, 3), cuenta = val(ws, r, 4);
                    if (!esFilaDatos(concepto) && !esFilaDatos(fechaTxt)) continue;
                    if (parseNum(monto) <= 0) continue;
                    const fecha = parseFecha(fechaTxt);
                    if (!fecha) continue;
                    ingresos.push({ cantidad: parseNum(monto), fecha, origen: cuentaNombreToId(cuenta), nota: (concepto || 'Ingreso').toString().trim() });
                }
                localStorage.setItem('ingresos', JSON.stringify(ingresos));
            }

            if (workbook.SheetNames.includes('Gastos')) {
                const ws = workbook.Sheets['Gastos'];
                const gastos = [];
                const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
                for (let r = 1; r <= range.e.r; r++) {
                    const fechaTxt = val(ws, r, 1), concepto = val(ws, r, 2), cat = val(ws, r, 3), monto = val(ws, r, 4), cuenta = val(ws, r, 5), nota = val(ws, r, 6), cuotasTxt = val(ws, r, 7);
                    if (!esFilaDatos(concepto) && !esFilaDatos(fechaTxt)) continue;
                    if (parseNum(monto) <= 0) continue;
                    const fecha = parseFecha(fechaTxt);
                    if (!fecha) continue;
                    const cuotas = parseInt(String(cuotasTxt).replace(/\D/g, '')) || 1;
                    const cantidad = parseNum(monto);
                    gastos.push({ nombre: (concepto||'').toString().trim(), cantidad, fecha, categoria: (cat||'').toString().trim(), origen: cuentaNombreToId(cuenta), nota: (nota||'').toString().trim() || null, cuotas: cuotas, cuotaMensual: cantidad / cuotas });
                }
                localStorage.setItem('gastos', JSON.stringify(gastos));
            }

            if (workbook.SheetNames.includes('Categorías')) {
                const ws = workbook.Sheets['Categorías'];
                const categorias = [];
                const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
                for (let r = 1; r <= range.e.r; r++) {
                    const nom = val(ws, r, 0), limite = val(ws, r, 1), color = val(ws, r, 2);
                    if (!nom || String(nom).toLowerCase().includes('categoría') || String(nom).toLowerCase().includes('sin categorías')) continue;
                    categorias.push({ nombre: String(nom).trim(), limite: parseNum(limite) || null, color: (color||'').toString().trim() || undefined });
                }
                localStorage.setItem('categorias', JSON.stringify(categorias));
            }

            const metasImportadas = [];
            if (workbook.SheetNames.includes('Metas')) {
                const ws = workbook.Sheets['Metas'];
                const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
                for (let r = 1; r <= range.e.r; r++) {
                    const nom = val(ws, r, 0), obj = val(ws, r, 1), plazo = val(ws, r, 4);
                    if (!nom || String(nom).toLowerCase().includes('meta') || String(nom).toLowerCase().includes('sin metas')) continue;
                    const id = 'meta_' + Date.now() + '_' + r;
                    metasImportadas.push({ id, nombre: String(nom).trim(), objetivo: parseNum(obj), plazo: (plazo||'—').toString().trim() || null });
                }
                localStorage.setItem('metas', JSON.stringify(metasImportadas));
            }

            if (workbook.SheetNames.includes('Aportes a metas') && metasImportadas.length > 0) {
                const ws = workbook.Sheets['Aportes a metas'];
                const contribuciones = [];
                const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
                for (let r = 1; r <= range.e.r; r++) {
                    const fechaTxt = val(ws, r, 1), metaNom = val(ws, r, 2), monto = val(ws, r, 3), cuenta = val(ws, r, 4);
                    if (!esFilaDatos(metaNom) && !esFilaDatos(fechaTxt)) continue;
                    if (parseNum(monto) <= 0) continue;
                    const fecha = parseFecha(fechaTxt);
                    if (!fecha) continue;
                    const meta = metasImportadas.find(m => m.nombre === String(metaNom).trim());
                    if (meta) contribuciones.push({ metaId: meta.id, cantidad: parseNum(monto), fecha, origen: cuentaNombreToId(cuenta) });
                }
                localStorage.setItem('contribucionesMetas', JSON.stringify(contribuciones));
            } else {
                localStorage.setItem('contribucionesMetas', '[]');
            }

            if (workbook.SheetNames.includes('Pagos programados')) {
                const ws = workbook.Sheets['Pagos programados'];
                const pagos = [];
                const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
                for (let r = 1; r <= range.e.r; r++) {
                    const concepto = val(ws, r, 0), monto = val(ws, r, 1), freq = val(ws, r, 2), dia = val(ws, r, 3), cuenta = val(ws, r, 4), cat = val(ws, r, 5), fechaIni = val(ws, r, 6), activo = val(ws, r, 7);
                    if (!concepto || String(concepto).toLowerCase().includes('concepto') || String(concepto).toLowerCase().includes('sin pagos')) continue;
                    if (parseNum(monto) <= 0) continue;
                    pagos.push({ id: 'pago_' + Date.now() + '_' + r, concepto: String(concepto).trim(), monto: parseNum(monto), frecuencia: (freq||'mensual').toString().toLowerCase(), diaPago: parseInt(dia) || 1, cuenta: cuentaNombreToId(cuenta), categoria: (cat||'').toString().trim(), fechaInicio: parseFecha(fechaIni) || new Date().toISOString().slice(0,10), activo: String(activo).toLowerCase() !== 'no', nota: '' });
                }
                localStorage.setItem('pagosProgramados', JSON.stringify(pagos));
            }

            alert('¡Importación completada! Se han cargado los datos del Excel. La página se recargará.');
            location.reload();
        } catch (err) {
            console.error(err);
            alert('Error al importar: ' + (err.message || 'Formato de archivo no reconocido. Asegúrate de usar una plantilla exportada desde MoneyTrack.'));
        }
    };
    reader.readAsArrayBuffer(archivo);
}

/**
 * Importa datos desde un archivo Excel (plantilla MoneyTrack).
 * Reconoce la estructura exportada y rellena el sistema automáticamente.
 */
function importarDesdeExcel(archivo) {
    if (typeof XLSX === 'undefined') {
        alert('Error: La librería Excel no está cargada. Recarga la página e intenta de nuevo.');
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

            const valorCelda = (ws, r, c) => {
                const cell = ws[XLSX.utils.encode_cell({ r, c })];
                return cell ? (cell.v !== undefined ? cell.v : '') : '';
            };

            const parseNum = (v) => {
                if (v === '' || v === null || v === undefined) return 0;
                if (typeof v === 'number' && !isNaN(v)) return v;
                const s = String(v).replace(/\s+/g, ' ').trim();
                const sinMoneda = s.replace(/\s*(USD|EUR|MXN|COP|ARS|CLP|PEN|GBP)\s*$/i, '').trim();
                const numStr = sinMoneda.replace(/\./g, '').replace(',', '.');
                const n = parseFloat(numStr);
                return isNaN(n) ? 0 : n;
            };

            const parseFecha = (v) => {
                if (!v) return null;
                if (v instanceof Date) return v;
                if (typeof v === 'number') {
                    const d = XLSX.SSF.parse_date_code ? XLSX.SSF.parse_date_code(v) : null;
                    if (d) return new Date(d.y, d.m - 1, d.d, d.H || 0, d.M || 0);
                    return new Date((v - 25569) * 86400 * 1000);
                }
                const s = String(v).trim();
                const m = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2}))?/);
                if (m) {
                    const [, d, mes, año, h = 12, min = 0] = m;
                    return new Date(parseInt(año), parseInt(mes) - 1, parseInt(d), parseInt(h), parseInt(min));
                }
                const d = new Date(s);
                return isNaN(d.getTime()) ? null : d;
            };

            const fechaToStr = (d) => {
                if (!d || !(d instanceof Date) || isNaN(d.getTime())) return null;
                const p = n => String(n).padStart(2, '0');
                return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}`;
            };

            const cuentaNombreToId = (nombre) => {
                const n = (nombre || '').toString().trim().toLowerCase();
                const map = { 'efectivo': 'efectivo', 'banco': 'banco', 'tarjeta de crédito': 'tarjetaCredito', 'nequi': 'nequi', 'daviplata': 'daviplata' };
                for (const [k, id] of Object.entries(map)) {
                    if (n.includes(k)) return id;
                }
                return CUENTAS.find(c => c.nombre.toLowerCase() === n)?.id || 'efectivo';
            };

            const mesTextoToDate = (texto) => {
                const t = (texto || '').toString();
                for (let i = 0; i < MESES.length; i++) {
                    if (t.includes(MESES[i])) {
                        const año = t.match(/\d{4}/);
                        return new Date(parseInt(año) || new Date().getFullYear(), i, 15, 12, 0);
                    }
                }
                return null;
            };

            let moneda = localStorage.getItem('moneda') || 'COP';
            const saldosIni = {};
            CUENTAS.forEach(c => { saldosIni[c.id] = 0; });

            const wsResumen = workbook.Sheets['Resumen'] || workbook.Sheets[workbook.SheetNames[0]];
            if (wsResumen) {
                for (let r = 0; r < 50; r++) {
                    const a = valorCelda(wsResumen, r, 0);
                    const b = valorCelda(wsResumen, r, 1);
                    const aStr = String(a || '').toLowerCase();
                    if (aStr === 'moneda' && b) moneda = String(b).trim().split(/\s/)[0] || moneda;
                    if (aStr === 'presupuesto mensual' && b) {
                        const val = parseNum(b);
                        if (val > 0) localStorage.setItem('presupuestoMensual', val.toString());
                    }
                    if (aStr.includes('límite tarjeta') && b) {
                        const val = parseNum(b);
                        if (val > 0) localStorage.setItem('limiteTarjetaCredito', val.toString());
                    }
                    if (aStr === 'saldos iniciales por cuenta') break;
                }
                let enSaldosIni = false;
                for (let r = 0; r < 50; r++) {
                    const a = valorCelda(wsResumen, r, 0);
                    const b = valorCelda(wsResumen, r, 1);
                    const aStr = String(a || '').trim();
                    if (aStr === 'SALDOS INICIALES POR CUENTA') { enSaldosIni = true; continue; }
                    if (enSaldosIni && aStr === 'SALDO ACTUAL POR CUENTA') break;
                    if (enSaldosIni && aStr && b) {
                        const id = cuentaNombreToId(aStr);
                        saldosIni[id] = parseNum(b);
                    }
                }
                localStorage.setItem('moneda', moneda);
                localStorage.setItem('saldosCuentas', JSON.stringify(saldosIni));
            }

            const ingresos = [];
            const wsIng = workbook.Sheets['Ingresos'];
            if (wsIng) {
                const range = XLSX.utils.decode_range(wsIng['!ref'] || 'A1');
                for (let r = 1; r <= range.e.r; r++) {
                    const mesCol = valorCelda(wsIng, r, 0);
                    const fechaCol = valorCelda(wsIng, r, 1);
                    const concepto = valorCelda(wsIng, r, 2);
                    const monto = valorCelda(wsIng, r, 3);
                    const cuenta = valorCelda(wsIng, r, 4);
                    if (String(mesCol || '').includes('═══') || !monto) continue;
                    const d = parseFecha(fechaCol) || mesTextoToDate(mesCol);
                    if (!d) continue;
                    ingresos.push({
                        cantidad: parseNum(monto),
                        fecha: fechaToStr(d),
                        origen: cuentaNombreToId(cuenta),
                        nota: (concepto || 'Ingreso').toString().trim()
                    });
                }
            }
            localStorage.setItem('ingresos', JSON.stringify(ingresos));

            const gastos = [];
            const wsGas = workbook.Sheets['Gastos'];
            if (wsGas) {
                const range = XLSX.utils.decode_range(wsGas['!ref'] || 'A1');
                for (let r = 1; r <= range.e.r; r++) {
                    const mesCol = valorCelda(wsGas, r, 0);
                    const fechaCol = valorCelda(wsGas, r, 1);
                    const concepto = valorCelda(wsGas, r, 2);
                    const categoria = valorCelda(wsGas, r, 3);
                    const monto = valorCelda(wsGas, r, 4);
                    const cuenta = valorCelda(wsGas, r, 5);
                    const nota = valorCelda(wsGas, r, 6);
                    const cuotasStr = valorCelda(wsGas, r, 7);
                    if (String(mesCol || '').includes('═══') || !monto) continue;
                    const d = parseFecha(fechaCol) || mesTextoToDate(mesCol);
                    if (!d) continue;
                    const cuotas = parseInt(String(cuotasStr).replace(/\D/g, '')) || 1;
                    const cant = parseNum(monto);
                    gastos.push({
                        nombre: (concepto || '').toString().trim(),
                        cantidad: cant,
                        fecha: fechaToStr(d),
                        categoria: (categoria || '').toString().trim(),
                        origen: cuentaNombreToId(cuenta),
                        nota: (nota || '').toString().trim() || null,
                        cuotas: cuotas,
                        cuotaMensual: cuotas > 1 ? cant / cuotas : cant
                    });
                }
            }
            localStorage.setItem('gastos', JSON.stringify(gastos));

            const categorias = [];
            const wsCat = workbook.Sheets['Categorías'];
            if (wsCat) {
                const range = XLSX.utils.decode_range(wsCat['!ref'] || 'A1');
                for (let r = 1; r <= range.e.r; r++) {
                    const nom = valorCelda(wsCat, r, 0);
                    const limite = valorCelda(wsCat, r, 1);
                    const color = valorCelda(wsCat, r, 2);
                    if (!nom || String(nom).toLowerCase().includes('sin categorías')) continue;
                    categorias.push({
                        nombre: String(nom).trim(),
                        limite: parseNum(limite) || null,
                        color: (color || '').toString().trim() || undefined
                    });
                }
            }
            localStorage.setItem('categorias', JSON.stringify(categorias));

            const metas = [];
            const metaNombreToId = {};
            const wsMet = workbook.Sheets['Metas'];
            if (wsMet) {
                const range = XLSX.utils.decode_range(wsMet['!ref'] || 'A1');
                for (let r = 1; r <= range.e.r; r++) {
                    const nom = valorCelda(wsMet, r, 0);
                    const obj = valorCelda(wsMet, r, 1);
                    const plazo = valorCelda(wsMet, r, 4);
                    if (!nom || String(nom).toLowerCase().includes('sin metas')) continue;
                    const id = 'meta_' + Date.now() + '_' + r;
                    metaNombreToId[String(nom).trim()] = id;
                    metas.push({
                        id,
                        nombre: String(nom).trim(),
                        objetivo: parseNum(obj),
                        plazo: (plazo || '').toString().trim() || null
                    });
                }
            }
            localStorage.setItem('metas', JSON.stringify(metas));

            const contribuciones = [];
            const wsAport = workbook.Sheets['Aportes a metas'];
            if (wsAport) {
                const range = XLSX.utils.decode_range(wsAport['!ref'] || 'A1');
                for (let r = 1; r <= range.e.r; r++) {
                    const mesCol = valorCelda(wsAport, r, 0);
                    const fechaCol = valorCelda(wsAport, r, 1);
                    const metaNom = valorCelda(wsAport, r, 2);
                    const monto = valorCelda(wsAport, r, 3);
                    const cuenta = valorCelda(wsAport, r, 4);
                    if (String(mesCol || '').includes('═══') || !monto) continue;
                    const metaId = metaNombreToId[String(metaNom).trim()];
                    if (!metaId) continue;
                    const d = parseFecha(fechaCol) || mesTextoToDate(mesCol);
                    if (!d) continue;
                    contribuciones.push({
                        metaId,
                        cantidad: parseNum(monto),
                        fecha: fechaToStr(d),
                        origen: cuentaNombreToId(cuenta)
                    });
                }
            }
            localStorage.setItem('contribucionesMetas', JSON.stringify(contribuciones));

            const pagosProgramados = [];
            const wsPagos = workbook.Sheets['Pagos programados'];
            if (wsPagos) {
                const range = XLSX.utils.decode_range(wsPagos['!ref'] || 'A1');
                for (let r = 1; r <= range.e.r; r++) {
                    const concepto = valorCelda(wsPagos, r, 0);
                    const monto = valorCelda(wsPagos, r, 1);
                    const frecuencia = valorCelda(wsPagos, r, 2);
                    const diaPago = valorCelda(wsPagos, r, 3);
                    const cuenta = valorCelda(wsPagos, r, 4);
                    const categoria = valorCelda(wsPagos, r, 5);
                    const fechaInicio = valorCelda(wsPagos, r, 6);
                    const activo = valorCelda(wsPagos, r, 7);
                    if (!concepto || String(concepto).toLowerCase().includes('sin pagos')) continue;
                    const d = parseFecha(fechaInicio);
                    pagosProgramados.push({
                        id: 'pago_' + Date.now() + '_' + r,
                        concepto: String(concepto).trim(),
                        monto: parseNum(monto),
                        frecuencia: (frecuencia || 'mensual').toString().toLowerCase(),
                        diaPago: parseInt(diaPago) || 1,
                        cuenta: cuentaNombreToId(cuenta),
                        categoria: (categoria || '').toString().trim(),
                        fechaInicio: d ? fechaToStr(d).slice(0, 10) : null,
                        activo: String(activo).toLowerCase() !== 'no',
                        nota: ''
                    });
                }
            }
            localStorage.setItem('pagosProgramados', JSON.stringify(pagosProgramados));

            alert('¡Importación completada! Se han cargado: ' + ingresos.length + ' ingresos, ' + gastos.length + ' gastos, ' + categorias.length + ' categorías, ' + metas.length + ' metas, ' + contribuciones.length + ' aportes y ' + pagosProgramados.length + ' pagos programados.');
            location.reload();
        } catch (err) {
            console.error(err);
            alert('Error al importar: ' + (err.message || 'Formato de archivo no reconocido. Asegúrate de usar una plantilla exportada de MoneyTrack.'));
        }
    };
    reader.readAsArrayBuffer(archivo);
}

/**
 * Importa datos desde un archivo Excel (plantilla MoneyTrack).
 * Reconoce la estructura exportada y rellena el sistema automáticamente.
 */
function importarDesdeExcel(archivo) {
    if (typeof XLSX === 'undefined') {
        alert('Error: La librería Excel no está cargada. Recarga la página e intenta de nuevo.');
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

            const val = (ws, r, c) => {
                const cell = ws[XLSX.utils.encode_cell({ r, c })];
                return cell ? (cell.v !== undefined ? cell.v : '') : '';
            };

            const parseNum = (v) => {
                if (v === '' || v === null || v === undefined) return 0;
                if (typeof v === 'number' && !isNaN(v)) return v;
                const s = String(v).trim().replace(/\s+[A-Z]{3}$/i, '').replace(/\./g, '').replace(',', '.');
                const n = parseFloat(s);
                return isNaN(n) ? 0 : n;
            };

            const parseFecha = (v) => {
                if (!v) return null;
                if (v instanceof Date) return v;
                if (typeof v === 'number') {
                    const d = XLSX.SSF.parse_date_code ? XLSX.SSF.parse_date_code(v) : null;
                    if (d) return new Date(d.y, d.m - 1, d.d, d.H || 0, d.M || 0);
                    return new Date((v - 25569) * 86400 * 1000);
                }
                const s = String(v).trim();
                const m = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2}))?/);
                if (m) return new Date(+m[3], +m[2] - 1, +m[1], +(m[4]||0), +(m[5]||0));
                return new Date(s) || null;
            };

            const fechaToStr = (d) => {
                if (!d || !(d instanceof Date) || isNaN(d)) return null;
                const p = n => String(n).padStart(2, '0');
                return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}:00`;
            };

            const cuentaNombreToId = (nombre) => {
                const n = (nombre || '').toString().trim().toLowerCase();
                const map = { 'efectivo':'efectivo','banco':'banco','tarjeta de crédito':'tarjetaCredito','nequi':'nequi','daviplata':'daviplata' };
                for (const [k, id] of Object.entries(map)) if (n.includes(k)) return id;
                return CUENTAS.find(c => c.nombre.toLowerCase() === n)?.id || 'efectivo';
            };

            const esFilaDatos = (txt) => txt && !String(txt).includes('═══') && !String(txt).match(/^(MES|FECHA|CONCEPTO|CATEGORÍA|META|MONTO|CUENTA|OBJETIVO|AHORRADO|% LOGRADO|PLAZO|Nº APORTES|ACTIVO|FRECUENCIA|DÍA PAGO|FECHA INICIO|COLOR|LÍMITE)/i);

            let moneda = 'COP';
            const saldosIni = {};

            if (workbook.SheetNames.includes('Resumen')) {
                const ws = workbook.Sheets['Resumen'];
                for (let r = 0; r < 50; r++) {
                    const a = val(ws, r, 0), b = val(ws, r, 1);
                    const label = String(a || '').trim();
                    if (label === 'Moneda' && b) moneda = String(b).trim().split(/\s/)[0] || moneda;
                    if (label === 'Presupuesto mensual' && b) {
                        const n = parseNum(b);
                        if (n > 0) localStorage.setItem('presupuestoMensual', n.toString());
                    }
                    if (label === 'Límite tarjeta de crédito' && b) {
                        const n = parseNum(b);
                        if (n > 0) localStorage.setItem('limiteTarjetaCredito', n.toString());
                    }
                    if (label && CUENTAS.some(c => c.nombre === label) && b) {
                        saldosIni[CUENTAS.find(c => c.nombre === label).id] = parseNum(b);
                    }
                }
                if (Object.keys(saldosIni).length > 0) {
                    const sc = CUENTAS.reduce((acc, c) => { acc[c.id] = saldosIni[c.id] || 0; return acc; }, {});
                    localStorage.setItem('saldosCuentas', JSON.stringify(sc));
                }
                localStorage.setItem('moneda', moneda);
            }

            if (workbook.SheetNames.includes('Ingresos')) {
                const ws = workbook.Sheets['Ingresos'];
                const ingresos = [];
                const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
                for (let r = 1; r <= range.e.r; r++) {
                    const fechaVal = val(ws, r, 1), concepto = val(ws, r, 2), montoVal = val(ws, r, 3), cuentaVal = val(ws, r, 4);
                    if (!esFilaDatos(concepto) && !fechaVal && !montoVal) continue;
                    if (String(concepto || '').includes('Sin ingresos')) break;
                    const monto = parseNum(montoVal);
                    if (monto <= 0) continue;
                    const d = parseFecha(fechaVal);
                    if (!d) continue;
                    ingresos.push({
                        cantidad: monto,
                        fecha: fechaToStr(d),
                        origen: cuentaNombreToId(cuentaVal),
                        nota: (concepto || 'Ingreso').toString().trim()
                    });
                }
                localStorage.setItem('ingresos', JSON.stringify(ingresos));
            }

            if (workbook.SheetNames.includes('Gastos')) {
                const ws = workbook.Sheets['Gastos'];
                const gastos = [];
                const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
                for (let r = 1; r <= range.e.r; r++) {
                    const fechaVal = val(ws, r, 1), concepto = val(ws, r, 2), cat = val(ws, r, 3), montoVal = val(ws, r, 4), cuentaVal = val(ws, r, 5), nota = val(ws, r, 6), cuotasStr = val(ws, r, 7);
                    if (!esFilaDatos(concepto) && !fechaVal && !montoVal) continue;
                    if (String(concepto || '').includes('Sin gastos')) break;
                    const monto = parseNum(montoVal);
                    if (monto <= 0) continue;
                    const d = parseFecha(fechaVal);
                    if (!d) continue;
                    const cuotas = parseInt(String(cuotasStr).replace(/\D/g, '')) || 1;
                    const cuotaMensual = cuotas > 1 ? monto / cuotas : monto;
                    gastos.push({
                        nombre: (concepto || '').toString().trim(),
                        cantidad: monto,
                        fecha: fechaToStr(d),
                        categoria: (cat || '').toString().trim(),
                        origen: cuentaNombreToId(cuentaVal),
                        nota: (nota || '').toString().trim() || null,
                        cuotas: cuotas,
                        cuotaMensual: cuotaMensual
                    });
                }
                localStorage.setItem('gastos', JSON.stringify(gastos));
            }

            if (workbook.SheetNames.includes('Categorías')) {
                const ws = workbook.Sheets['Categorías'];
                const categorias = [];
                const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
                for (let r = 1; r <= range.e.r; r++) {
                    const nom = val(ws, r, 0), limiteVal = val(ws, r, 1), color = val(ws, r, 2);
                    if (!nom || String(nom).includes('Sin categorías')) break;
                    categorias.push({
                        nombre: String(nom).trim(),
                        limite: parseNum(limiteVal) || null,
                        color: (color || '').toString().trim() || '#6b7280'
                    });
                }
                localStorage.setItem('categorias', JSON.stringify(categorias));
            }

            const mapaMetaNombreId = {};
            if (workbook.SheetNames.includes('Metas')) {
                const ws = workbook.Sheets['Metas'];
                const metas = [];
                const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
                for (let r = 1; r <= range.e.r; r++) {
                    const nom = val(ws, r, 0), objVal = val(ws, r, 1), plazo = val(ws, r, 4);
                    if (!nom || String(nom).includes('Sin metas')) break;
                    const id = 'meta_' + Date.now() + '_' + r;
                    mapaMetaNombreId[String(nom).trim()] = id;
                    metas.push({
                        id,
                        nombre: String(nom).trim(),
                        objetivo: parseNum(objVal),
                        plazo: (plazo || '').toString().trim() || null
                    });
                }
                localStorage.setItem('metas', JSON.stringify(metas));
            }

            if (workbook.SheetNames.includes('Aportes a metas') && Object.keys(mapaMetaNombreId).length > 0) {
                const ws = workbook.Sheets['Aportes a metas'];
                const contribuciones = [];
                const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
                for (let r = 1; r <= range.e.r; r++) {
                    const fechaVal = val(ws, r, 1), metaNom = val(ws, r, 2), montoVal = val(ws, r, 3), cuentaVal = val(ws, r, 4);
                    if (!metaNom || String(metaNom).includes('Sin aportes')) break;
                    if (String(metaNom).includes('═══')) continue;
                    const metaId = mapaMetaNombreId[String(metaNom).trim()];
                    if (!metaId) continue;
                    const monto = parseNum(montoVal);
                    if (monto <= 0) continue;
                    const d = parseFecha(fechaVal);
                    if (!d) continue;
                    contribuciones.push({
                        metaId,
                        cantidad: monto,
                        fecha: fechaToStr(d),
                        origen: cuentaNombreToId(cuentaVal)
                    });
                }
                localStorage.setItem('contribucionesMetas', JSON.stringify(contribuciones));
            } else if (workbook.SheetNames.includes('Metas')) {
                localStorage.setItem('contribucionesMetas', '[]');
            }

            if (workbook.SheetNames.includes('Pagos programados')) {
                const ws = workbook.Sheets['Pagos programados'];
                const pagos = [];
                const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
                for (let r = 1; r <= range.e.r; r++) {
                    const concepto = val(ws, r, 0), montoVal = val(ws, r, 1), freq = val(ws, r, 2), dia = val(ws, r, 3), cuentaVal = val(ws, r, 4), cat = val(ws, r, 5), fechaIni = val(ws, r, 6), activo = val(ws, r, 7);
                    if (!concepto || String(concepto).includes('Sin pagos')) break;
                    const monto = parseNum(montoVal);
                    if (monto <= 0) continue;
                    const d = parseFecha(fechaIni);
                    pagos.push({
                        id: 'pago_' + Date.now() + '_' + r,
                        concepto: String(concepto).trim(),
                        monto,
                        frecuencia: (freq || 'mensual').toString().toLowerCase(),
                        diaPago: parseInt(dia) || 1,
                        cuenta: cuentaNombreToId(cuentaVal),
                        categoria: (cat || '').toString().trim(),
                        fechaInicio: d ? fechaToStr(d).slice(0, 10) : new Date().toISOString().slice(0, 10),
                        activo: String(activo || 'Sí').toLowerCase() !== 'no',
                        nota: ''
                    });
                }
                localStorage.setItem('pagosProgramados', JSON.stringify(pagos));
            }

            alert('¡Importación completada! Los datos se han cargado correctamente.');
            location.reload();
        } catch (err) {
            console.error(err);
            alert('Error al importar: ' + (err.message || 'Formato de archivo no reconocido. Asegúrate de usar una plantilla exportada desde MoneyTrack.'));
        }
    };
    reader.readAsArrayBuffer(archivo);
}

/**
 * Importa datos desde un archivo Excel (plantilla MoneyTrack).
 * Reconoce la estructura exportada y rellena automáticamente el sistema.
 */
function importarDesdeExcel(archivo) {
    if (typeof XLSX === 'undefined') {
        alert('Error: La librería Excel no está cargada. Recarga la página e intenta de nuevo.');
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

            const val = (ws, r, c) => {
                const cell = ws[XLSX.utils.encode_cell({ r, c })];
                return cell ? (cell.v !== undefined ? cell.v : '') : '';
            };
            const parseNum = (v) => {
                if (v === '' || v === null || v === undefined) return 0;
                if (typeof v === 'number' && !isNaN(v)) return v;
                const s = String(v).replace(/\s+/g, ' ').trim();
                const numStr = s.replace(/[^\d,.\-]/g, '').replace(/\./g, '').replace(',', '.');
                return parseFloat(numStr) || 0;
            };
            const parseFecha = (v) => {
                if (!v) return null;
                if (v instanceof Date) return v;
                if (typeof v === 'number') {
                    const d = XLSX.SSF.parse_date_code ? XLSX.SSF.parse_date_code(v) : null;
                    if (d) return new Date(d.y, d.m - 1, d.d, d.H || 0, d.M || 0, d.S || 0);
                    return new Date((v - 25569) * 86400 * 1000);
                }
                const s = String(v).trim();
                const m = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2}))?/);
                if (m) {
                    const [, d, mo, y, h, min] = m;
                    return new Date(y, parseInt(mo, 10) - 1, parseInt(d, 10), parseInt(h || 0, 10), parseInt(min || 0, 10));
                }
                const d = new Date(s);
                return isNaN(d.getTime()) ? null : d;
            };
            const cuentaNombreToId = (nombre) => {
                const n = (nombre || '').toString().trim().toLowerCase();
                const map = { 'efectivo': 'efectivo', 'banco': 'banco', 'tarjeta de crédito': 'tarjetaCredito', 'nequi': 'nequi', 'daviplata': 'daviplata' };
                return map[n] || CUENTAS.find(c => c.nombre.toLowerCase() === n)?.id || 'efectivo';
            };
            const fechaToStr = (d) => {
                if (!d || !(d instanceof Date) || isNaN(d.getTime())) return null;
                const p = n => String(n).padStart(2, '0');
                return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}:00`;
            };

            let moneda = localStorage.getItem('moneda') || '';
            const saldosCuentas = {};
            CUENTAS.forEach(c => { saldosCuentas[c.id] = 0; });

            if (workbook.SheetNames.includes('Resumen')) {
                const ws = workbook.Sheets['Resumen'];
                for (let r = 0; r < 50; r++) {
                    const a = String(val(ws, r, 0)).trim();
                    const b = val(ws, r, 1);
                    if (a === 'Moneda' && b) moneda = String(b).trim();
                    if (a === 'Presupuesto mensual' && b) {
                        const n = parseNum(b);
                        if (n > 0) localStorage.setItem('presupuestoMensual', n.toString());
                    }
                    if (a === 'Límite tarjeta de crédito' && b) {
                        const n = parseNum(b);
                        if (n > 0) localStorage.setItem('limiteTarjetaCredito', n.toString());
                    }
                    if (['Efectivo','Banco','Tarjeta de crédito','Nequi','Daviplata'].includes(a)) {
                        const id = cuentaNombreToId(a);
                        saldosCuentas[id] = parseNum(b);
                    }
                }
                localStorage.setItem('moneda', moneda);
                localStorage.setItem('saldosCuentas', JSON.stringify(saldosCuentas));
            }

            if (workbook.SheetNames.includes('Ingresos')) {
                const ws = workbook.Sheets['Ingresos'];
                const ingresos = [];
                const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
                for (let r = 1; r <= range.e.r; r++) {
                    const mesVal = val(ws, r, 0);
                    const fechaVal = val(ws, r, 1);
                    const concepto = val(ws, r, 2);
                    const monto = parseNum(val(ws, r, 3));
                    const cuenta = val(ws, r, 4);
                    if (String(mesVal).includes('═══') || !fechaVal || monto <= 0) continue;
                    const d = parseFecha(fechaVal);
                    if (!d) continue;
                    ingresos.push({
                        cantidad: monto,
                        fecha: fechaToStr(d),
                        origen: cuentaNombreToId(cuenta),
                        nota: concepto || 'Ingreso importado'
                    });
                }
                localStorage.setItem('ingresos', JSON.stringify(ingresos));
            }

            if (workbook.SheetNames.includes('Gastos')) {
                const ws = workbook.Sheets['Gastos'];
                const gastos = [];
                const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
                for (let r = 1; r <= range.e.r; r++) {
                    const fechaVal = val(ws, r, 1);
                    const concepto = val(ws, r, 2);
                    const categoria = val(ws, r, 3);
                    const monto = parseNum(val(ws, r, 4));
                    const cuenta = val(ws, r, 5);
                    const nota = val(ws, r, 6);
                    const cuotasStr = val(ws, r, 7);
                    if (String(val(ws, r, 0)).includes('═══') || !fechaVal || monto <= 0) continue;
                    const d = parseFecha(fechaVal);
                    if (!d) continue;
                    const cuotas = parseInt(String(cuotasStr).replace(/\D/g, ''), 10) || 1;
                    const cuotaMensual = cuotas > 1 ? monto / cuotas : monto;
                    gastos.push({
                        nombre: concepto || 'Gasto importado',
                        cantidad: monto,
                        fecha: fechaToStr(d),
                        categoria: categoria || 'Otros',
                        origen: cuentaNombreToId(cuenta),
                        nota: nota || null,
                        cuotas,
                        cuotaMensual
                    });
                }
                localStorage.setItem('gastos', JSON.stringify(gastos));
            }

            if (workbook.SheetNames.includes('Categorías')) {
                const ws = workbook.Sheets['Categorías'];
                const categorias = [];
                const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
                for (let r = 1; r <= range.e.r; r++) {
                    const nombre = val(ws, r, 0);
                    const limite = parseNum(val(ws, r, 1));
                    const color = val(ws, r, 2);
                    if (!nombre || String(nombre).toLowerCase().includes('sin categorías')) continue;
                    categorias.push({ nombre: String(nombre).trim(), limite: limite > 0 ? limite : null, color: color || '#6b7280' });
                }
                localStorage.setItem('categorias', JSON.stringify(categorias));
            }

            const mapaMetaNombreId = {};
            if (workbook.SheetNames.includes('Metas')) {
                const ws = workbook.Sheets['Metas'];
                const metas = [];
                const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
                for (let r = 1; r <= range.e.r; r++) {
                    const nombre = val(ws, r, 0);
                    const objetivo = parseNum(val(ws, r, 1));
                    const plazo = val(ws, r, 4);
                    if (!nombre || String(nombre).toLowerCase().includes('sin metas')) continue;
                    const id = 'meta_' + Date.now() + '_' + r;
                    metas.push({ id, nombre: String(nombre).trim(), objetivo, plazo: plazo || null });
                    mapaMetaNombreId[String(nombre).trim().toLowerCase()] = id;
                }
                localStorage.setItem('metas', JSON.stringify(metas));
            }

            if (workbook.SheetNames.includes('Aportes a metas') && Object.keys(mapaMetaNombreId).length > 0) {
                const ws = workbook.Sheets['Aportes a metas'];
                const contribuciones = [];
                const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
                for (let r = 1; r <= range.e.r; r++) {
                    const fechaVal = val(ws, r, 1);
                    const metaNombre = val(ws, r, 2);
                    const monto = parseNum(val(ws, r, 3));
                    const cuenta = val(ws, r, 4);
                    if (String(val(ws, r, 0)).includes('═══') || !fechaVal || !metaNombre || monto <= 0) continue;
                    const d = parseFecha(fechaVal);
                    if (!d) continue;
                    const metaId = mapaMetaNombreId[String(metaNombre).trim().toLowerCase()];
                    if (!metaId) continue;
                    contribuciones.push({ metaId, cantidad: monto, fecha: fechaToStr(d), origen: cuentaNombreToId(cuenta) });
                }
                localStorage.setItem('contribucionesMetas', JSON.stringify(contribuciones));
            }

            if (workbook.SheetNames.includes('Pagos programados')) {
                const ws = workbook.Sheets['Pagos programados'];
                const pagos = [];
                const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
                for (let r = 1; r <= range.e.r; r++) {
                    const concepto = val(ws, r, 0);
                    const monto = parseNum(val(ws, r, 1));
                    const frecuencia = val(ws, r, 2);
                    const diaPago = val(ws, r, 3);
                    const cuenta = val(ws, r, 4);
                    const categoria = val(ws, r, 5);
                    const fechaInicio = val(ws, r, 6);
                    const activo = val(ws, r, 7);
                    if (!concepto || String(concepto).toLowerCase().includes('sin pagos')) continue;
                    const id = 'pago_' + Date.now() + '_' + r;
                    let fechaInicioStr = '';
                    if (fechaInicio) {
                        const df = parseFecha(fechaInicio);
                        if (df) fechaInicioStr = df.toISOString().slice(0, 10);
                    }
                    pagos.push({
                        id,
                        concepto: String(concepto).trim(),
                        monto,
                        frecuencia: frecuencia || 'mensual',
                        diaPago: diaPago || 1,
                        cuenta: cuentaNombreToId(cuenta),
                        categoria: categoria || '',
                        fechaInicio: fechaInicioStr,
                        activo: String(activo).toLowerCase() !== 'no',
                        nota: ''
                    });
                }
                localStorage.setItem('pagosProgramados', JSON.stringify(pagos));
            }

            alert('¡Importación completada! Se han cargado todos los datos del Excel.');
            location.reload();
        } catch (err) {
            console.error(err);
            alert('Error al importar: ' + (err.message || 'Formato de archivo no reconocido. Asegúrate de usar una plantilla MoneyTrack exportada.'));
        }
    };
    reader.readAsArrayBuffer(archivo);
}

/**
 * Importa datos desde un archivo Excel (plantilla MoneyTrack).
 * Reconoce la estructura exportada y rellena automáticamente el sistema.
 */
function importarDesdeExcel(archivo) {
    if (typeof XLSX === 'undefined') {
        alert('Error: La librería Excel no está cargada. Recarga la página e intenta de nuevo.');
        return;
    }
    if (!archivo || !archivo.name) {
        alert('Selecciona un archivo Excel (.xlsx) para importar.');
        return;
    }
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            const MESES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];

            const val = (ws, r, c) => {
                const cell = ws[XLSX.utils.encode_cell({ r, c })];
                return cell ? (cell.v !== undefined ? cell.v : '') : '';
            };

            const parseNum = (v) => {
                if (v === '' || v === null || v === undefined) return 0;
                if (typeof v === 'number' && !isNaN(v)) return v;
                const s = String(v).trim().replace(/\s+[A-Z]{3}$/i, '').replace(/\./g, '').replace(',', '.');
                const n = parseFloat(s);
                return isNaN(n) ? 0 : n;
            };

            const parseFecha = (v) => {
                if (!v) return null;
                if (v instanceof Date) return v;
                if (typeof v === 'number') {
                    const d = XLSX.SSF.parse_date_code ? XLSX.SSF.parse_date_code(v) : null;
                    if (d) return new Date(d.y, d.m - 1, d.d, d.H || 0, d.M || 0);
                    return new Date((v - 25569) * 86400 * 1000);
                }
                const s = String(v).trim();
                const m = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2}))?/);
                if (m) return new Date(+m[3], +m[2] - 1, +m[1], m[4] ? +m[4] : 12, m[5] ? +m[5] : 0);
                return new Date(s);
            };

            const fechaToStr = (d) => {
                if (!d || !(d instanceof Date) || isNaN(d.getTime())) return null;
                const p = n => String(n).padStart(2, '0');
                return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}`;
            };

            const cuentaNombreToId = (nombre) => {
                const n = (nombre || '').toString().trim().toLowerCase();
                const map = { 'efectivo': 'efectivo', 'banco': 'banco', 'tarjeta de crédito': 'tarjetaCredito', 'nequi': 'nequi', 'daviplata': 'daviplata' };
                return map[n] || CUENTAS.find(c => c.nombre.toLowerCase() === n)?.id || 'efectivo';
            };

            let moneda = localStorage.getItem('moneda') || '';

            if (workbook.SheetNames.includes('Resumen')) {
                const ws = workbook.Sheets['Resumen'];
                for (let r = 0; r < 30; r++) {
                    const a = String(val(ws, r, 0)).trim();
                    const b = val(ws, r, 1);
                    if (a === 'Moneda' && b) { moneda = String(b).trim(); break; }
                }
                if (moneda) localStorage.setItem('moneda', moneda);

                for (let r = 0; r < 30; r++) {
                    const a = String(val(ws, r, 0)).trim();
                    const b = val(ws, r, 1);
                    if (a === 'Presupuesto mensual' && b && String(b).toLowerCase() !== 'no definido') {
                        const n = parseNum(b);
                        if (n > 0) localStorage.setItem('presupuestoMensual', n.toString());
                    }
                    if (a === 'Límite tarjeta de crédito' && b && String(b).toLowerCase() !== 'no definido') {
                        const n = parseNum(b);
                        if (n > 0) localStorage.setItem('limiteTarjetaCredito', n.toString());
                    }
                }

                let enSaldosIni = false;
                const saldosIni = {};
                for (let r = 0; r < 30; r++) {
                    const a = String(val(ws, r, 0)).trim();
                    const b = val(ws, r, 1);
                    if (a === 'SALDOS INICIALES POR CUENTA') { enSaldosIni = true; continue; }
                    if (enSaldosIni && a === '') break;
                    if (enSaldosIni && a && a !== 'SALDO ACTUAL POR CUENTA') {
                        const cuenta = CUENTAS.find(c => c.nombre === a);
                        if (cuenta) saldosIni[cuenta.id] = parseNum(b);
                    }
                }
                if (Object.keys(saldosIni).length > 0) localStorage.setItem('saldosCuentas', JSON.stringify(saldosIni));
            }

            if (workbook.SheetNames.includes('Categorías')) {
                const ws = workbook.Sheets['Categorías'];
                const cats = [];
                for (let r = 1; r < 200; r++) {
                    const nom = String(val(ws, r, 0)).trim();
                    if (!nom || nom.toLowerCase() === 'sin categorías creadas') break;
                    const limite = parseNum(val(ws, r, 1));
                    const color = String(val(ws, r, 2)).trim();
                    cats.push({ nombre: nom, limite: limite > 0 ? limite : null, color: color || '#6b7280' });
                }
                if (cats.length > 0) localStorage.setItem('categorias', JSON.stringify(cats));
            }

            if (workbook.SheetNames.includes('Metas')) {
                const ws = workbook.Sheets['Metas'];
                const metas = [];
                for (let r = 1; r < 200; r++) {
                    const nom = String(val(ws, r, 0)).trim();
                    if (!nom || nom.toLowerCase() === 'sin metas creadas') break;
                    const obj = parseNum(val(ws, r, 1));
                    const id = 'meta_' + Date.now() + '_' + r;
                    metas.push({ id, nombre: nom, objetivo: obj, plazo: String(val(ws, r, 4)).trim() || null });
                }
                if (metas.length > 0) localStorage.setItem('metas', JSON.stringify(metas));
            }

            const metaNombreToId = {};
            const metasData = JSON.parse(localStorage.getItem('metas') || '[]');
            metasData.forEach((m, i) => { metaNombreToId[m.nombre] = m.id; });

            if (workbook.SheetNames.includes('Ingresos')) {
                const ws = workbook.Sheets['Ingresos'];
                const ingresos = [];
                for (let r = 1; r < 2000; r++) {
                    const mesVal = String(val(ws, r, 0)).trim();
                    const fechaVal = val(ws, r, 1);
                    const concepto = String(val(ws, r, 2)).trim();
                    const monto = parseNum(val(ws, r, 3));
                    const cuenta = String(val(ws, r, 4)).trim();
                    if (!mesVal && !fechaVal && !concepto && monto === 0) continue;
                    if (concepto.includes('═══') || concepto.toLowerCase() === 'sin ingresos registrados') continue;
                    const d = parseFecha(fechaVal);
                    if (!d || isNaN(d.getTime())) continue;
                    ingresos.push({ cantidad: monto, fecha: fechaToStr(d), origen: cuentaNombreToId(cuenta), nota: concepto || 'Ingreso' });
                }
                if (ingresos.length > 0) localStorage.setItem('ingresos', JSON.stringify(ingresos));
            }

            if (workbook.SheetNames.includes('Gastos')) {
                const ws = workbook.Sheets['Gastos'];
                const gastos = [];
                for (let r = 1; r < 2000; r++) {
                    const fechaVal = val(ws, r, 1);
                    const concepto = String(val(ws, r, 2)).trim();
                    const categoria = String(val(ws, r, 3)).trim();
                    const monto = parseNum(val(ws, r, 4));
                    const cuenta = String(val(ws, r, 5)).trim();
                    const nota = String(val(ws, r, 6)).trim();
                    const cuotasStr = String(val(ws, r, 7)).trim();
                    if (!concepto && !categoria && monto === 0) continue;
                    if (concepto.includes('═══') || concepto.toLowerCase() === 'sin gastos registrados') continue;
                    const d = parseFecha(fechaVal);
                    if (!d || isNaN(d.getTime())) continue;
                    const cuotas = cuotasStr.match(/(\d+)/) ? parseInt(cuotasStr.match(/(\d+)/)[1], 10) : 1;
                    const cuotaMensual = cuotas > 1 ? monto / cuotas : monto;
                    gastos.push({
                        nombre: concepto, cantidad: monto, fecha: fechaToStr(d), categoria: categoria || 'Otros',
                        origen: cuentaNombreToId(cuenta), nota: nota || null,
                        cuotas: cuentaNombreToId(cuenta) === 'tarjetaCredito' ? cuotas : 1,
                        cuotaMensual: cuentaNombreToId(cuenta) === 'tarjetaCredito' ? cuotaMensual : monto
                    });
                }
                if (gastos.length > 0) localStorage.setItem('gastos', JSON.stringify(gastos));
            }

            if (workbook.SheetNames.includes('Aportes a metas')) {
                const ws = workbook.Sheets['Aportes a metas'];
                const contribuciones = [];
                for (let r = 1; r < 2000; r++) {
                    const fechaVal = val(ws, r, 1);
                    const metaNom = String(val(ws, r, 2)).trim();
                    const monto = parseNum(val(ws, r, 3));
                    const cuenta = String(val(ws, r, 4)).trim();
                    if (!metaNom || metaNom.includes('═══') || metaNom.toLowerCase() === 'sin aportes registrados') continue;
                    const metaId = metaNombreToId[metaNom] || metasData[0]?.id;
                    if (!metaId) continue;
                    const d = parseFecha(fechaVal);
                    if (!d || isNaN(d.getTime())) continue;
                    contribuciones.push({ metaId, cantidad: monto, fecha: fechaToStr(d), origen: cuentaNombreToId(cuenta) });
                }
                if (contribuciones.length > 0) localStorage.setItem('contribucionesMetas', JSON.stringify(contribuciones));
            }

            if (workbook.SheetNames.includes('Pagos programados')) {
                const ws = workbook.Sheets['Pagos programados'];
                const pagos = [];
                for (let r = 1; r < 200; r++) {
                    const concepto = String(val(ws, r, 0)).trim();
                    if (!concepto || concepto.toLowerCase() === 'sin pagos programados') break;
                    const monto = parseNum(val(ws, r, 1));
                    const frecuencia = String(val(ws, r, 2)).trim();
                    const diaPago = val(ws, r, 3);
                    const cuenta = String(val(ws, r, 4)).trim();
                    const categoria = String(val(ws, r, 5)).trim();
                    const fechaInicio = String(val(ws, r, 6)).trim();
                    const activo = String(val(ws, r, 7)).toLowerCase();
                    pagos.push({
                        id: 'pago_' + Date.now() + '_' + r,
                        concepto, monto, frecuencia: frecuencia || 'mensual',
                        diaPago: diaPago ? parseInt(diaPago, 10) || 1 : 1,
                        cuenta: cuentaNombreToId(cuenta), categoria: categoria || '',
                        fechaInicio: fechaInicio || new Date().toISOString().slice(0, 10),
                        activo: activo !== 'no', nota: ''
                    });
                }
                if (pagos.length > 0) localStorage.setItem('pagosProgramados', JSON.stringify(pagos));
            }

            alert('¡Importación completada! Los datos del Excel se han cargado correctamente.');
            location.reload();
        } catch (err) {
            console.error(err);
            alert('Error al importar: ' + (err.message || 'El archivo no tiene el formato esperado de MoneyTrack.'));
        }
    };
    reader.readAsArrayBuffer(archivo);
}

/**
 * Importa datos desde un archivo Excel (plantilla MoneyTrack).
 * Reconoce la estructura exportada y rellena el sistema automáticamente.
 */
function importarDesdeExcel(archivo) {
    if (typeof XLSX === 'undefined') {
        alert('Error: La librería Excel no está cargada. Recarga la página e intenta de nuevo.');
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

            const val = (ws, r, c) => {
                const cell = ws[XLSX.utils.encode_cell({ r, c })];
                return cell ? (cell.v !== undefined ? cell.v : '') : '';
            };

            const parseNum = (v) => {
                if (v === '' || v === null || v === undefined) return 0;
                if (typeof v === 'number' && !isNaN(v)) return v;
                const s = String(v).replace(/\s+/g, ' ').trim();
                const sinMoneda = s.replace(/\s*(USD|EUR|MXN|COP|ARS|CLP|PEN|GBP)\s*$/i, '').trim();
                const numStr = sinMoneda.replace(/\./g, '').replace(',', '.');
                const n = parseFloat(numStr);
                return isNaN(n) ? 0 : n;
            };

            const parseFecha = (v) => {
                if (!v) return null;
                if (v instanceof Date) return v;
                if (typeof v === 'number') {
                    const d = XLSX.SSF.parse_date_code ? XLSX.SSF.parse_date_code(v) : null;
                    if (d) return new Date(d.y, d.m - 1, d.d, d.H || 0, d.M || 0, d.S || 0);
                    return new Date((v - 25569) * 86400 * 1000);
                }
                const s = String(v).trim();
                const m = s.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})\s*(\d{1,2})?:?(\d{2})?/);
                if (m) {
                    const dia = parseInt(m[1], 10);
                    const mes = parseInt(m[2], 10) - 1;
                    const año = parseInt(m[3], 10) < 100 ? 2000 + parseInt(m[3], 10) : parseInt(m[3], 10);
                    const h = m[4] ? parseInt(m[4], 10) : 12;
                    const min = m[5] ? parseInt(m[5], 10) : 0;
                    return new Date(año, mes, dia, h, min, 0);
                }
                const d = new Date(s);
                return isNaN(d.getTime()) ? null : d;
            };

            const fechaToStr = (d) => {
                if (!d || !(d instanceof Date) || isNaN(d.getTime())) return null;
                const p = n => String(n).padStart(2, '0');
                return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}:00`;
            };

            const cuentaNombreToId = (nombre) => {
                const n = (nombre || '').toString().trim().toLowerCase();
                const map = { 'efectivo': 'efectivo', 'banco': 'banco', 'tarjeta de crédito': 'tarjetaCredito', 'tarjeta de credito': 'tarjetaCredito', 'nequi': 'nequi', 'daviplata': 'daviplata' };
                return map[n] || CUENTAS.find(c => c.nombre.toLowerCase() === n)?.id || 'efectivo';
            };

            const genId = () => 'id_' + Date.now() + '_' + Math.random().toString(36).slice(2, 9);

            let moneda = localStorage.getItem('moneda') || '';

            // ========== HOJA Resumen ==========
            const wsResumen = workbook.Sheets['Resumen'] || workbook.Sheets[workbook.SheetNames[0]];
            const saldosIni = { efectivo: 0, banco: 0, tarjetaCredito: 0, nequi: 0, daviplata: 0 };
            if (wsResumen) {
                let enSeccionSaldosIniciales = false;
                for (let r = 0; r < 30; r++) {
                    const a = String(val(wsResumen, r, 0)).trim();
                    const b = val(wsResumen, r, 1);
                    if (a === 'Moneda' && b) { moneda = String(b).trim().split(/\s/)[0] || moneda; }
                    if (a === 'Presupuesto mensual' && b) {
                        const n = parseNum(b);
                        if (n > 0) localStorage.setItem('presupuestoMensual', n.toString());
                    }
                    if (a === 'Límite tarjeta de crédito' && b) {
                        const n = parseNum(b);
                        if (n > 0) localStorage.setItem('limiteTarjetaCredito', n.toString());
                    }
                    if (a === 'SALDOS INICIALES POR CUENTA') enSeccionSaldosIniciales = true;
                    if (a === 'SALDO ACTUAL POR CUENTA' || a === 'Saldo total disponible') enSeccionSaldosIniciales = false;
                    if (enSeccionSaldosIniciales && ['Efectivo','Banco','Tarjeta de crédito','Nequi','Daviplata'].includes(a)) {
                        const id = cuentaNombreToId(a);
                        if (id) saldosIni[id] = parseNum(b);
                    }
                }
                if (moneda) localStorage.setItem('moneda', moneda);
                localStorage.setItem('saldosCuentas', JSON.stringify(saldosIni));
            }

            // ========== HOJA Categorías (primero, para que existan antes de gastos) ==========
            const wsCat = workbook.Sheets['Categorías'];
            if (wsCat) {
                const categorias = [];
                const range = XLSX.utils.decode_range(wsCat['!ref'] || 'A1');
                for (let r = 1; r <= range.e.r; r++) {
                    const nom = String(val(wsCat, r, 0)).trim();
                    if (!nom || nom.includes('Sin categorías') || nom === 'CATEGORÍA') continue;
                    const limite = parseNum(val(wsCat, r, 1));
                    const color = String(val(wsCat, r, 2)).trim() || '#6b7280';
                    categorias.push({ nombre: nom, color, limite: limite > 0 ? limite : null });
                }
                if (categorias.length > 0) localStorage.setItem('categorias', JSON.stringify(categorias));
            }

            // ========== HOJA Metas (antes de Aportes) ==========
            const wsMetas = workbook.Sheets['Metas'];
            const metaNombreToId = {};
            if (wsMetas) {
                const metas = [];
                const range = XLSX.utils.decode_range(wsMetas['!ref'] || 'A1');
                for (let r = 1; r <= range.e.r; r++) {
                    const nom = String(val(wsMetas, r, 0)).trim();
                    if (!nom || nom.includes('Sin metas') || nom === 'META') continue;
                    const id = genId();
                    metaNombreToId[nom] = id;
                    metas.push({
                        id,
                        nombre: nom,
                        objetivo: parseNum(val(wsMetas, r, 1)),
                        plazo: String(val(wsMetas, r, 4)).trim() || null
                    });
                }
                if (metas.length > 0) localStorage.setItem('metas', JSON.stringify(metas));
            }

            // ========== HOJA Ingresos ==========
            const wsIng = workbook.Sheets['Ingresos'];
            if (wsIng) {
                const ingresos = [];
                const range = XLSX.utils.decode_range(wsIng['!ref'] || 'A1');
                for (let r = 1; r <= range.e.r; r++) {
                    const fechaVal = val(wsIng, r, 1);
                    const montoVal = val(wsIng, r, 3);
                    const cuentaVal = val(wsIng, r, 4);
                    const notaVal = val(wsIng, r, 2);
                    if (!fechaVal && !montoVal) continue;
                    const d = parseFecha(fechaVal);
                    const monto = parseNum(montoVal);
                    if (!d || monto <= 0) continue;
                    const cuenta = cuentaNombreToId(cuentaVal);
                    ingresos.push({ cantidad: monto, fecha: fechaToStr(d), origen: cuenta, nota: (notaVal || 'Ingreso').toString().trim() });
                }
                if (ingresos.length > 0) localStorage.setItem('ingresos', JSON.stringify(ingresos));
            }

            // ========== HOJA Gastos ==========
            const wsGas = workbook.Sheets['Gastos'];
            if (wsGas) {
                const gastos = [];
                const range = XLSX.utils.decode_range(wsGas['!ref'] || 'A1');
                for (let r = 1; r <= range.e.r; r++) {
                    const fechaVal = val(wsGas, r, 1);
                    const concepto = String(val(wsGas, r, 2)).trim();
                    const categoria = String(val(wsGas, r, 3)).trim();
                    const montoVal = val(wsGas, r, 4);
                    const cuentaVal = val(wsGas, r, 5);
                    const notaVal = val(wsGas, r, 6);
                    const cuotasVal = String(val(wsGas, r, 7)).trim();
                    if (!fechaVal && !montoVal) continue;
                    const d = parseFecha(fechaVal);
                    const monto = parseNum(montoVal);
                    if (!d || monto <= 0) continue;
                    const cuotas = parseInt(cuotasVal, 10) || 1;
                    const cuotaMensual = cuotas > 1 ? monto / cuotas : monto;
                    gastos.push({
                        nombre: concepto || 'Gasto',
                        cantidad: monto,
                        fecha: fechaToStr(d),
                        categoria: categoria || 'Otros',
                        origen: cuentaNombreToId(cuentaVal),
                        nota: notaVal || null,
                        cuotas: cuotas,
                        cuotaMensual
                    });
                }
                if (gastos.length > 0) localStorage.setItem('gastos', JSON.stringify(gastos));
            }

            // ========== HOJA Aportes a metas ==========
            const wsAport = workbook.Sheets['Aportes a metas'];
            if (wsAport && Object.keys(metaNombreToId).length > 0) {
                const contribuciones = [];
                const range = XLSX.utils.decode_range(wsAport['!ref'] || 'A1');
                for (let r = 1; r <= range.e.r; r++) {
                    const fechaVal = val(wsAport, r, 1);
                    const metaNom = String(val(wsAport, r, 2)).trim();
                    const montoVal = val(wsAport, r, 3);
                    const cuentaVal = val(wsAport, r, 4);
                    if (!fechaVal || !metaNom || !montoVal) continue;
                    const metaId = metaNombreToId[metaNom];
                    if (!metaId) continue;
                    const d = parseFecha(fechaVal);
                    const monto = parseNum(montoVal);
                    if (!d || monto <= 0) continue;
                    contribuciones.push({ metaId, cantidad: monto, fecha: fechaToStr(d), origen: cuentaNombreToId(cuentaVal) });
                }
                if (contribuciones.length > 0) localStorage.setItem('contribucionesMetas', JSON.stringify(contribuciones));
            }

            // ========== HOJA Pagos programados ==========
            const wsPagos = workbook.Sheets['Pagos programados'];
            if (wsPagos) {
                const pagos = [];
                const range = XLSX.utils.decode_range(wsPagos['!ref'] || 'A1');
                for (let r = 1; r <= range.e.r; r++) {
                    const concepto = String(val(wsPagos, r, 0)).trim();
                    const montoVal = val(wsPagos, r, 1);
                    if (!concepto || concepto.includes('Sin pagos')) continue;
                    const monto = parseNum(montoVal);
                    if (monto <= 0) continue;
                    const activoStr = String(val(wsPagos, r, 7)).toLowerCase();
                    pagos.push({
                        id: genId(),
                        concepto,
                        monto,
                        frecuencia: String(val(wsPagos, r, 2)).trim() || 'mensual',
                        diaPago: parseInt(val(wsPagos, r, 3), 10) || 1,
                        cuenta: cuentaNombreToId(val(wsPagos, r, 4)),
                        categoria: String(val(wsPagos, r, 5)).trim() || '',
                        fechaInicio: val(wsPagos, r, 6) ? parseFecha(val(wsPagos, r, 6)) : null,
                        activo: activoStr !== 'no' && activoStr !== 'false',
                        nota: ''
                    });
                }
                pagos.forEach(p => {
                    if (p.fechaInicio && p.fechaInicio instanceof Date) {
                        const d = p.fechaInicio;
                        p.fechaInicio = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
                    }
                });
                if (pagos.length > 0) localStorage.setItem('pagosProgramados', JSON.stringify(pagos));
            }

            alert('¡Importación completada! Los datos del Excel se han cargado correctamente.');
            location.reload();
        } catch (err) {
            console.error(err);
            alert('Error al importar: ' + (err.message || 'Formato de archivo no reconocido. Asegúrate de usar una plantilla exportada desde MoneyTrack.'));
        }
    };
    reader.readAsArrayBuffer(archivo);
}
