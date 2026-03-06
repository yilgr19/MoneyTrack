/**
 * Formatea un número con separadores de miles y millones (ej: 940000.50 → "940.000,50")
 * Usa locale español para máxima legibilidad.
 */
function formatearNumero(num, decimales = 2) {
    if (num === null || num === undefined || isNaN(num)) return '0,00';
    const n = parseFloat(num);
    return n.toLocaleString('es', {
        minimumFractionDigits: decimales,
        maximumFractionDigits: decimales
    });
}
