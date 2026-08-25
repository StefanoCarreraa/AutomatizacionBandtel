const { aplicarEstiloCelda } = require('../../utils/excelTools');

function s4_1(hojaDestino, fPruebas, filaInicio) {
    if (!fPruebas) return filaInicio;

    // --- PROTECCIÓN ANTI-NaN ---
    // Localización dinámica de la sección 4.1
    let filaSeccion = -1;
    hojaDestino.eachRow((row, rowNumber) => {
        const valorB = row.getCell('B').value;
        if (valorB && valorB.toString().includes("4.1. Sistema Rectificador")) {
            filaSeccion = rowNumber;
        }
    });

    // Inicio 2 filas después del encabezado o respaldo basado en filaInicio + 8
    let filaActual = (filaSeccion !== -1) ? (filaSeccion + 2) : (filaInicio + 8);

    // Validación de seguridad para evitar Address: BNaN
    if (isNaN(filaActual)) filaActual = 100;

    let bloquesEscritos = 0;

    /**
     * Generación de bloques desde E (col 5) hasta BY (col 77)
     * Cada equipo tiene 6 parámetros: Equipo, Marca, Modelo, Serie, Capacidad, Rango Temp.
     */
    const bloquesMatriz4_1 = [];
    const colInicio = fPruebas.getCell('E').col;
    const colFin = fPruebas.getCell('BY').col;

    for (let c = colInicio; c <= colFin; c += 6) {
        bloquesMatriz4_1.push({ colIdx: c });
    }

    const limpiarMergesFilaExacta = (nFila) => {
        if (!hojaDestino._merges) return;
        Object.keys(hojaDestino._merges).forEach(key => {
            const m = hojaDestino._merges[key];
            if (m.top === nFila || m.bottom === nFila) delete hojaDestino._merges[key];
        });
    };

    const escribirFilaS4_1 = (numFila, valores, insertar) => {
        if (insertar) {
            hojaDestino.insertRow(numFila, {});
        }

        limpiarMergesFilaExacta(numFila);
        const row = hojaDestino.getRow(numFila);
        row.height = 30;

        // Reset de B a M con formato de pruebas
        for (let c = 2; c <= 13; c++) {
            const cell = row.getCell(c);
            cell.value = null;
            cell.border = {
                top: { style: 'thin' }, left: { style: 'thin' },
                bottom: { style: 'thin' }, right: { style: 'thin' }
            };
            cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
            cell.font = { name: 'Arial', size: 11 };
        }

        // --- MAPEO DE COLUMNAS POR PARES ---
        row.getCell('B').value = valores[0]; // Equipo (E, K...) -> B-C
        row.getCell('D').value = valores[1]; // Marca (F, L...) -> D-E
        row.getCell('F').value = valores[2]; // Modelo (G, M...) -> F-G
        row.getCell('H').value = valores[3]; // Serie (H, N...) -> H-I
        row.getCell('J').value = valores[4]; // Capacidad KW (I, O...) -> J-K
        row.getCell('L').value = valores[5]; // Rango Temp (J, P...) -> L-M

        try {
            hojaDestino.mergeCells(numFila, 2, numFila, 3);   // Equipos
            hojaDestino.mergeCells(numFila, 4, numFila, 5);   // Marca
            hojaDestino.mergeCells(numFila, 6, numFila, 7);   // Modelo
            hojaDestino.mergeCells(numFila, 8, numFila, 9);   // Serie
            hojaDestino.mergeCells(numFila, 10, numFila, 11); // Capacidad
            hojaDestino.mergeCells(numFila, 12, numFila, 13); // Rango Temp
        } catch (e) { }
    };

    bloquesMatriz4_1.forEach((bloque) => {
        const celdaEquipo = fPruebas.getCell(bloque.colIdx);
        const valorEquipo = celdaEquipo.value ? celdaEquipo.value.toString().trim().toUpperCase() : "";

        // --- FILTRO ESTRICTO ---
        // Ignora si está vacío, es "NA", o es un guion "-"
        const esInvalido = !valorEquipo ||
            valorEquipo === 'NA' ||
            valorEquipo === '-';

        if (esInvalido) return;

        const datos = [];
        for (let i = 0; i < 6; i++) {
            // Usamos .result para obtener el valor limpio antes de la interpretación de fecha de ExcelJS
            const celda = fPruebas.getCell(bloque.colIdx + i);
            datos.push(celda.result !== undefined ? celda.result : celda.value);
        }

        escribirFilaS4_1(filaActual, datos, bloquesEscritos > 0);

        filaActual++;
        bloquesEscritos++;
    });

    return filaActual;
}

module.exports = s4_1;