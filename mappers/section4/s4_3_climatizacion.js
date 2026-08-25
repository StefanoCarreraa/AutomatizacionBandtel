// const { aplicarEstiloCelda } = require('../../utils/excelTools');

// function s4_3(hojaDestino, fPruebas, filaInicio) {
//     if (!fPruebas) return filaInicio;

//     // +3 filas después de la sección anterior
//     let filaActual = filaInicio + 3;
//     let bloquesEscritos = 0;

//     // Definición explícita de los últimos dos bloques
//     const bloquesMatriz4_3 = [
//         { start: 'FK' }, 
//         { start: 'FQ' }
//     ];

//     const limpiarMergesFilaExacta = (nFila) => {
//         if (!hojaDestino._merges) return;
//         Object.keys(hojaDestino._merges).forEach(key => {
//             const m = hojaDestino._merges[key];
//             if (m.top === nFila || m.bottom === nFila) delete hojaDestino._merges[key];
//         });
//     };

//     const escribirFilaS4_3 = (numFila, valores, insertar) => {
//         if (insertar) hojaDestino.insertRow(numFila, {});

//         limpiarMergesFilaExacta(numFila);
//         const row = hojaDestino.getRow(numFila);
//         row.height = 25;

//         // Limpieza y estilo base
//         for (let c = 2; c <= 13; c++) {
//             const cell = row.getCell(c);
//             cell.value = null;
//             cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
//             cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
//             cell.font = { name: 'Arial', size: 11 };
//         }

//         // --- MAPEO DE COLUMNAS (Pares combinados B-M) ---
//         row.getCell('B').value = valores[0]; 
//         row.getCell('D').value = valores[1]; 
//         row.getCell('F').value = valores[2]; 
//         row.getCell('H').value = valores[3]; 
//         row.getCell('J').value = valores[4]; 
//         row.getCell('L').value = valores[5]; 

//         try {
//             hojaDestino.mergeCells(numFila, 2, numFila, 3);   // B-C
//             hojaDestino.mergeCells(numFila, 4, numFila, 5);   // D-E
//             hojaDestino.mergeCells(numFila, 6, numFila, 7);   // F-G
//             hojaDestino.mergeCells(numFila, 8, numFila, 9);   // H-I
//             hojaDestino.mergeCells(numFila, 10, numFila, 11); // J-K
//             hojaDestino.mergeCells(numFila, 12, numFila, 13); // L-M
//         } catch (e) {}
//     };

//     bloquesMatriz4_3.forEach((bloque) => {
//         const celdaBase = fPruebas.getCell(bloque.start);
//         const valorRef = celdaBase.value;

//         // Si el bloque es NA o está vacío, no lo imprimimos
//         if (!valorRef || valorRef.toString().trim().toUpperCase() === '') return;

//         const colIdx = celdaBase.col;
//         const datos = [];
//         for (let i = 0; i < 6; i++) {
//             datos.push(fPruebas.getCell(colIdx + i).value);
//         }

//         escribirFilaS4_3(filaActual, datos, bloquesEscritos > 0);

//         filaActual++;
//         bloquesEscritos++;
//     });

//     return filaActual;
// }

// module.exports = s4_3;

const { aplicarEstiloCelda } = require('../../utils/excelTools');

function s4_3(hojaDestino, fPruebas, filaInicio) {
    if (!fPruebas) return filaInicio;

    // --- PROTECCIÓN ANTI-NaN ---
    // Localización dinámica mediante el texto del encabezado
    let filaSeccion = -1;
    hojaDestino.eachRow((row, rowNumber) => {
        const valorB = row.getCell('B').value;
        if (valorB && valorB.toString().includes("4.3. Sistema de Climatización")) {
            filaSeccion = rowNumber;
        }
    });

    // Inicio 2 filas después del encabezado o respaldo en filaInicio + 3
    let filaActual = (filaSeccion !== -1) ? (filaSeccion + 2) : (filaInicio + 3);

    // Validación final para evitar errores de puntero Address: BNaN
    if (isNaN(filaActual)) filaActual = 140;

    let bloquesEscritos = 0;

    // Definición de bloques: FK y FQ (Cada bloque tiene 6 parámetros en la matriz)
    const bloquesMatriz4_3 = [
        { start: 'FK' },
        { start: 'FQ' }
    ];

    const limpiarMergesFilaExacta = (nFila) => {
        if (!hojaDestino._merges) return;
        Object.keys(hojaDestino._merges).forEach(key => {
            const m = hojaDestino._merges[key];
            if (m.top === nFila || m.bottom === nFila) delete hojaDestino._merges[key];
        });
    };

    const escribirFilaS4_3 = (numFila, valores, insertar) => {
        if (insertar) {
            hojaDestino.insertRow(numFila, {});
        }

        limpiarMergesFilaExacta(numFila);
        const row = hojaDestino.getRow(numFila);
        row.height = 25;

        // Reset y estilo base de B a M
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
        row.getCell('B').value = valores[0]; // Equipo (FK, FQ) -> B-C
        row.getCell('D').value = valores[1]; // Marca (FL, FR) -> D-E
        row.getCell('F').value = valores[2]; // Modelo (FM, FS) -> F-G
        row.getCell('H').value = valores[3]; // Serie (FN, FT) -> H-I
        row.getCell('J').value = valores[4]; // Capacidad (FO, FU) -> J-K
        row.getCell('L').value = valores[5]; // Temp. Seteo °C (FP, FV) -> L-M

        try {
            hojaDestino.mergeCells(numFila, 2, numFila, 3);   // Equipos
            hojaDestino.mergeCells(numFila, 4, numFila, 5);   // Marca
            hojaDestino.mergeCells(numFila, 6, numFila, 7);   // Modelo
            hojaDestino.mergeCells(numFila, 8, numFila, 9);   // Serie
            hojaDestino.mergeCells(numFila, 10, numFila, 11); // Capacidad
            hojaDestino.mergeCells(numFila, 12, numFila, 13); // Temperatura
        } catch (e) { }
    };

    bloquesMatriz4_3.forEach((bloque) => {
        const celdaEquipo = fPruebas.getCell(bloque.start);
        const valorEquipo = celdaEquipo.value ? celdaEquipo.value.toString().trim().toUpperCase() : "";

        // --- FILTRO ESTRICTO ---
        // Ignora si está vacío, es "NA", o es un guion "-"
        const esInvalido = !valorEquipo ||
            valorEquipo === 'NA' ||
            valorEquipo === '-';

        if (esInvalido) return;

        const colIdx = celdaEquipo.col;
        const datos = [];
        for (let i = 0; i < 6; i++) {
            // Usamos .result para obtener el valor puro, evitando la conversión automática a fecha de ExcelJS
            const celda = fPruebas.getCell(colIdx + i);
            const valorCrudo = celda.result !== undefined ? celda.result : celda.value;
            datos.push(valorCrudo);
        }

        escribirFilaS4_3(filaActual, datos, bloquesEscritos > 0);

        filaActual++;
        bloquesEscritos++;
    });
    
    return filaActual;
}

module.exports = s4_3;