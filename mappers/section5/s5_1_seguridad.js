// const { aplicarEstiloCelda } = require('../../utils/excelTools');

// function s5_1(hojaDestino, fPestana5, filaInicio) {
//     if (!fPestana5) return filaInicio;

//     // +4 filas después de la sección anterior
//     let filaActual = filaInicio + 4;
//     let bloquesEscritos = 0;

//     /**
//      * Generamos los inicios de bloque de 5 en 5 
//      * desde L (col 12) hasta EG (col 137)
//      */
//     const bloquesMatriz5_1 = [];
//     const colInicio = fPestana5.getCell('L').col;
//     const colFin = fPestana5.getCell('EG').col;

//     for (let c = colInicio; c <= colFin; c += 5) {
//         bloquesMatriz5_1.push({ colIdx: c });
//     }

//     const limpiarMergesFilaExacta = (nFila) => {
//         if (!hojaDestino._merges) return;
//         Object.keys(hojaDestino._merges).forEach(key => {
//             const m = hojaDestino._merges[key];
//             if (m.top === nFila || m.bottom === nFila) delete hojaDestino._merges[key];
//         });
//     };

//     const escribirFilaS5_1 = (numFila, valores, insertar) => {
//         if (insertar) hojaDestino.insertRow(numFila, {});

//         limpiarMergesFilaExacta(numFila);
//         const row = hojaDestino.getRow(numFila);
//         row.height = 25;

//         // Reset y Estilo base
//         for (let c = 2; c <= 13; c++) {
//             const cell = row.getCell(c);
//             cell.value = null;
//             cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
//             cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
//             cell.font = { name: 'Arial', size: 11 };
//         }

//         // --- MAPEO DE COLUMNAS (Formato mixto: 3-2-2-2-3) ---
//         row.getCell('B').value = valores[0]; // (B,C,D)
//         row.getCell('E').value = valores[1]; // (E,F)
//         row.getCell('G').value = valores[2]; // (G,H)
//         row.getCell('I').value = valores[3]; // (I,J)
//         row.getCell('K').value = valores[4]; // (K,L,M)

//         try {
//             hojaDestino.mergeCells(numFila, 2, numFila, 4);   // B-D
//             hojaDestino.mergeCells(numFila, 5, numFila, 6);   // E-F
//             hojaDestino.mergeCells(numFila, 7, numFila, 8);   // G-H
//             hojaDestino.mergeCells(numFila, 9, numFila, 10);  // I-J
//             hojaDestino.mergeCells(numFila, 11, numFila, 13); // K-M
//         } catch (e) {}
//     };

//     bloquesMatriz5_1.forEach((bloque) => {
//         const valorRef = fPestana5.getCell(bloque.colIdx).value;

//         // Validación extendida: Saltamos si es NA, -, -- o vacío
//         if (!valorRef) return;
//         const valStr = valorRef.toString().trim();
//         if (['NA', '-', '--'].includes(valStr.toUpperCase())) return;

//         const datos = [];
//         for (let i = 0; i < 5; i++) {
//             datos.push(fPestana5.getCell(bloque.colIdx + i).value);
//         }

//         escribirFilaS5_1(filaActual, datos, bloquesEscritos > 0);

//         filaActual++;
//         bloquesEscritos++;
//     });

//     return filaActual;
// }

// module.exports = s5_1;

const { aplicarEstiloCelda } = require('../../utils/excelTools');

function s5_1(hojaDestino, fPestana5, filaInicio) {
    if (!fPestana5) return filaInicio;

    // --- PROTECCIÓN ANTI-NaN ---
    // Localización dinámica de la sección 5.1
    let filaSeccion = -1;
    hojaDestino.eachRow((row, rowNumber) => {
        const valorB = row.getCell('B').value;
        if (valorB && valorB.toString().includes("5.1. Sistema de Seguridad")) {
            filaSeccion = rowNumber;
        }
    });

    // Inicio 2 filas después del encabezado o respaldo basado en filaInicio + 4
    let filaActual = (filaSeccion !== -1) ? (filaSeccion + 2) : (filaInicio + 4);

    // Validación de seguridad para evitar Address: BNaN
    if (isNaN(filaActual)) filaActual = 160;

    let bloquesEscritos = 0;

    /**
     * Generación de bloques desde L (col 12) hasta EG (col 137)
     * Parámetros: Equipos, Marca, Modelo, Serie, Comentarios.
     */
    const bloquesMatriz5_1 = [];
    const colInicio = fPestana5.getCell('L').col;
    const colFin = fPestana5.getCell('EG').col;

    for (let c = colInicio; c <= colFin; c += 5) {
        bloquesMatriz5_1.push({ colIdx: c });
    }

    const limpiarMergesFilaExacta = (nFila) => {
        if (!hojaDestino._merges) return;
        Object.keys(hojaDestino._merges).forEach(key => {
            const m = hojaDestino._merges[key];
            if (m.top === nFila || m.bottom === nFila) delete hojaDestino._merges[key];
        });
    };

    const escribirFilaS5_1 = (numFila, valores, insertar) => {
        if (insertar) {
            hojaDestino.insertRow(numFila, {});
        }

        limpiarMergesFilaExacta(numFila);
        const row = hojaDestino.getRow(numFila);
        row.height = 25;

        // Reset y Estilo base de B a M
        for (let c = 2; c <= 13; c++) {
            const cell = row.getCell(c);
            cell.value = null;
            cell.border = {
                top: { style: 'thin' }, left: { style: 'thin' },
                bottom: { style: 'thin' }, right: { style: 'thin' }
            };
            cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
            cell.font = { name: 'Arial', size: 10 };
        }

        // --- MAPEO DE COLUMNAS SEGÚN FORMATO MIXTO ---
        row.getCell('B').value = valores[0]; // Equipos (L, Q...) -> B-D
        row.getCell('E').value = valores[1]; // Marca (M, R...)   -> E-F
        row.getCell('G').value = valores[2]; // Modelo (N, S...)  -> G-H
        row.getCell('I').value = valores[3]; // Serie (O, T...)   -> I-J
        row.getCell('K').value = valores[4]; // Comentarios (P, U...) -> K-M

        try {
            hojaDestino.mergeCells(numFila, 2, numFila, 4);   // B-D (Equipos)
            hojaDestino.mergeCells(numFila, 5, numFila, 6);   // E-F (Marca)
            hojaDestino.mergeCells(numFila, 7, numFila, 8);   // G-H (Modelo)
            hojaDestino.mergeCells(numFila, 9, numFila, 10);  // I-J (Serie)
            hojaDestino.mergeCells(numFila, 11, numFila, 13); // K-M (Comentarios)
        } catch (e) { }
    };

    bloquesMatriz5_1.forEach((bloque) => {
        const valorRef = fPestana5.getCell(bloque.colIdx).value;

        // Validación extendida para filtrar equipos no instalados
        if (!valorRef) return;
        const valStr = valorRef.toString().trim().toUpperCase();
        if (['NA', '-', '--', 'N/A'].includes(valStr)) return;

        const datos = [];
        for (let i = 0; i < 5; i++) {
            datos.push(fPestana5.getCell(bloque.colIdx + i).value);
        }

        escribirFilaS5_1(filaActual, datos, bloquesEscritos > 0);

        filaActual++;
        bloquesEscritos++;
    });

    return filaActual;
}

module.exports = s5_1;