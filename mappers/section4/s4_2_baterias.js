// const { aplicarEstiloCelda } = require('../../utils/excelTools');

// function s4_2(hojaDestino, fPruebas, filaInicio) {
//     if (!fPruebas) return filaInicio;

//     // +3 filas después de la sección anterior
//     let filaActual = filaInicio + 3;
//     let bloquesEscritos = 0;

//     /**
//      * Generamos los inicios de bloque de 6 en 6 
//      * desde CE (col 83) hasta FE (col 161)
//      */
//     const bloquesMatriz4_2 = [];
//     const colInicio = fPruebas.getCell('CE').col;
//     const colFin = fPruebas.getCell('FE').col;

//     for (let c = colInicio; c <= colFin; c += 6) {
//         bloquesMatriz4_2.push({ colIdx: c });
//     }

//     const limpiarMergesFilaExacta = (nFila) => {
//         if (!hojaDestino._merges) return;
//         Object.keys(hojaDestino._merges).forEach(key => {
//             const m = hojaDestino._merges[key];
//             if (m.top === nFila || m.bottom === nFila) delete hojaDestino._merges[key];
//         });
//     };

//     const escribirFilaS4_2 = (numFila, valores, insertar) => {
//         if (insertar) hojaDestino.insertRow(numFila, {});
        
//         limpiarMergesFilaExacta(numFila);
//         const row = hojaDestino.getRow(numFila);
//         row.height = 25;

//         // Estilo y limpieza de B a M
//         for (let c = 2; c <= 13; c++) {
//             const cell = row.getCell(c);
//             cell.value = null;
//             cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
//             cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
//             cell.font = { name: 'Arial', size: 11 };
//         }

//         // --- MAPEO DE COLUMNAS (Pares combinados) ---
//         row.getCell('B').value = valores[0]; // CE, CK, CQ... (B-C)
//         row.getCell('D').value = valores[1]; // CF, CL, CR... (D-E)
//         row.getCell('F').value = valores[2]; // CG, CM, CS... (F-G)
//         row.getCell('H').value = valores[3]; // CH, CN, CT... (H-I)
//         row.getCell('J').value = valores[4]; // CI, CO, CU... (J-K)
//         row.getCell('L').value = valores[5]; // CJ, CP, CV... (L-M)

//         try {
//             hojaDestino.mergeCells(numFila, 2, numFila, 3);   // B-C
//             hojaDestino.mergeCells(numFila, 4, numFila, 5);   // D-E
//             hojaDestino.mergeCells(numFila, 6, numFila, 7);   // F-G
//             hojaDestino.mergeCells(numFila, 8, numFila, 9);   // H-I
//             hojaDestino.mergeCells(numFila, 10, numFila, 11); // J-K
//             hojaDestino.mergeCells(numFila, 12, numFila, 13); // L-M
//         } catch (e) {}
//     };

//     bloquesMatriz4_2.forEach((bloque) => {
//         const valorRef = fPruebas.getCell(bloque.colIdx).value;

//         // Validación de contenido para evitar filas vacías
//         if (!valorRef || valorRef.toString().trim().toUpperCase() === '') return;

//         const datos = [];
//         for (let i = 0; i < 6; i++) {
//             datos.push(fPruebas.getCell(bloque.colIdx + i).value);
//         }

//         escribirFilaS4_2(filaActual, datos, bloquesEscritos > 0);
        
//         filaActual++;
//         bloquesEscritos++;
//     });

//     return filaActual;
// }

// module.exports = s4_2;

const { aplicarEstiloCelda } = require('../../utils/excelTools');

function s4_2(hojaDestino, fPruebas, filaInicio) {
    if (!fPruebas) return filaInicio;

    // --- PROTECCIÓN ANTI-NaN ---
    // Localización dinámica de la sección 4.2
    let filaSeccion = -1;
    hojaDestino.eachRow((row, rowNumber) => {
        const valorB = row.getCell('B').value;
        if (valorB && valorB.toString().includes("4.2. Banco de Baterias")) {
            filaSeccion = rowNumber;
        }
    });

    // Inicio 2 filas después del encabezado o respaldo basado en filaInicio + 3
    let filaActual = (filaSeccion !== -1) ? (filaSeccion + 2) : (filaInicio + 3);
    
    // Validación final para evitar Address: BNaN
    if (isNaN(filaActual)) filaActual = 120; 

    let bloquesEscritos = 0;

    /**
     * Generación de bloques desde CE (col 83) hasta FE (col 161)
     * Parámetros: Banco N°, Marca, Modelo, Capacidad (AH), Tensión (VDC), Rango Temp.
     */
    const bloquesMatriz4_2 = [];
    const colInicio = fPruebas.getCell('CE').col;
    const colFin = fPruebas.getCell('FE').col;

    for (let c = colInicio; c <= colFin; c += 6) {
        bloquesMatriz4_2.push({ colIdx: c });
    }

    const limpiarMergesFilaExacta = (nFila) => {
        if (!hojaDestino._merges) return;
        Object.keys(hojaDestino._merges).forEach(key => {
            const m = hojaDestino._merges[key];
            if (m.top === nFila || m.bottom === nFila) delete hojaDestino._merges[key];
        });
    };

    const escribirFilaS4_2 = (numFila, valores, insertar) => {
        if (insertar) {
            hojaDestino.insertRow(numFila, {});
        }
        
        limpiarMergesFilaExacta(numFila);
        const row = hojaDestino.getRow(numFila);
        row.height = 30; // Altura para lectura clara

        // Reset de B a M con formato de pruebas
        for (let c = 2; c <= 13; c++) {
            const cell = row.getCell(c);
            cell.value = null;
            cell.border = { 
                top: {style:'thin'}, left: {style:'thin'}, 
                bottom: {style:'thin'}, right: {style:'thin'} 
            };
            cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
            cell.font = { name: 'Arial', size: 11 };
        }

        // --- MAPEO DE COLUMNAS POR PARES ---
        row.getCell('B').value = valores[0]; // Banco N° (CE, CK...) -> B-C
        row.getCell('D').value = valores[1]; // Marca (CF, CL...) -> D-E
        row.getCell('F').value = valores[2]; // Modelo (CG, CM...) -> F-G
        row.getCell('H').value = valores[3]; // Capacidad AH (CH, CN...) -> H-I
        row.getCell('J').value = valores[4]; // Tensión VDC (CI, CO...) -> J-K
        row.getCell('L').value = valores[5]; // Rango Temp (CJ, CP...) -> L-M

        try {
            hojaDestino.mergeCells(numFila, 2, numFila, 3);   // Banco N°
            hojaDestino.mergeCells(numFila, 4, numFila, 5);   // Marca
            hojaDestino.mergeCells(numFila, 6, numFila, 7);   // Modelo
            hojaDestino.mergeCells(numFila, 8, numFila, 9);   // Capacidad
            hojaDestino.mergeCells(numFila, 10, numFila, 11); // Tensión
            hojaDestino.mergeCells(numFila, 12, numFila, 13); // Rango Temp
        } catch (e) {}
    };

    bloquesMatriz4_2.forEach((bloque) => {
        const celdaBanco = fPruebas.getCell(bloque.colIdx);
        const valorBanco = celdaBanco.value ? celdaBanco.value.toString().trim().toUpperCase() : "";

        // --- FILTRO ESTRICTO ---
        // Ignora si está vacío, es "NA", o es un guion "-"
        const esInvalido = !valorBanco || 
                           valorBanco === 'NA' || 
                           valorBanco === '-';

        if (esInvalido) return;

        const datos = [];
        for (let i = 0; i < 6; i++) {
            // Usamos .result para obtener el valor limpio y evitar que ExcelJS lo convierta en fecha
            const celda = fPruebas.getCell(bloque.colIdx + i);
            const valorCrudo = celda.result !== undefined ? celda.result : celda.value;
            datos.push(valorCrudo);
        }

        escribirFilaS4_2(filaActual, datos, bloquesEscritos > 0);
        
        filaActual++;
        bloquesEscritos++;
    });

    return filaActual;
}

module.exports = s4_2;