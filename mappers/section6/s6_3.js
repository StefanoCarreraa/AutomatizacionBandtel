// const { aplicarEstiloCelda } = require('../../utils/excelTools');

// function s6_3(hojaDestino, fPestana6, filaInicio) {
//     if (!fPestana6) return filaInicio;

//     let filaActual = filaInicio + 4;

//     const bloquesMatriz6_3 = [
//         { cols: ['Q', 'S', 'T'] },
//         { cols: ['U', 'W', 'X'] },
//         { cols: ['Y', 'AA', 'AB'] },
//         { cols: ['AC', 'AE', 'AF'] },
//         { cols: ['AG', 'AI', 'AJ'] },
//         { cols: ['AK', 'AM', 'AN'] }
//     ];

//     const formatearCasilla = (val) => {
//         const v = val ? val.toString().trim().toUpperCase() : "";
//         if (v === 'SI' || v === 'S' || v === 'X') return "SI (  X  )  NO (     )";
//         return "SI (     )  NO (  X  )";
//     };

//     bloquesMatriz6_3.forEach((bloque) => {
//         const row = hojaDestino.getRow(filaActual);
//         row.height = 25;

//         // Extraer valores de la matriz
//         const val1 = fPestana6.getCell(bloque.cols[0]).value; 
//         const val2 = fPestana6.getCell(bloque.cols[1]).value; 
//         const val3 = fPestana6.getCell(bloque.cols[2]).value; 

//         // --- MAPEO AL ACTA ---
//         row.getCell('E').value = formatearCasilla(val1);
//         row.getCell('J').value = val2;
//         row.getCell('K').value = val3;

//         // --- APLICACIÓN DE ESTILOS (Refuerzo de Negrita) ---
//         // Definimos todas las columnas que llevan borde y formato base
//         const todasLasColumnas = [2, 5, 7, 10, 11];

//         todasLasColumnas.forEach(colIdx => {
//             const cell = row.getCell(colIdx);
            
//             // Aplicar bordes y alineación
//             cell.border = { 
//                 top: {style:'thin'}, left: {style:'thin'}, 
//                 bottom: {style:'thin'}, right: {style:'thin'} 
//             };
//             cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };

//             // Forzar Negrita solo en las columnas específicas
//             cell.font = { 
//                 name: 'Arial', 
//                 size: 11, 
//             };
//         });

//         // Manejo de Merges (para evitar el error de "already merged")
//         try {
//             // Verificamos si la celda ya es parte de un merge antes de intentar unir
//             if (!row.getCell('E').isMerged) hojaDestino.mergeCells(filaActual, 5, filaActual, 6);
//             if (!row.getCell('K').isMerged) hojaDestino.mergeCells(filaActual, 11, filaActual, 13);
//         } catch (e) {
//             // Silenciamos errores de merge en secciones estáticas
//         }

//         filaActual++;
//     });

//     return filaActual;
// }

// module.exports = s6_3;

const { aplicarEstiloCelda } = require('../../utils/excelTools');

function s6_3(hojaDestino, fPestana6, filaInicio) {
    if (!fPestana6) return filaInicio;

    // --- PROTECCIÓN ANTI-NaN ---
    let filaSeccion = -1;
    hojaDestino.eachRow((row, rowNumber) => {
        const valorB = row.getCell('B').value;
        if (valorB && valorB.toString().includes("6.3. Patio")) {
            filaSeccion = rowNumber;
        }
    });

    let filaActual = (filaSeccion !== -1) ? (filaSeccion + 2) : (filaInicio + 4);
    if (isNaN(filaActual)) filaActual = 220; 

    const bloquesMatriz6_3 = [
        { cols: ['Q', 'S', 'T'] }, // Losa Shelter
        { cols: ['U', 'W', 'X'] }, // Losa de ingreso
        { cols: ['Y', 'AA', 'AB'] }, // Pedestales
        { cols: ['AC', 'AE', 'AF'] }, // Piso de ripio
        { cols: ['AG', 'AI', 'AJ'] }, // Luminaria
        { cols: ['AK', 'AM', 'AN'] }  // Tomacorrientes
    ];

    const formatearCasilla = (val) => {
        const v = val ? val.toString().trim().toUpperCase() : "";
        if (['SI', 'S', 'X', '1', 'VERDADERO'].includes(v)) {
            return "SI (  X  )  NO (     )";
        }
        return "SI (     )  NO (  X  )";
    };

    bloquesMatriz6_3.forEach((bloque) => {
        const row = hojaDestino.getRow(filaActual);
        row.height = 30;

        // 1. RESET DE ESTILOS (Para limpiar formatos heredados de la plantilla)
        for (let c = 2; c <= 13; c++) {
            const cell = row.getCell(c);
            cell.style = {}; // <--- RESET TOTAL
            cell.border = { 
                top: {style:'thin'}, left: {style:'thin'}, 
                bottom: {style:'thin'}, right: {style:'thin'} 
            };
            cell.font = { name: 'Arial', size: 11 };
        }

        // 2. EXTRAER VALORES Y ASIGNAR
        const val1 = fPestana6.getCell(bloque.cols[0]).value; 
        const val2 = fPestana6.getCell(bloque.cols[1]).value; 
        const val3 = fPestana6.getCell(bloque.cols[2]).value; 

        row.getCell('E').value = formatearCasilla(val1);
        row.getCell('J').value = val2;
        row.getCell('K').value = val3;

        // 3. MERGES (Diseño visual del acta)
        try {
            hojaDestino.mergeCells(filaActual, 5, filaActual, 6);   // E-F: SI/NO
            hojaDestino.mergeCells(filaActual, 7, filaActual, 9);   // G-I: Unidad Medición
            hojaDestino.mergeCells(filaActual, 11, filaActual, 13); // K-M: Comentarios
        } catch (e) {}

        // 4. ALINEACIÓN FORZADA (Post-Merge y Post-Valores)
        const estiloCentrado = { 
            vertical: 'middle', 
            horizontal: 'center', 
            wrapText: true 
        };

        for (let c = 2; c <= 13; c++) {
            row.getCell(c).alignment = estiloCentrado;
        }

        filaActual++;
    });

    return filaActual;
}

module.exports = s6_3;