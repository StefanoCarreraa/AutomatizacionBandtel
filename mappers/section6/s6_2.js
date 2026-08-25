// const { aplicarEstiloCelda } = require('../../utils/excelTools');

// function s6_2(hojaDestino, fPestana6, filaInicio) {
//     if (!fPestana6) return filaInicio;

//     // +3 filas después de la sección anterior (Estático)
//     let filaActual = filaInicio + 4;

//     const colInicio = fPestana6.getCell('K').col; // Rango K:P
//     const datos = [];
//     for (let i = 0; i < 6; i++) {
//         datos.push(fPestana6.getCell(colInicio + i).value);
//     }

//     const formatearCasilla = (val) => {
//         const v = val ? val.toString().trim().toUpperCase() : "";
//         if (v === 'SI' || v === 'S' || v === 'X') return "SI (  X  )  NO (     )";
//         return "SI (     )  NO (  X  )";
//     };

//     const row = hojaDestino.getRow(filaActual);
//     row.height = 25;

//     // Mapeo: B:C es K. De L a P (datos[1] a datos[5]) son SI/NO
//     row.getCell('B').value = datos[0];                  // K de matriz
//     row.getCell('D').value = formatearCasilla(datos[1]); // L
//     row.getCell('F').value = formatearCasilla(datos[2]); // M
//     row.getCell('H').value = formatearCasilla(datos[3]); // N
//     row.getCell('J').value = formatearCasilla(datos[4]); // O
//     row.getCell('L').value = formatearCasilla(datos[5]); // P

//     try {
//         hojaDestino.mergeCells(filaActual, 2, filaActual, 3);
//         hojaDestino.mergeCells(filaActual, 4, filaActual, 5);
//         hojaDestino.mergeCells(filaActual, 6, filaActual, 7);
//         hojaDestino.mergeCells(filaActual, 8, filaActual, 9);
//         hojaDestino.mergeCells(filaActual, 10, filaActual, 11);
//         hojaDestino.mergeCells(filaActual, 12, filaActual, 13);
//         aplicarEstiloCelda(row);
//     } catch (e) {}

//     return filaActual;
// }

// module.exports = s6_2;

const { aplicarEstiloCelda } = require('../../utils/excelTools');

function s6_2(hojaDestino, fPestana6, filaInicio) {
    if (!fPestana6) return filaInicio;

    // --- PROTECCIÓN ANTI-NaN ---
    // Localización dinámica mediante el texto del encabezado
    let filaSeccion = -1;
    hojaDestino.eachRow((row, rowNumber) => {
        const valorB = row.getCell('B').value;
        if (valorB && valorB.toString().includes("6.2. Descripción de Torre")) {
            filaSeccion = rowNumber;
        }
    });

    // Inicio 2 filas después del encabezado o respaldo en filaInicio + 4
    let filaActual = (filaSeccion !== -1) ? (filaSeccion + 2) : (filaInicio + 4);
    
    // Validación de seguridad para evitar Address: BNaN
    if (isNaN(filaActual)) filaActual = 200; 

    // Extraemos el bloque de datos desde K:P en la pestaña 6
    const colInicio = fPestana6.getCell('K').col; 
    const datos = [];
    for (let i = 0; i < 6; i++) {
        datos.push(fPestana6.getCell(colInicio + i).value);
    }

    /**
     * Función para formatear SI/NO con casillas visuales
     * Soporta valores booleanos, texto y marcas 'X'
     */
    const formatearCasilla = (val) => {
        const v = val ? val.toString().trim().toUpperCase() : "";
        if (['SI', 'S', 'X', '1', 'VERDADERO'].includes(v)) {
            return "SI (  X  )  NO (     )";
        }
        return "SI (     )  NO (  X  )";
    };

    const row = hojaDestino.getRow(filaActual);
    row.height = 35; // Altura para descripciones técnicas

    // Reset y Estilo base de B a M
    for (let c = 2; c <= 13; c++) {
        const cell = row.getCell(c);
        cell.border = { 
            top: {style:'thin'}, left: {style:'thin'}, 
            bottom: {style:'thin'}, right: {style:'thin'} 
        };
        cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        cell.font = { name: 'Arial', size: 11 };
    }

    // --- MAPEO DE COLUMNAS POR PARES (B-M) ---
    row.getCell('B').value = datos[0]; // Altura de torre (m) (K) -> B-C
    
    // Parámetros de instalación (L a P) mapeados a casillas SI/NO
    row.getCell('D').value = formatearCasilla(datos[1]); // Torre galvanizada (L) -> D-E
    row.getCell('F').value = formatearCasilla(datos[2]); // Soportes cableado (M) -> F-G
    row.getCell('H').value = formatearCasilla(datos[3]); // Luz de Balizaje (N) -> H-I
    row.getCell('J').value = formatearCasilla(datos[4]); // Escalera peatonal (O) -> J-K
    row.getCell('L').value = formatearCasilla(datos[5]); // Pararrayos (P) -> L-M

    // Aplicar merges por pares para mantener la simetría del acta
    try {
        hojaDestino.mergeCells(filaActual, 2, filaActual, 3);   // Altura
        hojaDestino.mergeCells(filaActual, 4, filaActual, 5);   // Galvanizada
        hojaDestino.mergeCells(filaActual, 6, filaActual, 7);   // Soportes
        hojaDestino.mergeCells(filaActual, 8, filaActual, 9);   // Balizaje
        hojaDestino.mergeCells(filaActual, 10, filaActual, 11); // Escalera
        hojaDestino.mergeCells(filaActual, 12, filaActual, 13); // Pararrayos
    } catch (e) {}

    return filaActual;
}

module.exports = s6_2;