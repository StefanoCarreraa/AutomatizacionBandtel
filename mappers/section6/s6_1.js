// const { aplicarEstiloCelda } = require('../../utils/excelTools');

// function s6_1(hojaDestino, fPestana6, filaInicio) {
//     if (!fPestana6) return filaInicio;

//     // +4 filas después de la sección anterior (Estático)
//     let filaActual = filaInicio + 4;
    
//     const colInicio = fPestana6.getCell('E').col; // Rango E:J
//     const datos = [];
//     for (let i = 0; i < 6; i++) {
//         datos.push(fPestana6.getCell(colInicio + i).value);
//     }

//     // Función para formatear SI/NO con casillas
//     const formatearCasilla = (val) => {
//         const v = val ? val.toString().trim().toUpperCase() : "";
//         if (v === 'SI' || v === 'S' || v === 'X') return "SI (  X  )  NO (     )";
//         return "SI (     )  NO (  X  )";
//     };

//     const row = hojaDestino.getRow(filaActual);
//     row.height = 25;

//     // Mapeo de pares combinados B:C, D:E, F:G, H:I, J:K, L:M
//     row.getCell('B').value = datos[0]; // E
//     row.getCell('D').value = datos[1]; // F
//     row.getCell('F').value = datos[2]; // G
    
//     // H e I son SI/NO según tu instrucción (mapeados a la columna H e J del acta)
//     row.getCell('H').value = formatearCasilla(datos[3]); // H de matriz
//     row.getCell('J').value = formatearCasilla(datos[4]); // I de matriz
    
//     row.getCell('L').value = datos[5]; // J de matriz

//     // Aplicar estilos y merges (Sin insertar fila)
//     try {
//         hojaDestino.mergeCells(filaActual, 2, filaActual, 3);   // B-C
//         hojaDestino.mergeCells(filaActual, 4, filaActual, 5);   // D-E
//         hojaDestino.mergeCells(filaActual, 6, filaActual, 7);   // F-G
//         hojaDestino.mergeCells(filaActual, 8, filaActual, 9);   // H-I
//         hojaDestino.mergeCells(filaActual, 10, filaActual, 11); // J-K
//         hojaDestino.mergeCells(filaActual, 12, filaActual, 13); // L-M
//         aplicarEstiloCelda(row);
//     } catch (e) {}

//     return filaActual;
// }

// module.exports = s6_1;

const { aplicarEstiloCelda } = require('../../utils/excelTools');

function s6_1(hojaDestino, fPestana6, filaInicio) {
    if (!fPestana6) return filaInicio;

    // --- PROTECCIÓN ANTI-NaN ---
    // Localización dinámica mediante el texto del encabezado
    let filaSeccion = -1;
    hojaDestino.eachRow((row, rowNumber) => {
        const valorB = row.getCell('B').value;
        if (valorB && valorB.toString().includes("6.1. Cerco Perimétrico")) {
            filaSeccion = rowNumber;
        }
    });

    // Inicio 2 filas después del encabezado o respaldo en filaInicio + 4
    let filaActual = (filaSeccion !== -1) ? (filaSeccion + 2) : (filaInicio + 4);
    
    // Validación de seguridad para evitar Address: BNaN
    if (isNaN(filaActual)) filaActual = 180; 

    // Extraemos el bloque único de datos desde E:J en la pestaña 6
    const colInicio = fPestana6.getCell('E').col; 
    const datos = [];
    for (let i = 0; i < 6; i++) {
        datos.push(fPestana6.getCell(colInicio + i).value);
    }

    /**
     * Función para formatear SI/NO con casillas visuales
     * Basado en la estructura de la imagen image_f04f01.png
     */
    const formatearCasilla = (val) => {
        const v = val ? val.toString().trim().toUpperCase() : "";
        if (['SI', 'S', 'X', '1', 'VERDADERO'].includes(v)) {
            return "SI (  X  )  NO (     )";
        }
        return "SI (     )  NO (  X  )";
    };

    const row = hojaDestino.getRow(filaActual);
    row.height = 35; // Altura mayor para textos largos de obras civiles

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
    row.getCell('B').value = datos[0]; // Tipo de Cerco (E) -> B-C
    row.getCell('D').value = datos[1]; // Área (F) -> D-E
    row.getCell('F').value = datos[2]; // Altura (G) -> F-G
    
    // Formateo de casillas SI/NO para Concertina y Puerta
    row.getCell('H').value = formatearCasilla(datos[3]); // Concertina (H) -> H-I
    row.getCell('J').value = formatearCasilla(datos[4]); // Puerta/Marco (I) -> J-K
    
    row.getCell('L').value = datos[5]; // Dimensiones Puerta (J) -> L-M

    // Aplicar merges por pares
    try {
        hojaDestino.mergeCells(filaActual, 2, filaActual, 3);   // B-C
        hojaDestino.mergeCells(filaActual, 4, filaActual, 5);   // D-E
        hojaDestino.mergeCells(filaActual, 6, filaActual, 7);   // F-G
        hojaDestino.mergeCells(filaActual, 8, filaActual, 9);   // H-I
        hojaDestino.mergeCells(filaActual, 10, filaActual, 11); // J-K
        hojaDestino.mergeCells(filaActual, 12, filaActual, 13); // L-M
    } catch (e) {}

    return filaActual;
}

module.exports = s6_1;