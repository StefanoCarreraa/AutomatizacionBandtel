// const { aplicarEstiloCelda } = require('../../utils/excelTools');

// function s7_2(hojaDestino, fPestana6, filaInicio) {
//     if (!fPestana6) return filaInicio;

//     // --- BLOQUE 1: ELECTRODOS ---
//     // +3 filas después de la sección anterior (7.1)
//     let filaElectrodo = filaInicio + 4;

//     // --- BLOQUE 2: DATOS TÉCNICOS ---
//     // +2 filas después del bloque de electrodos
//     let filaDatos = filaElectrodo + 2;

//     // 1. Lógica para determinar el tipo de Electrodo
//     const valI = fPestana6.getCell('I').value;
//     const esVertical = valI && valI.toString().trim().toUpperCase() === 'SI';

//     // Texto con espaciado amplio para cubrir el ancho de B:M
//     let textoElectrodo = `Electrodo Vertical (  ${esVertical ? 'X' : ' '}  )                                                                                                    Electrodo Horizontal (  ${esVertical ? ' ' : 'X'}  )`;

//     // 2. Extracción de datos de la matriz
//     const datosTecnicos = {
//         k: fPestana6.getCell('K').value, // Para D:E
//         l: fPestana6.getCell('L').value, // Para F:G
//         m: fPestana6.getCell('M').value, // Para H:I
//         n: fPestana6.getCell('N').value  // Para J:K (SI/NO)
//     };

//     const formatearCasilla = (val) => {
//         const v = val ? val.toString().trim().toUpperCase() : "";
//         if (v === 'SI' || v === 'S' || v === 'X') return "SI (  X  )  NO (     )";
//         return "SI (     )  NO (  X  )";
//     };

//     // --- ESCRITURA EN EL ACTA ---

//     // Fila Electrodo (B:M)
//     const rowEl = hojaDestino.getRow(filaElectrodo);
//     rowEl.height = 30;
//     rowEl.getCell('B').value = textoElectrodo;

//     // Fila Datos Técnicos (D:E, F:G, H:I, J:K)
//     const rowDt = hojaDestino.getRow(filaDatos);
//     rowDt.height = 39.5; // Altura solicitada para el último bloque

//     rowDt.getCell('D').value = datosTecnicos.k;
//     rowDt.getCell('F').value = datosTecnicos.l;
//     rowDt.getCell('H').value = datosTecnicos.m;
//     rowDt.getCell('J').value = formatearCasilla(datosTecnicos.n);

//     // --- ESTILOS ---

//     // Estilo Electrodo
//     const cellEl = rowEl.getCell(2);
//     cellEl.font = { name: 'Arial', size: 10, bold: true };
//     cellEl.alignment = { vertical: 'middle', horizontal: 'center' };

//     // Estilo Datos Técnicos con bordes y negrita
//     [4, 6, 8, 10].forEach(colIdx => {
//         const cell = rowDt.getCell(colIdx);
//         cell.border = {
//             top: { style: 'thin' }, left: { style: 'thin' },
//             bottom: { style: 'thin' }, right: { style: 'thin' }
//         };
//         cell.alignment = {
//             vertical: 'middle',
//             horizontal: 'center',
//             wrapText: true
//         };
//         cell.font = {
//             name: 'Arial',
//             size: 10,
//         };
//     });

//     return filaDatos;
// }

// module.exports = s7_2;

const { aplicarEstiloCelda } = require('../../utils/excelTools');

/**
 * Función para llenar la sección 7.2. Mediciones del Sistema SPAT
 * @param {Object} hojaDestino - La hoja del acta donde se escribirá
 * @param {Object} fPestana6 - La pestaña de la matriz con los datos de origen
 */
function s7_2(hojaDestino, fPestana6) {
    if (!fPestana6) return;

    // 1. BUSCAR DINÁMICAMENTE LA FILA DEL TÍTULO EN LA COLUMNA B
    let filaTitulo = -1;
    hojaDestino.eachRow((row, rowNumber) => {
        const valorB = row.getCell('B').value;
        if (valorB && valorB.toString().includes('7.2.  Mediciones (*)')) {
            filaTitulo = rowNumber;
        }
    });

    // Si no encuentra la sección, termina la función para evitar errores
    if (filaTitulo === -1) {
        console.warn("No se encontró la sección 7.2. Mediciones (*) en la columna B.");
        return;
    }

    // --- DEFINICIÓN DE FILAS BASADAS EN EL HALLAZGO ---
    // Según image_f14054.png:
    // Fila 120: Selección de Electrodo (Vertical/Horizontal)
    // Fila 122: Valores de Medida 1, 2, 3 y Cumplimiento
    let filaElectrodo = filaTitulo + 1;
    let filaDatos = filaTitulo + 3; // Salta la fila de cabeceras de la tabla

    // 2. LÓGICA PARA EL TIPO DE ELECTRODO (Proviene de la columna I de la matriz)
    const valI = fPestana6.getCell('I').value;
    const esVertical = valI && valI.toString().trim().toUpperCase() === 'SI';

    // Construcción del texto con espacios para alinear "X" en los paréntesis del acta
    let textoElectrodo = `Electrodo Vertical (  ${esVertical ? 'X' : ' '}  )                                                                    Electrodo Horizontal (  ${esVertical ? ' ' : 'X'}  )`;

    // 3. EXTRACCIÓN DE DATOS TÉCNICOS DE LA MATRIZ
    const datosTecnicos = {
        medida1: fPestana6.getCell('K').value, // Medida 1 (52%)
        medida2: fPestana6.getCell('L').value, // Medida 2 (62%)
        medida3: fPestana6.getCell('M').value, // Medida 3 (72%)
        cumple: fPestana6.getCell('N').value   // Resultado SI/NO
    };

    const formatearCheck = (val) => {
        const v = val ? val.toString().trim().toUpperCase() : "";
        if (v === 'SI' || v === 'S' || v === 'X') return "SI (  X  )  NO (     )";
        return "SI (     )  NO (  X  )";
    };

    // --- ESCRITURA EN EL ACTA ---

    // Escribir tipo de Electrodo (Combinación visual en B)
    const rowEl = hojaDestino.getRow(filaElectrodo);
    rowEl.getCell('B').value = textoElectrodo;

    // Escribir Mediciones y Cumplimiento
    const rowDt = hojaDestino.getRow(filaDatos);
    rowDt.height = 35; // Ajuste de altura para visibilidad

    rowDt.getCell('D').value = datosTecnicos.medida1;
    rowDt.getCell('F').value = datosTecnicos.medida2;
    rowDt.getCell('H').value = datosTecnicos.medida3;
    rowDt.getCell('J').value = formatearCheck(datosTecnicos.cumple);

    // --- APLICACIÓN DE ESTILOS ---

    // Estilo para la fila de Electrodos
    const cellEl = rowEl.getCell('B');
    cellEl.font = { name: 'Arial', size: 11, };
    cellEl.alignment = { vertical: 'middle', horizontal: 'left' };

    // Estilo para las celdas de datos (D, F, H, J) con bordes
    [4, 6, 8, 10].forEach(colIdx => {
        const cell = rowDt.getCell(colIdx);
        cell.border = {
            top: { style: 'thin' }, left: { style: 'thin' },
            bottom: { style: 'thin' }, right: { style: 'thin' }
        };
        cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        cell.font = { name: 'Arial', size: 11 };
    });

    return filaDatos; // Retorna la última fila modificada
}

module.exports = s7_2;