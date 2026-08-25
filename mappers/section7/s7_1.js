// const { aplicarEstiloCelda } = require('../../utils/excelTools');

// function s7_1(hojaDestino, fPestana6, filaInicio) {
//     if (!fPestana6) return filaInicio;

//     // +5 filas después de la sección anterior
//     let fila1 = filaInicio + 6;
//     let fila2 = fila1 + 1;

//     // --- EXTRACCIÓN DE DATOS ---
//     const datosRaw = {
//         e: fPestana6.getCell('E').value,
//         f: fPestana6.getCell('F').value,
//         g: fPestana6.getCell('G').value,
//         h: fPestana6.getCell('H').value
//     };

//     // Función para el formato de casillas SI/NO
//     const formatearCasilla = (val) => {
//         const v = val ? val.toString().trim().toUpperCase() : "";
//         if (v === 'SI' || v === 'S' || v === 'X') return "SI (  X  )  NO (     )";
//         return "SI (     )  NO (  X  )";
//     };

//     // --- MAPEO AL ACTA ---
//     const row1 = hojaDestino.getRow(fila1);
//     row1.getCell('F').value = formatearCasilla(datosRaw.e);
//     row1.getCell('L').value = formatearCasilla(datosRaw.f);

//     const row2 = hojaDestino.getRow(fila2);
//     row2.getCell('F').value = formatearCasilla(datosRaw.g);
//     row2.getCell('L').value = formatearCasilla(datosRaw.h);

//     // --- ESTILOS Y ALTURA DE FILA ---
//     [row1, row2].forEach(row => {
//         // Establecer altura solicitada
//         row.height = 37.5;

//         // Columnas F(6) y L(12) que son los inicios de los merges
//         [6, 12].forEach(colIdx => {
//             const cell = row.getCell(colIdx);
//             cell.border = {
//                 top: { style: 'thin' }, left: { style: 'thin' },
//                 bottom: { style: 'thin' }, right: { style: 'thin' }
//             };
//             cell.alignment = {
//                 vertical: 'middle',
//                 horizontal: 'center',
//                 wrapText: true
//             };
//             cell.font = {
//                 name: 'Arial',
//                 size: 11,

//             };
//         });
//     });

//     return fila2;
// }

// module.exports = s7_1;

const { aplicarEstiloCelda } = require('../../utils/excelTools');

function s7_1(hojaDestino, fPestana6) {
    if (!fPestana6) return;

    let filaSeccion = -1;
    hojaDestino.eachRow((row, rowNumber) => {
        const valorB = row.getCell('B').value;
        // Buscamos el ancla en la columna B
        if (valorB && valorB.toString().includes("7.1.  Carácteristicas Generales")) {
            filaSeccion = rowNumber;
        }
    });

    if (filaSeccion === -1) return;

    // "Pégalo 1 fila abajo": Si el título está en X, los datos empiezan en X + 2
    // (X+1 suele ser la cabecera de la tabla o espacio, X+2 el primer dato)
    let fila1 = filaSeccion + 2;
    let fila2 = fila1 + 1;

    const datosRaw = {
        e: fPestana6.getCell('E').value,
        f: fPestana6.getCell('F').value,
        g: fPestana6.getCell('G').value,
        h: fPestana6.getCell('H').value
    };

    const formatearCasilla = (val) => {
        const v = val ? val.toString().trim().toUpperCase() : "";
        return (['SI', 'S', 'X', '1', 'VERDADERO'].includes(v))
            ? "SI (  X  )  NO (     )"
            : "SI (     )  NO (  X  )";
    };

    const row1 = hojaDestino.getRow(fila1);
    row1.height = 37.5;
    row1.getCell('F').value = formatearCasilla(datosRaw.e);
    row1.getCell('L').value = formatearCasilla(datosRaw.f);

    const row2 = hojaDestino.getRow(fila2);
    row2.height = 37.5;
    row2.getCell('F').value = formatearCasilla(datosRaw.g);
    row2.getCell('L').value = formatearCasilla(datosRaw.h);

    [row1, row2].forEach(row => {
        [6, 12].forEach(colIdx => {
            const cell = row.getCell(colIdx);
            cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
            cell.font = { name: 'Arial', size: 11 };
        });
    });
}

module.exports = s7_1;