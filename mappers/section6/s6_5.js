// const { aplicarEstiloCelda } = require('../../utils/excelTools');

// function s6_5(hojaDestino, fPestana6, filaInicio) {
//     if (!fPestana6) return filaInicio;

//     // --- PRIMER BLOQUE (AV) ---
//     // +2 filas después de la sección anterior (6.4)
//     let filaAV = filaInicio + 3;

//     // --- SEGUNDO BLOQUE (AW, AX, AY) ---
//     // +2 filas después de la fila AV
//     let filaFinal = filaAV + 2;

//     // Extracción de datos de la matriz
//     const datoAV = fPestana6.getCell('AV').value;
//     const datoAW = fPestana6.getCell('AW').value;
//     const datoAX = fPestana6.getCell('AX').value;
//     const datoAY = fPestana6.getCell('AY').value;

//     // Formato estándar SI/NO
//     const formatearCasilla = (val) => {
//         const v = val ? val.toString().trim().toUpperCase() : "";
//         if (v === 'SI' || v === 'S' || v === 'X') return "SI (  X  )  NO (     )";
//         return "SI (     )  NO (  X  )";
//     };

//     // Formato EXTENDIDO para AV
//     const formatearCasillaExtendida = (val) => {
//         const v = val ? val.toString().trim().toUpperCase() : "";
//         if (v === 'SI' || v === 'S' || v === 'X') return "SI (  X  )                          NO (     )";
//         return "SI (     )                          NO (  X  )";
//     };

//     // --- ESCRITURA FILA AV ---
//     const rowAV = hojaDestino.getRow(filaAV);
//     rowAV.height = 25;
//     // H:M es AV (Escribimos en H porque es la celda de inicio del merge)
//     rowAV.getCell('H').value = formatearCasillaExtendida(datoAV);

//     // --- ESCRITURA FILA AW, AX, AY ---
//     const rowFin = hojaDestino.getRow(filaFinal);
//     rowFin.height = 25;
//     rowFin.getCell('B').value = formatearCasilla(datoAW); // B:E es AW
//     rowFin.getCell('F').value = datoAX;                  // F:I es AX
//     rowFin.getCell('J').value = datoAY;                  // J:M es AY

//     // --- ESTILOS Y NEGRITAS (Imagen 1) ---
//     const filasParaEstilo = [
//         { row: rowAV, cols: [8] },             // Solo H
//         { row: rowFin, cols: [2, 6, 10] }      // B, F, J
//     ];

//     filasParaEstilo.forEach(item => {
//         item.cols.forEach(colIdx => {
//             const cell = item.row.getCell(colIdx);
//             cell.border = {
//                 top: { style: 'thin' }, left: { style: 'thin' },
//                 bottom: { style: 'thin' }, right: { style: 'thin' }
//             };
//             cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
//             cell.font = { name: 'Arial', size: 11 }; // Negrita según Imagen 1
//         });
//     });

//     // Nota: No incluimos mergeCells para evitar el error "already merged"
//     // ya que el modelo base ya debería tener las celdas combinadas.

//     return filaFinal;
// }

// module.exports = s6_5;

const { aplicarEstiloCelda } = require('../../utils/excelTools');

function s6_5(hojaDestino, fPestana6, filaInicio) {
    if (!fPestana6) return filaInicio;

    // --- LOCALIZACIÓN DINÁMICA ---
    let filaSeccion = -1;
    hojaDestino.eachRow((row, rowNumber) => {
        const valorB = row.getCell('B').value;
        if (valorB && valorB.toString().includes("6.5. Suministro de Energía")) {
            filaSeccion = rowNumber;
        }
    });

    // AJUSTE: Reducimos de +2 a +1 para pegar el contenido al encabezado
    let filaAV = (filaSeccion !== -1) ? (filaSeccion + 1) : (filaInicio + 2);

    if (isNaN(filaAV)) filaAV = 260;
    let filaFinal = filaAV + 2; // Mantiene la separación interna del bloque de energía

    // Extracción de datos
    const datoAV = fPestana6.getCell('AV').value;
    const datoAW = fPestana6.getCell('AW').value;
    const datoAX = fPestana6.getCell('AX').value;
    const datoAY = fPestana6.getCell('AY').value;

    const formatearCasilla = (val) => {
        const v = val ? val.toString().trim().toUpperCase() : "";
        if (['SI', 'S', 'X', '1', 'VERDADERO'].includes(v)) return "SI (  X  )  NO (     )";
        return "SI (     )  NO (  X  )";
    };

    const formatearCasillaExtendida = (val) => {
        const v = val ? val.toString().trim().toUpperCase() : "";
        if (['SI', 'S', 'X', '1', 'VERDADERO'].includes(v)) {
            return "SI (  X  )                   NO (     )";
        }
        return "SI (     )                   NO (  X  )";
    };

    // --- ESCRITURA ---
    const rowAV = hojaDestino.getRow(filaAV);
    rowAV.height = 25;
    rowAV.getCell('H').value = formatearCasillaExtendida(datoAV);

    const rowFin = hojaDestino.getRow(filaFinal);
    rowFin.height = 28; // Ajuste leve de altura
    rowFin.getCell('B').value = formatearCasilla(datoAW);
    rowFin.getCell('F').value = datoAX;
    rowFin.getCell('J').value = datoAY;

    // --- ESTILOS ---
    const mapeoEstilos = [
        { row: rowAV, cols: [8] },
        { row: rowFin, cols: [2, 6, 10] }
    ];

    mapeoEstilos.forEach(item => {
        item.cols.forEach(colIdx => {
            const cell = item.row.getCell(colIdx);
            cell.border = {
                top: { style: 'thin' }, left: { style: 'thin' },
                bottom: { style: 'thin' }, right: { style: 'thin' }
            };
            cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
            cell.font = { name: 'Arial', size: 11 };
        });
    });

    return filaFinal;
}

module.exports = s6_5;