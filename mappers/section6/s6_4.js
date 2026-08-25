// const { aplicarEstiloCelda } = require('../../utils/excelTools');

// function s6_4(hojaDestino, fPestana6, filaInicio) {
//     if (!fPestana6) return filaInicio;

//     // Ajuste de filas según tu modelo
//     let filaSuperior = filaInicio + 2;
//     let filaIntermedia = filaSuperior + 1; // Fila de etiquetas (Puerta, Climatización...)
//     let filaInferior = filaSuperior + 2;    // Fila de los SI/NO

//     // --- EXTRACCIÓN DE DATOS ---
//     const material = fPestana6.getCell('AO').value;
//     const dimensiones = fPestana6.getCell('AP').value;

//     const datosCheck = [
//         fPestana6.getCell('AQ').value,
//         fPestana6.getCell('AR').value,
//         fPestana6.getCell('AS').value,
//         fPestana6.getCell('AT').value,
//         fPestana6.getCell('AU').value
//     ];

//     const formatearCasilla = (val) => {
//         const v = val ? val.toString().trim().toUpperCase() : "";
//         if (v === 'SI' || v === 'S' || v === 'X') return "SI (  X  )  NO (     )";
//         return "SI (     )  NO (  X  )";
//     };

//     // --- ESCRITURA DE DATOS ---
//     const rowSup = hojaDestino.getRow(filaSuperior);
//     rowSup.getCell('D').value = material;
//     rowSup.getCell('J').value = dimensiones;

//     const rowInf = hojaDestino.getRow(filaInferior);
//     rowInf.height = 25;
//     rowInf.getCell('B').value = formatearCasilla(datosCheck[0]);
//     rowInf.getCell('D').value = formatearCasilla(datosCheck[1]);
//     rowInf.getCell('F').value = formatearCasilla(datosCheck[2]);
//     rowInf.getCell('J').value = formatearCasilla(datosCheck[3]);
//     rowInf.getCell('L').value = formatearCasilla(datosCheck[4]);

//     // --- APLICACIÓN DE ESTILOS Y NEGRITAS ---
//     // Definimos qué columnas deben mantener la NEGRITA (etiquetas fijas)
//     const columnasNegrita = [2, 7, 10]; // B: Material, G: Dimensiones, etc.

//     [filaSuperior, filaIntermedia, filaInferior].forEach((nFila) => {
//         const row = hojaDestino.getRow(nFila);

//         // Recorremos las columnas de la B a la M (2 a 13)
//         for (let i = 2; i <= 13; i++) {
//             const cell = row.getCell(i);

//             // Mantenemos el borde y alineación
//             cell.border = {
//                 top: { style: 'thin' }, left: { style: 'thin' },
//                 bottom: { style: 'thin' }, right: { style: 'thin' }
//             };
//             cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };

//             // Aplicamos NEGRITA si es una columna de etiqueta o si es la fila intermedia de títulos
//             const esEtiqueta = columnasNegrita.includes(i);
//             const esFilaTitulos = (nFila === filaIntermedia);

//             cell.font = {
//                 name: 'Arial',
//                 size: 11,
//                 bold: esEtiqueta || esFilaTitulos
//             };
//         }
//     });

//     return filaInferior;
// }

// module.exports = s6_4;
const { aplicarEstiloCelda } = require('../../utils/excelTools');

function s6_4(hojaDestino, fPestana6, filaInicio) {
    if (!fPestana6) return filaInicio;

    // --- PROTECCIÓN ANTI-NaN Y AJUSTE DE POSICIÓN ---
    let filaSeccion = -1;
    hojaDestino.eachRow((row, rowNumber) => {
        const valorB = row.getCell('B').value;
        if (valorB && valorB.toString().includes("6.4. Shelter")) {
            filaSeccion = rowNumber;
        }
    });

    // AJUSTE: Si se encuentra la sección, empezamos 1 fila después (no 2)
    // para que el contenido encaje justo debajo del encabezado.
    let filaSuperior = (filaSeccion !== -1) ? (filaSeccion + 1) : (filaInicio + 1);

    if (isNaN(filaSuperior)) filaSuperior = 240;

    let filaIntermedia = filaSuperior + 1;
    let filaInferior = filaSuperior + 2;

    // --- EXTRACCIÓN DE DATOS ---
    const material = fPestana6.getCell('AO').value;
    const dimensiones = fPestana6.getCell('AP').value;

    const datosCheck = [
        fPestana6.getCell('AQ').value,
        fPestana6.getCell('AR').value,
        fPestana6.getCell('AS').value,
        fPestana6.getCell('AT').value,
        fPestana6.getCell('AU').value
    ];

    const formatearCasilla = (val) => {
        const v = val ? val.toString().trim().toUpperCase() : "";
        if (['SI', 'S', 'X', '1', 'VERDADERO'].includes(v)) return "SI (  X  )  NO (     )";
        return "SI (     )  NO (  X  )";
    };

    // --- ESCRITURA ---
    const rowSup = hojaDestino.getRow(filaSuperior);
    rowSup.height = 30; // Reducido ligeramente para compactar
    rowSup.getCell('D').value = material;
    rowSup.getCell('J').value = dimensiones;

    const rowInf = hojaDestino.getRow(filaInferior);
    rowInf.height = 25;
    rowInf.getCell('B').value = formatearCasilla(datosCheck[0]);
    rowInf.getCell('D').value = formatearCasilla(datosCheck[1]);
    rowInf.getCell('F').value = formatearCasilla(datosCheck[2]);
    rowInf.getCell('J').value = formatearCasilla(datosCheck[3]);
    rowInf.getCell('L').value = formatearCasilla(datosCheck[4]);

    // --- ESTILOS ---
    const filasProcesar = [filaSuperior, filaIntermedia, filaInferior];
    filasProcesar.forEach((nFila) => {
        const row = hojaDestino.getRow(nFila);
        for (let i = 2; i <= 13; i++) {
            const cell = row.getCell(i);
            cell.border = {
                top: { style: 'thin' }, left: { style: 'thin' },
                bottom: { style: 'thin' }, right: { style: 'thin' }
            };
            cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
            cell.font = { name: 'Arial', size: 11, bold: (nFila === filaIntermedia || [2, 7, 10].includes(i)) };
        }
    });

    // Merges manuales solo si son necesarios
    try {
        if (!rowSup.getCell('D').isMerged) hojaDestino.mergeCells(filaSuperior, 4, filaSuperior, 6);
        if (!rowSup.getCell('J').isMerged) hojaDestino.mergeCells(filaSuperior, 10, filaSuperior, 13);
        if (!rowInf.getCell('D').isMerged) hojaDestino.mergeCells(filaInferior, 4, filaInferior, 5);
        if (!rowInf.getCell('F').isMerged) hojaDestino.mergeCells(filaInferior, 6, filaInferior, 9);
        if (!rowInf.getCell('L').isMerged) hojaDestino.mergeCells(filaInferior, 12, filaInferior, 13);
    } catch (e) { }

    return filaInferior;
}

module.exports = s6_4;