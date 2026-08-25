const { aplicarEstiloCelda, limpiarMergesFila } = require('../../utils/excelTools');

function s2_4(hojaDestino, fEquip, filaReferenciaAnterior) {
    let filaSeccion = -1;
    hojaDestino.eachRow((row, rowNumber) => {
        const valorB = row.getCell('B').value;
        if (valorB && valorB.toString().includes("2.4. Descripción de Enlaces (Ópticos)")) {
            filaSeccion = rowNumber;
        }
    });

    let filaActual = filaSeccion !== -1 ? filaSeccion + 2 : filaReferenciaAnterior + 3;
    if (!fEquip) return filaActual;

    const FILA_BASE_PLANTILLA = filaActual;
    const iniciosColumnas = ['UC', 'UJ', 'UQ', 'UX', 'VE', 'VL'];
    let itemCorrelativo = 1;

   iniciosColumnas.forEach((letraInicio, index) => {
        const celdaMatrizBase = fEquip.getCell(letraInicio);
        const valorRef = celdaMatrizBase.value;

        // 1. TRATAMIENTO SEGURO DEL VALOR
        const valorMatriz = valorRef ? valorRef.toString().trim() : "";
        const valorUpper = valorMatriz.toUpperCase();
        
        // 2. FILTRO NA: Si es NA, lo ignoramos totalmente
        if (valorUpper === 'NA') return;

        // 3. CONTROL DE FILAS VACÍAS: 
        // Si la celda está vacía pero NO es la primera (index > 0), saltamos para no repetir 6 filas de guiones.
        if (valorMatriz === "" && index > 0) return;

        // --- A partir de aquí se crea la fila ---
        if (filaActual > FILA_BASE_PLANTILLA) {
            hojaDestino.insertRow(filaActual, {});
        }

        const row = hojaDestino.getRow(filaActual);
        limpiarMergesFila(hojaDestino, filaActual);
        aplicarEstiloCelda(row);

        const colIdx = celdaMatrizBase.col;
        const esInvalido = valorMatriz === "-" || valorMatriz === "0" || valorMatriz === "";
        if (esInvalido) {
            row.getCell('B').value = "-";
            row.getCell('C').value = "-";
            row.getCell('E').value = "-";
            row.getCell('F').value = "-";
            row.getCell('G').value = "-";
            row.getCell('I').value = "-";
            row.getCell('K').value = "-";
            row.getCell('M').value = "-";
        } else {
            row.getCell('B').value = itemCorrelativo++;
            row.getCell('C').value = fEquip.getCell(colIdx).value || "-";
            row.getCell('E').value = fEquip.getCell(colIdx + 1).value || "-";
            row.getCell('F').value = fEquip.getCell(colIdx + 2).value || "-";
            row.getCell('G').value = fEquip.getCell(colIdx + 3).value || "-";
            row.getCell('I').value = fEquip.getCell(colIdx + 4).value || "-";
            row.getCell('K').value = fEquip.getCell(colIdx + 5).value || "-";
            row.getCell('M').value = fEquip.getCell(colIdx + 6).value || "-";
        }

        try {
            hojaDestino.mergeCells(filaActual, 3, filaActual, 4);
            hojaDestino.mergeCells(filaActual, 7, filaActual, 8);
            hojaDestino.mergeCells(filaActual, 9, filaActual, 10);
            hojaDestino.mergeCells(filaActual, 11, filaActual, 12);
        } catch (e) { }
        row.height = undefined;
        filaActual++;
    });

    return filaActual;
}

module.exports = s2_4;