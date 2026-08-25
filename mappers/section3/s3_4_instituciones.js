const { aplicarEstiloCelda } = require('../../utils/excelTools');

function s3_4(hojaDestino, fPop, filaInicio) {
    if (!fPop) return filaInicio;
    let filaSeccion = -1;
    hojaDestino.eachRow((row, rowNumber) => {
        const valorB = row.getCell('B').value;
        if (valorB && valorB.toString().includes("3.4. Instituciones Abonadas Obligatorias")) {
            filaSeccion = rowNumber;
        }
    });
    let filaActual = (filaSeccion !== -1) ? (filaSeccion + 2) : (filaInicio + 4);
    if (isNaN(filaActual)) filaActual = 60;
    const bloquesMatriz3_4 = [
        { start: 'EH' }, { start: 'EO' }, { start: 'EV' },
        { start: 'FC' }, { start: 'FJ' }, { start: 'FQ' },
        { start: 'FX' }, { start: 'GE' }, { start: 'GL' }
    ];
    const esValorVacio = (val) => {
        if (!val) return true;
        const v = val.toString().trim().toUpperCase();
        return v === 'NA' || v === '-' || v === 'N/A' || v === '';
    };
    const limpiarMergesFila = (numFila) => {
        // 1. Forzar desvinculación por rango físico (B a M)
        // unMergeCells lanzará error si el rango no está combinado, por eso usamos try-catch
        try {
            hojaDestino.unMergeCells(numFila, 2, numFila, 13);
        } catch (e) {
            // Silenciamos el error si no hay nada que descombinar
        }

        // 2. Limpieza de metadatos internos de la celda
        const row = hojaDestino.getRow(numFila);
        for (let c = 2; c <= 13; c++) {
            const cell = row.getCell(c);
            // Esto es vital: obliga a la celda a ser su propia "maestra",
            // rompiendo cualquier vínculo fantasma de merges previos.
            cell.master = cell;
        }

        // 3. Limpieza manual del índice interno (por si acaso)
        if (hojaDestino._merges) {
            Object.keys(hojaDestino._merges).forEach(key => {
                const m = hojaDestino._merges[key];
                if (m.top === numFila || m.bottom === numFila) {
                    delete hojaDestino._merges[key];
                }
            });
        }
    };
    const escribirFilaS3_4 = (numFila, valores, numItem, insertar, modoNoAplica = false) => {
        if (insertar) {
            hojaDestino.insertRow(numFila, {});
        }

        limpiarMergesFila(numFila);
        const row = hojaDestino.getRow(numFila);
        row.height = 25;

        // --- RESET Y ESTILOS ---
        for (let c = 2; c <= 13; c++) {
            const cell = row.getCell(c);
            cell.value = null;
            cell.border = {
                top: { style: 'thin' }, left: { style: 'thin' },
                bottom: { style: 'thin' }, right: { style: 'thin' }
            };
            cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
            cell.font = { name: 'Arial', size: 11 };
        }

        // --- ASIGNACIÓN DE VALORES ---
        if (modoNoAplica) {
            row.getCell('B').value = "-"; // N°
            row.getCell('C').value = "-"; // ID IAO (C-D)
            row.getCell('E').value = "-"; // Localidad (E-F)
            row.getCell('G').value = "-"; // Latitud (G-H)
            row.getCell('I').value = "-"; // Longitud (I-J)
            row.getCell('K').value = "-"; // RSS (K-L)
            row.getCell('M').value = "-"; // SSID (M)
        } else {
            row.getCell('B').value = numItem;
            row.getCell('C').value = valores[1];
            row.getCell('E').value = valores[2];
            row.getCell('G').value = valores[3];
            row.getCell('I').value = valores[4];
            row.getCell('K').value = valores[5];
            row.getCell('M').value = valores[6];
        }

        // --- ESTRUCTURA DE MERGES SEGURA ---
        const safeMerge = (r, c1, c2) => {
            const cell = hojaDestino.getCell(r, c1);
            if (!cell.isMerged) {
                try {
                    hojaDestino.mergeCells(r, c1, r, c2);
                } catch (e) {
                    console.warn(`Merge omitido en fila ${r}: ${e.message}`);
                }
            }
        };

        // Mantener la estructura visual siempre
        safeMerge(numFila, 3, 4);   // C-D
        safeMerge(numFila, 5, 6);   // E-F
        safeMerge(numFila, 7, 8);   // G-H
        safeMerge(numFila, 9, 10);  // I-J
        safeMerge(numFila, 11, 12); // K-L
    };
    let bloquesEscritos = 0;
    const primerColIdx = fPop.getCell(bloquesMatriz3_4[0].start).col;
    const primerIdIAO = fPop.getCell(primerColIdx + 1).value;
    if (esValorVacio(primerIdIAO)) {
        escribirFilaS3_4(filaActual, null, "-", false, true);
        return filaActual + 1;
    }
    bloquesMatriz3_4.forEach((bloque) => {
        const celdaBase = fPop.getCell(bloque.start);
        const idIAO = fPop.getCell(celdaBase.col + 1).value;
        if (esValorVacio(idIAO)) return;
        const colIdx = celdaBase.col;
        const datos = [];
        for (let i = 0; i < 7; i++) {
            datos.push(fPop.getCell(colIdx + i).value);
        }
        escribirFilaS3_4(filaActual, datos, bloquesEscritos + 1, bloquesEscritos > 0);
        filaActual++;
        bloquesEscritos++;
    });
    return filaActual;
}
module.exports = s3_4;