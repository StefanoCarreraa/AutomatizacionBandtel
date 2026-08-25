const { aplicarEstiloCelda } = require('../../utils/excelTools');

function s3_2(hojaDestino, fPop, filaReferenciaAnterior) {
    let filaSeccion = -1;
    hojaDestino.eachRow((row, rowNumber) => {
        const valorB = row.getCell('B').value;
        if (valorB && valorB.toString().includes("3.2. Descripción de Antenas Sectoriales")) {
            filaSeccion = rowNumber;
        }
    });
    let filaActual = filaSeccion !== -1 ? filaSeccion + 2 : filaReferenciaAnterior + 4;
    if (!fPop) return filaActual;
    const FILA_BASE_PLANTILLA = filaActual;
    const limpiarMergesFila = (numFila) => {
        if (!hojaDestino._merges) return;

        Object.keys(hojaDestino._merges).forEach(key => {
            const m = hojaDestino._merges[key];
            // Si el merge intersecta con la fila, lo eliminamos físicamente
            if (m.top <= numFila && m.bottom >= numFila) {
                hojaDestino.unMergeCells(m.top, m.left, m.bottom, m.right);
                delete hojaDestino._merges[key];
            }
        });

        // LIMPIEZA ADICIONAL: Aseguramos que las celdas del rango de trabajo (2 a 13)
        // pierdan su estado de "combinadas" internamente
        const row = hojaDestino.getRow(numFila);
        for (let c = 2; c <= 13; c++) {
            const cell = row.getCell(c);
            // Esto desvincula a la celda de cualquier merge previo
            cell.master = cell;
        }
    };
    const mapeoBloques = [
        { id: 'N', det: 'AA' },
        { id: 'AH', det: 'AU' },
        { id: 'BB', det: 'BO' },
        { id: 'BV', det: 'CI' },
        { id: 'CP', det: 'DC' }
    ];
    const esValorVacio = (val) => {
        if (!val) return true;
        const v = val.toString().trim().toUpperCase();
        return v === 'NA' || v === '-' || v === 'N/A' || v === '';
    };
    const escribirFilaS3_2 = (numFila, valores, insertar, modoNoAplica = false) => {
        if (insertar || numFila > FILA_BASE_PLANTILLA) {
            hojaDestino.insertRow(numFila, {});
        }

        limpiarMergesFila(numFila);
        const row = hojaDestino.getRow(numFila);
        row.height = 30;

        // --- RESET DE CELDAS (B a M) ---
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
            // Ponemos "-" en cada celda que inicia un bloque de merge o columna simple
            row.getCell('B').value = "-"; // B
            row.getCell('C').value = "-"; // C-D
            row.getCell('E').value = "-"; // E
            row.getCell('F').value = "-"; // F
            row.getCell('G').value = "-"; // G
            row.getCell('H').value = "-"; // H
            row.getCell('I').value = "-"; // I
            row.getCell('J').value = "-"; // J-K
            row.getCell('L').value = "-"; // L-M
        } else {
            // Asignación normal de datos
            row.getCell('B').value = valores[0]; 
            row.getCell('C').value = valores[1]; 
            row.getCell('E').value = valores[2]; 
            row.getCell('F').value = valores[3]; 
            row.getCell('G').value = valores[4]; 
            row.getCell('H').value = valores[5]; 
            row.getCell('I').value = valores[6]; 
            row.getCell('J').value = valores[7]; 
            row.getCell('L').value = valores[8]; 
        }

        // --- MERGES CON ESTRUCTURA CONSTANTE ---
        const safeMerge = (r, c1, c2) => {
            try {
                const cell = hojaDestino.getCell(r, c1);
                if (!cell.isMerged) {
                    hojaDestino.mergeCells(r, c1, r, c2);
                }
            } catch (e) {
                console.warn(`Merge omitido en fila ${r}, col ${c1}-${c2}: ${e.message}`);
            }
        };

        // Aplicamos la misma estructura independientemente de modoNoAplica
        safeMerge(numFila, 3, 4);   // C-D
        safeMerge(numFila, 10, 11); // J-K
        safeMerge(numFila, 12, 13); // L-M
    };
    let bloquesEscritos = 0;
    const primerValorId = fPop.getCell(mapeoBloques[0].id).value;
    if (esValorVacio(primerValorId)) {
        escribirFilaS3_2(filaActual, null, false, true);
        return filaActual + 1;
    }
    mapeoBloques.forEach((bloque) => {
        const celdaId = fPop.getCell(bloque.id);
        const valorId = celdaId.value;
        if (esValorVacio(valorId)) return;
        const colIdIdx = celdaId.col;
        const colDetIdx = fPop.getCell(bloque.det).col;
        const datos = [
            fPop.getCell(colIdIdx).value,
            fPop.getCell(colIdIdx + 1).value,
            fPop.getCell(colDetIdx).value,
            fPop.getCell(colDetIdx + 1).value,
            fPop.getCell(colDetIdx + 2).value,
            fPop.getCell(colDetIdx + 3).value,
            fPop.getCell(colDetIdx + 4).value,
            fPop.getCell(colDetIdx + 5).value,
            fPop.getCell(colDetIdx + 6).value
        ];
        escribirFilaS3_2(filaActual, datos, bloquesEscritos > 0);
        filaActual++;
        bloquesEscritos++;
    });
    return filaActual;
}
module.exports = s3_2;