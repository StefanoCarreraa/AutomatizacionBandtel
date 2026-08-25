const { aplicarEstiloCelda } = require('../../utils/excelTools');

function s3_5(hojaDestino, fPop, filaInicio) {
    if (!fPop) return filaInicio;
    let filaSeccion = -1;
    hojaDestino.eachRow((row, rowNumber) => {
        const valorB = row.getCell('B').value;
        if (valorB && valorB.toString().includes("3.5. Radio de cobertura del POP")) {
            filaSeccion = rowNumber;
        }
    });
    let filaActual = (filaSeccion !== -1) ? (filaSeccion + 2) : (filaInicio + 4);
    if (isNaN(filaActual)) filaActual = 75;
    const colInicio = fPop.getCell('GS').col;
    const colFin = fPop.getCell('NW').col;
    const bloquesMatriz3_5 = [];
    for (let c = colInicio; c <= colFin; c += 6) {
        bloquesMatriz3_5.push({ colIdx: c });
    }
    const esValorVacio = (val) => {
        if (!val) return true;
        const v = val.toString().trim().toUpperCase();
        return v === 'NA' || v === '-' || v === 'N/A' || v === '';
    };
    const limpiarMergesFila = (numFila) => {
        // 1. Limpieza física: forzar unmerge por rango exacto
        try {
            hojaDestino.unMergeCells(numFila, 2, numFila, 13);
        } catch (e) { }

        // 2. Limpieza de metadatos de celda (evita herencia de 'master')
        const row = hojaDestino.getRow(numFila);
        for (let c = 2; c <= 13; c++) {
            const cell = row.getCell(c);
            cell.master = cell;
        }

        // 3. ¡IMPORTANTE! Limpiar el índice interno de la hoja de forma manual
        // Esto previene que el objeto interno de merges tenga referencias viejas
        if (hojaDestino._merges) {
            Object.keys(hojaDestino._merges).forEach(key => {
                const m = hojaDestino._merges[key];
                // Si el merge intersecta la fila, lo eliminamos del mapa de la hoja
                if (m.top === numFila || m.bottom === numFila) {
                    delete hojaDestino._merges[key];
                }
            });
        }
    };
    const escribirFilaS3_5 = (numFila, valores, numItem, insertar, modoNoAplica = false) => {
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
            // Repetimos "-" en las celdas iniciales de cada grupo/columna
            row.getCell('B').value = "-"; // N°
            row.getCell('C').value = "-"; // Punto (C-E)
            row.getCell('F').value = "-"; // Localidad (F-G)
            row.getCell('H').value = "-"; // Latitud (H)
            row.getCell('I').value = "-"; // Longitud (I-J)
            row.getCell('K').value = "-"; // RSS (K-L)
            row.getCell('M').value = "-"; // SSID (M)
        } else {
            row.getCell('B').value = numItem;
            row.getCell('C').value = valores[0];
            row.getCell('F').value = valores[1];
            row.getCell('H').value = valores[2];
            row.getCell('I').value = valores[3];
            row.getCell('K').value = valores[4];
            row.getCell('M').value = valores[5];
        }

        // --- ESTRUCTURA DE MERGES SEGURA ---
        const safeMerge = (r, c1, c2) => {
            const cell = hojaDestino.getCell(r, c1);
            if (!cell.isMerged) {
                try {
                    hojaDestino.mergeCells(r, c1, r, c2);
                } catch (e) {
                    console.warn(`Merge omitido en fila ${r}, cols ${c1}-${c2}: ${e.message}`);
                }
            }
        };

        // Mantener la estructura visual idéntica siempre
        safeMerge(numFila, 3, 5);   // Punto: C-E
        safeMerge(numFila, 6, 7);   // Localidad: F-G
        safeMerge(numFila, 9, 10);  // Longitud: I-J
        safeMerge(numFila, 11, 12); // RSS: K-L
    };
    let bloquesEscritos = 0;
    const primerValorRef = fPop.getCell(bloquesMatriz3_5[0].colIdx).value;
    if (esValorVacio(primerValorRef)) {
        escribirFilaS3_5(filaActual, null, "-", false, true);
        return filaActual + 1;
    }
    bloquesMatriz3_5.forEach((bloque) => {
        const valorRef = fPop.getCell(bloque.colIdx).value;
        if (esValorVacio(valorRef)) return;
        const datos = [];
        for (let i = 0; i < 6; i++) {
            datos.push(fPop.getCell(bloque.colIdx + i).value);
        }
        escribirFilaS3_5(filaActual, datos, bloquesEscritos + 1, bloquesEscritos > 0);
        filaActual++;
        bloquesEscritos++;
    });
    return filaActual;
}
module.exports = s3_5;