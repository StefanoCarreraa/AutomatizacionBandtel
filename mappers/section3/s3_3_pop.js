const { aplicarEstiloCelda } = require('../../utils/excelTools');

function s3_3(hojaDestino, fPop, filaInicio) {
    if (!fPop) return filaInicio;
    let filaSeccion = -1;
    hojaDestino.eachRow((row, rowNumber) => {
        const valorB = row.getCell('B').value;
        if (valorB && valorB.toString().includes("3.3. Localidades atendidas por el POP")) {
            filaSeccion = rowNumber;
        }
    });
    let filaActual = (filaSeccion !== -1) ? (filaSeccion + 2) : (filaInicio + 3);
    if (isNaN(filaActual)) filaActual = 50;
    const bloquesMatriz3_3 = [
        { start: 'DJ' }, { start: 'DP' }, { start: 'DV' }, { start: 'EB' }
    ];
    const esValorVacio = (val) => {
        if (!val) return true;
        const v = val.toString().trim().toUpperCase();
        return v === 'NA' || v === '-' || v === 'N/A' || v === '';
    };
   const limpiarMergesFila = (numFila) => {
        // 1. Forzar desvinculación por rango físico (B a M)
        // Usamos un try-catch porque unMergeCells lanza error si el rango no está combinado
        try {
            hojaDestino.unMergeCells(numFila, 2, numFila, 13);
        } catch (e) {
            // Silenciamos: si no había merge, no pasa nada
        }

        // 2. Limpieza de metadatos internos de exceljs
        const row = hojaDestino.getRow(numFila);
        for (let c = 2; c <= 13; c++) {
            const cell = row.getCell(c);
            // IMPORTANTE: Esto rompe la relación de "Celda Maestra" 
            // que causa el error "Already merged"
            cell.master = cell; 
        }
        
        // 3. Limpieza manual del índice (opcional pero recomendado)
        if (hojaDestino._merges) {
            Object.keys(hojaDestino._merges).forEach(key => {
                const m = hojaDestino._merges[key];
                if (m.top === numFila || m.bottom === numFila) {
                    delete hojaDestino._merges[key];
                }
            });
        }
    };
    const escribirFilaS3_3 = (numFila, valores, numItem, insertar, modoNoAplica = false) => {
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
            cell.font = { name: 'Arial', size: 10 };
            cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        }
        if (modoNoAplica) {
            row.getCell('B').value = "-"; // N°
            row.getCell('C').value = "-"; // Ubigeo (C-D)
            row.getCell('E').value = "-"; // Localidad (E-F)
            row.getCell('G').value = "-"; // Distrito (G-H)
            row.getCell('I').value = "-"; // Provincia (I-J)
            row.getCell('K').value = "-"; // Tipo Localidad (K-L)
            row.getCell('M').value = "-"; // SSID (M)
        } else {
            row.getCell('B').value = numItem;
            row.getCell('C').value = valores[1];
            row.getCell('E').value = valores[2];
            row.getCell('G').value = valores[3];
            row.getCell('I').value = valores[4];
            row.getCell('K').value = valores[0];
            row.getCell('M').value = valores[5];
        }
        const safeMerge = (r, c1, c2) => {
            const cell = hojaDestino.getCell(r, c1);
            if (!cell.isMerged) {
                try {
                    hojaDestino.mergeCells(r, c1, r, c2);
                } catch (e) {
                    console.warn(`Merge fallido en fila ${r}: ${e.message}`);
                }
            }
        };
        safeMerge(numFila, 3, 4);   // Ubigeo: C-D
        safeMerge(numFila, 5, 6);   // Localidad: E-F
        safeMerge(numFila, 7, 8);   // Distrito: G-H
        safeMerge(numFila, 9, 10);  // Provincia: I-J
        safeMerge(numFila, 11, 12); // Tipo Localidad: K-L
    };
    let bloquesEscritos = 0;
    const primerValorReferencia = fPop.getCell(bloquesMatriz3_3[0].start).value;
    if (esValorVacio(primerValorReferencia)) {
        escribirFilaS3_3(filaActual, null, "-", false, true);
        return filaActual + 1;
    }
    bloquesMatriz3_3.forEach((bloque) => {
        const celdaBase = fPop.getCell(bloque.start);
        const valorRef = celdaBase.value;
        if (esValorVacio(valorRef)) return;
        const colIdx = celdaBase.col;
        const datos = [];
        for (let i = 0; i < 6; i++) {
            datos.push(fPop.getCell(colIdx + i).value);
        }
        escribirFilaS3_3(filaActual, datos, bloquesEscritos + 1, bloquesEscritos > 0);
        filaActual++;
        bloquesEscritos++;
    });
    return filaActual;
}
module.exports = s3_3;