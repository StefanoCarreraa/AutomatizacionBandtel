const { aplicarEstiloCelda } = require('../../utils/excelTools');

function s3_1(hojaDestino, fPop, filaInicioSugerida) {
    if (!fPop) return filaInicioSugerida;
    let filaTitulo = -1;
    hojaDestino.eachRow((row, rowNumber) => {
        const valorB = row.getCell('B').value;
        if (valorB && valorB.toString().includes("3.1. Radio Enlaces Punto a Multipunto")) {
            filaTitulo = rowNumber;
        }
    });
    let filaActual = filaTitulo !== -1 ? (filaTitulo + 2) :
        (typeof filaInicioSugerida === 'number' && !isNaN(filaInicioSugerida) ? filaInicioSugerida : 42);
    const FILA_BASE_TABLA = filaActual;
    const bloquesMatriz = [
        { start: 'P' }, { start: 'AJ' }, { start: 'BD' }, { start: 'BX' }, { start: 'CR' }
    ];
    const esValorVacio = (val) => {
        if (!val) return true;
        const v = val.toString().trim().toUpperCase();
        return v === 'NA' || v === '-' || v === 'N/A' || v === '';
    };
    const limpiarMergesFila = (numFila) => {
        if (!hojaDestino._merges) return;
        const llaves = Object.keys(hojaDestino._merges);
        llaves.forEach(key => {
            const m = hojaDestino._merges[key];
            if (m.top <= numFila && m.bottom >= numFila) {
                hojaDestino.unMergeCells(m.top, m.left, m.bottom, m.right);
                delete hojaDestino._merges[key];
            }
        });
    };
    const escribirFilaS3 = (numFila, valores, forzarInsercion, modoNoAplica = false) => {
        if (forzarInsercion || numFila > FILA_BASE_TABLA) {
            hojaDestino.insertRow(numFila, {});
        }
        limpiarMergesFila(numFila);
        const row = hojaDestino.getRow(numFila);
        row.height = 25;
        for (let c = 2; c <= 13; c++) {
            const cell = row.getCell(c);
            cell.value = null;
            cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
            cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        }
        if (modoNoAplica) {
            row.getCell(2).value = "-";  // Bloque Equipos (B)
            row.getCell(5).value = "-";  // Bloque Cantidad (E)
            row.getCell(6).value = "-";  // Bloque Marca (F)
            row.getCell(8).value = "-";  // Bloque Modelo (H)
            row.getCell(9).value = "-";  // Bloque Serie (I)
            row.getCell(11).value = "-"; // Bloque Características (M)
        } else {
            row.getCell(2).value = valores[0];  // Equipos (B)
            row.getCell(5).value = valores[1];  // Cantidad (E)
            row.getCell(6).value = valores[2];  // Marca (F)
            row.getCell(8).value = valores[3];  // Modelo (H)
            row.getCell(9).value = valores[4];  // Serie (I)
            row.getCell(11).value = valores[5]; // Características (M)
        }
        const safeMerge = (r, c1, c2) => {
            const cell = hojaDestino.getCell(r, c1);
            if (!cell.isMerged) {
                try {
                    hojaDestino.mergeCells(r, c1, r, c2);
                } catch (e) {
                    console.warn(`Omitido merge fila ${r}, col ${c1}-${c2}: ${e.message}`);
                }
            }
        };

        // Ejecutamos la estructura de la tabla
        safeMerge(numFila, 2, 4);   // Equipos: B(2) a D(4)
        // E(5) queda sola (Cantidad)
        safeMerge(numFila, 6, 7);   // Marca: F(6) a G(7)
        // H(8) queda sola (Modelo)
        safeMerge(numFila, 9, 10);  // Serie: I(9) a J(10)
        safeMerge(numFila, 11, 13); // Características: K(11) a M(13)
    };
    let bloquesEscritos = 0;
    const primerColBase = fPop.getCell(bloquesMatriz[0].start).col;
    const primerValorValido = fPop.getCell(primerColBase).value;
    if (esValorVacio(primerValorValido)) {
        escribirFilaS3(filaActual, null, false, true);
        return filaActual + 1;
    }
    bloquesMatriz.forEach((bloque) => {
        const colBase = fPop.getCell(bloque.start).col;
        const valorRef = fPop.getCell(colBase).value;
        if (esValorVacio(valorRef)) return;
        const datosA = [
            fPop.getCell(colBase).value,
            fPop.getCell(colBase + 1).value,
            fPop.getCell(colBase + 2).value,
            fPop.getCell(colBase + 3).value,
            fPop.getCell(colBase + 4).value,
            "NA"
        ];
        escribirFilaS3(filaActual, datosA, bloquesEscritos > 0);
        filaActual++;
        const datosB = [
            fPop.getCell(colBase + 5).value,
            fPop.getCell(colBase + 6).value,
            fPop.getCell(colBase + 7).value,
            fPop.getCell(colBase + 8).value,
            fPop.getCell(colBase + 9).value,
            fPop.getCell(colBase + 10).value
        ];
        escribirFilaS3(filaActual, datosB, true);
        filaActual++;
        bloquesEscritos++;
    });
    return filaActual;
}
module.exports = s3_1;