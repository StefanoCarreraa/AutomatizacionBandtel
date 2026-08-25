const { aplicarEstiloCelda, limpiarMergesFila } = require('../../utils/excelTools');

function s2_1(hojaDestino, fEquip) {
    // 1. BUSQUEDA DINÁMICA DEL INICIO
    let filaInicio = -1;
    hojaDestino.eachRow((row, rowNumber) => {
        const valorB = row.getCell('B').value;
        if (valorB && valorB.toString().includes("2.1. Radio Enlaces para el Backhaul")) {
            filaInicio = rowNumber + 2; 
        }
    });

    let filaActual = filaInicio !== -1 ? filaInicio : 21;
    if (!fEquip) return filaActual;

    const FILA_FIN_TABLA_ORIGINAL = filaActual;
    let itemCorrelativo = 1;

    const configuracion = [
        { c: 'G', start: 8 }, { c: 'AY', start: 52 }, { c: 'CQ', start: 96 },
        { c: 'EJ', start: 139 }, { c: 'GA', start: 184 }, { c: 'HS', start: 228 },
        { c: 'JK', start: 272 }, { c: 'LC', start: 316 }, { c: 'MU', start: 360 },
        { c: 'OM', start: 404 }, { c: 'QE', start: 448 }, { c: 'RW', start: 491 }
    ];

    configuracion.forEach((enlace) => {
        let colMatriz = enlace.start;
        for (let i = 0; i < 7; i++) {
            // --- BLINDAJE: Obtenemos el valor de forma segura ---
            const celdaRaw = fEquip.getCell(colMatriz);
            const valorD_Raw = celdaRaw.value;
            
            // Si el valor es NA, saltamos esta iteración
            if (valorD_Raw && valorD_Raw.toString().toUpperCase() === 'NA') {
                colMatriz += 4; continue;
            }

            // Convertimos a string de forma segura
            const valorD = valorD_Raw ? valorD_Raw.toString().trim() : "";

            if (filaActual > FILA_FIN_TABLA_ORIGINAL) {
                hojaDestino.insertRow(filaActual, {});
            }

            const row = hojaDestino.getRow(filaActual);
            row.height = 30;

            limpiarMergesFila(hojaDestino, filaActual);
            aplicarEstiloCelda(row);

            // --- LÓGICA DE VALIDACIÓN ---
            // Consideramos inválido si es vacío, "-", o "0"
            const esInvalido = valorD === "" || valorD === "-" || valorD === "0";

            if (esInvalido) {
                row.getCell('B').value = "-";
                row.getCell('C').value = "-";
                row.getCell('D').value = "-";
                row.getCell('G').value = "-";
                row.getCell('I').value = "-";
                row.getCell('K').value = "-";
                row.getCell('L').value = "-";
            } else {
                row.getCell('B').value = itemCorrelativo++; 
                row.getCell('C').value = fEquip.getCell(enlace.c).value || "-";
                row.getCell('D').value = valorD;
                row.getCell('G').value = fEquip.getCell(colMatriz + 1).value || "-";
                row.getCell('I').value = fEquip.getCell(colMatriz + 2).value || "-";
                row.getCell('K').value = fEquip.getCell(colMatriz + 3).value || "-";
                row.getCell('L').value = "-";
            }

            // 2. REALIZAR MERGES
            try {
                hojaDestino.mergeCells(filaActual, 4, filaActual, 6);  
                hojaDestino.mergeCells(filaActual, 7, filaActual, 8);  
                hojaDestino.mergeCells(filaActual, 9, filaActual, 10); 
                hojaDestino.mergeCells(filaActual, 12, filaActual, 13);
            } catch (e) { }

            // Alineación forzada
            const estiloFinal = { vertical: 'middle', horizontal: 'center', wrapText: true };
            for (let c = 2; c <= 13; c++) {
                hojaDestino.getCell(filaActual, c).alignment = estiloFinal;
            }

            filaActual++;
            colMatriz += 4;
        }
    });

    return filaActual;
}

module.exports = s2_1;