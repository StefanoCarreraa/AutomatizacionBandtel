const { aplicarEstiloCelda, limpiarMergesFila } = require('../../utils/excelTools');

function s2_3(hojaDestino, fEquip, filaReferenciaAnterior) {
    let filaSeccion = -1;
    hojaDestino.eachRow((row, rowNumber) => {
        const valorB = row.getCell('B').value;
        if (valorB && valorB.toString().includes("2.3. Descripción de Enlaces")) {
            filaSeccion = rowNumber;
        }
    });

    let filaActual = filaSeccion !== -1 ? filaSeccion + 2 : filaReferenciaAnterior + 3;
    if (!fEquip) return filaActual;

    let itemCorrelativo = 1;
    const FILA_BASE_PLANTILLA = filaActual;

    const CONFIG_BLOQUES = [
        { labelCols: ['AK', 'AL'], dataStart: 'AM' },
        { labelCols: ['CC', 'CD'], dataStart: 'CE' },
        { labelCols: ['DU', 'DV'], dataStart: 'DW' },
        { labelCols: ['FM', 'FN'], dataStart: 'FO' },
        { labelCols: ['HE', 'HF'], dataStart: 'HG' },
        { labelCols: ['IW', 'IX'], dataStart: 'IY' },
        { labelCols: ['KO', 'KP'], dataStart: 'KQ' },
        { labelCols: ['MG', 'MH'], dataStart: 'MI' },
        { labelCols: ['NY', 'NZ'], dataStart: 'OA' },
        { labelCols: ['PQ', 'PR'], dataStart: 'PS' },
        { labelCols: ['RI', 'RJ'], dataStart: 'RK' },
        { labelCols: ['TA', 'TB'], dataStart: 'TC' }
    ];

    CONFIG_BLOQUES.forEach((bloque, index) => {
        const celdaBase = fEquip.getCell(bloque.dataStart);
        // Tratamiento seguro contra null
        const valorRefRaw = celdaBase.value ? celdaBase.value.toString().trim() : "";
        const valorUpper = valorRefRaw.toUpperCase();

        // 1. FILTRO NA: Ignoramos totalmente si es NA
        if (valorUpper === 'NA') return;

        // 2. FILTRO DE FILAS VACÍAS:
        // Si no hay datos y NO es el primer bloque, no creamos fila (evita saturar de guiones)
        if (valorRefRaw === "" && index > 0) return;

        // Inserción de fila física
        if (filaActual > FILA_BASE_PLANTILLA) {
            hojaDestino.insertRow(filaActual, {});
        }

        const row = hojaDestino.getRow(filaActual);
        limpiarMergesFila(hojaDestino, filaActual);
        aplicarEstiloCelda(row);

        const colBase = celdaBase.col;

        // --- LÓGICA DE VALIDACIÓN ---
        const esInvalido = valorRefRaw === "" || valorRefRaw === "-" || valorRefRaw === "0";

        if (esInvalido) {
            row.getCell('B').value = "-";
            row.getCell('C').value = "-";
            for (let i = 0; i < 10; i++) {
                row.getCell(4 + i).value = "-";
            }
        } else {
            row.getCell('B').value = itemCorrelativo++;
            const label1 = fEquip.getCell(bloque.labelCols[0]).value || '';
            const label2 = fEquip.getCell(bloque.labelCols[1]).value || '';
            row.getCell('C').value = `${label1} / ${label2}`;
            
            // Llenado de columnas D a M (4 a 13)
            for (let i = 0; i < 10; i++) {
                const val = fEquip.getCell(colBase + i).value;
                row.getCell(4 + i).value = (val !== null && val !== undefined) ? val : "-";
            }
        }

        row.height = undefined;
        filaActual++;
    });

    return filaActual;
}

module.exports = s2_3;