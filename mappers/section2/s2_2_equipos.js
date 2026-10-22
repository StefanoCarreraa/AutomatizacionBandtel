const { aplicarEstiloCelda, limpiarMergesFila } = require('../../utils/excelTools');

function s2_2(hojaDestino, fEquip, filaReferenciaAnterior) {
    let filaSeccion = -1;
    hojaDestino.eachRow((row, rowNumber) => {
        const valorB = row.getCell('B').value;
        if (valorB && valorB.toString().includes("2.2. Equipo de Red / Ópticos")) {
            filaSeccion = rowNumber;
        }
    });

    let filaActual = filaSeccion !== -1 ? filaSeccion + 2 : filaReferenciaAnterior + 5;

    const estiloBold = { name: 'Arial', size: 11, bold: true };
    const bloquesEquipos = [
        { fijo: 'TM', datos: ['TM', 'TN', 'TO', 'TP'] },
        { fijo: 'TQ', datos: ['TQ', 'TR', 'TS', 'TT'] },
        { fijo: 'TU', datos: ['TU', 'TV', 'TW', 'TX'] },
        { fijo: 'TY', datos: ['TY', 'TZ', 'UA', 'UB'] }
    ];

    const FILA_BASE_PLANTILLA = filaActual;
    const normalizarDato = (cell) => {
        if (!cell || cell.value === null || cell.value === undefined) return 'NA';
        const texto = cell.text.trim();
        return texto === '' || texto.toUpperCase() === 'NA' || texto === '-' ? 'NA' : cell.value;
    };
    const tieneDato = (cell) => {
        if (!cell || cell.value === null || cell.value === undefined) return false;
        const texto = cell.text.trim();
        return texto !== '' && texto.toUpperCase() !== 'NA' && texto !== '-';
    };

    const escribirFila = (datos) => {
        if (filaActual > FILA_BASE_PLANTILLA) {
            hojaDestino.insertRow(filaActual, {});
        }

        const row = hojaDestino.getRow(filaActual);
        row.height = 25;
        limpiarMergesFila(hojaDestino, filaActual);
        aplicarEstiloCelda(row);

        const valores = [
            ['B', datos.equipo],
            ['E', datos.marca],
            ['G', datos.modelo],
            ['I', datos.serie],
            ['K', datos.comentarios]
        ];
        valores.forEach(([columna, valor]) => {
            const cell = row.getCell(columna);
            cell.value = valor;
            if (valor === 'NA') {
                cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
            }
        });
        row.getCell('B').font = estiloBold;
        if (datos.equipo !== 'NA') {
            row.getCell('B').alignment = { horizontal: 'left', vertical: 'middle' };
        }

        try {
            hojaDestino.mergeCells(filaActual, 2, filaActual, 4);
            hojaDestino.mergeCells(filaActual, 5, filaActual, 6);
            hojaDestino.mergeCells(filaActual, 7, filaActual, 8);
            hojaDestino.mergeCells(filaActual, 9, filaActual, 10);
            hojaDestino.mergeCells(filaActual, 11, filaActual, 13);
        } catch (e) { }

        filaActual++;
    };

    if (!fEquip) {
        escribirFila({
            equipo: 'NA',
            marca: 'NA',
            modelo: 'NA',
            serie: 'NA',
            comentarios: 'NA'
        });
        return filaActual;
    }

    const bloquesConDatos = bloquesEquipos.filter((bloque) =>
        bloque.datos.some((columna) => tieneDato(fEquip.getCell(columna)))
    );

    if (bloquesConDatos.length === 0) {
        escribirFila({
            equipo: 'NA',
            marca: 'NA',
            modelo: 'NA',
            serie: 'NA',
            comentarios: 'NA'
        });
        return filaActual;
    }

    bloquesConDatos.forEach((bloque) => {
        const celdaNombre = fEquip.worksheet.getRow(3).getCell(bloque.fijo);
        escribirFila({
            equipo: normalizarDato(celdaNombre),
            marca: normalizarDato(fEquip.getCell(bloque.datos[0])),
            modelo: normalizarDato(fEquip.getCell(bloque.datos[1])),
            serie: normalizarDato(fEquip.getCell(bloque.datos[2])),
            comentarios: normalizarDato(fEquip.getCell(bloque.datos[3]))
        });
    });

    return filaActual;
}

module.exports = s2_2;