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
    const FILA_BASE_PLANTILLA = filaActual;
    const iniciosColumnas = ['UC', 'UJ', 'UQ', 'UX', 'VE', 'VL'];
    let itemCorrelativo = 1;

    const esVacio = (valor) => {
        if (valor === null || valor === undefined) return true;
        const texto = valor.toString().trim().toUpperCase();
        return texto === '' || texto === 'NA' || texto === '-' || texto === '0';
    };

    const escribirFila = (valores, esFilaNA = false) => {
        if (filaActual > FILA_BASE_PLANTILLA) {
            hojaDestino.insertRow(filaActual, {});
        }

        const row = hojaDestino.getRow(filaActual);
        limpiarMergesFila(hojaDestino, filaActual);
        aplicarEstiloCelda(row);

        row.getCell('B').value = esFilaNA ? 'NA' : itemCorrelativo++;
        row.getCell('C').value = valores[0] || (esFilaNA ? 'NA' : '-');
        row.getCell('E').value = valores[1] || (esFilaNA ? 'NA' : '-');
        row.getCell('F').value = valores[2] || (esFilaNA ? 'NA' : '-');
        row.getCell('G').value = valores[3] || (esFilaNA ? 'NA' : '-');
        row.getCell('I').value = valores[4] || (esFilaNA ? 'NA' : '-');
        row.getCell('K').value = valores[5] || (esFilaNA ? 'NA' : '-');
        row.getCell('M').value = valores[6] || (esFilaNA ? 'NA' : '-');

        try {
            hojaDestino.mergeCells(filaActual, 3, filaActual, 4);
            hojaDestino.mergeCells(filaActual, 7, filaActual, 8);
            hojaDestino.mergeCells(filaActual, 9, filaActual, 10);
            hojaDestino.mergeCells(filaActual, 11, filaActual, 12);
        } catch (e) { }

        if (esFilaNA) {
            for (let col = 2; col <= 13; col++) {
                row.getCell(col).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
            }
        }

        row.height = undefined;
        filaActual++;
    };

    const bloquesEvaluados = fEquip
        ? iniciosColumnas.map((letraInicio) => {
            const celdaBase = fEquip.getCell(letraInicio);
            const colIdx = celdaBase.col;
            const valores = Array.from({ length: 7 }, (_, index) => fEquip.getCell(colIdx + index).value);
            const cantidadVacios = valores.filter(esVacio).length;
            return { valores, cantidadVacios, tieneDatos: valores.some((valor) => !esVacio(valor)) };
        })
        : [];

    const bloquesConDatos = bloquesEvaluados.filter(({ cantidadVacios, tieneDatos }) =>
        tieneDatos && cantidadVacios <= 3
    );
    const primerBloqueTieneExcesoDeVacios = !bloquesEvaluados.length || bloquesEvaluados[0].cantidadVacios > 3;

    if (bloquesConDatos.length === 0 || primerBloqueTieneExcesoDeVacios) {
        escribirFila(Array(7).fill('NA'), true);
    }

    bloquesConDatos.forEach(({ valores }) => {
        escribirFila(valores);
    });

    return filaActual;
}

module.exports = s2_4;