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
    const FILA_FIN_TABLA_ORIGINAL = filaActual;
    let itemCorrelativo = 1;

    const configuracion = [
        { c: 'G', start: 8 }, { c: 'AY', start: 52 }, { c: 'CQ', start: 96 },
        { c: 'EJ', start: 139 }, { c: 'GA', start: 184 }, { c: 'HS', start: 228 },
        { c: 'JK', start: 272 }, { c: 'LC', start: 316 }, { c: 'MU', start: 360 },
        { c: 'OM', start: 404 }, { c: 'QE', start: 448 }, { c: 'RW', start: 491 }
    ];

    const esVacio = (valor) => {
        if (valor === null || valor === undefined) return true;
        const texto = valor.toString().trim().toUpperCase();
        return texto === '' || texto === 'NA' || texto === '-' || texto === '0';
    };

    const registros = [];
    if (fEquip) {
        configuracion.forEach((enlace) => {
            for (let i = 0; i < 7; i++) {
                const colMatriz = enlace.start + i * 4;
                const valores = [
                    fEquip.getCell(enlace.c).value,
                    fEquip.getCell(colMatriz).value,
                    fEquip.getCell(colMatriz + 1).value,
                    fEquip.getCell(colMatriz + 2).value,
                    fEquip.getCell(colMatriz + 3).value
                ];
                const cantidadVacios = valores.filter(esVacio).length;

                if (cantidadVacios <= 3) {
                    registros.push({ valores });
                }
            }
        });
    }

    const escribirFila = (valores, numero = 'NA') => {
        if (filaActual > FILA_FIN_TABLA_ORIGINAL) {
            hojaDestino.insertRow(filaActual, {});
        }

        const row = hojaDestino.getRow(filaActual);
        row.height = 30;
        limpiarMergesFila(hojaDestino, filaActual);
        aplicarEstiloCelda(row);

        row.getCell('B').value = numero;
        row.getCell('C').value = valores[0] || '-';
        row.getCell('D').value = valores[1] || '-';
        row.getCell('G').value = valores[2] || '-';
        row.getCell('I').value = valores[3] || '-';
        row.getCell('K').value = valores[4] || '-';
        row.getCell('L').value = numero === 'NA' ? 'NA' : '-';

        try {
            hojaDestino.mergeCells(filaActual, 4, filaActual, 6);
            hojaDestino.mergeCells(filaActual, 7, filaActual, 8);
            hojaDestino.mergeCells(filaActual, 9, filaActual, 10);
            hojaDestino.mergeCells(filaActual, 12, filaActual, 13);
        } catch (e) { }

        const estiloFinal = { vertical: 'middle', horizontal: 'center', wrapText: true };
        for (let c = 2; c <= 13; c++) {
            hojaDestino.getCell(filaActual, c).alignment = estiloFinal;
        }

        filaActual++;
    };

    const primerRegistro = configuracion.length && fEquip
        ? [
            fEquip.getCell(configuracion[0].c).value,
            fEquip.getCell(configuracion[0].start).value,
            fEquip.getCell(configuracion[0].start + 1).value,
            fEquip.getCell(configuracion[0].start + 2).value,
            fEquip.getCell(configuracion[0].start + 3).value
        ]
        : null;
    const primerRegistroTieneExcesoDeVacios = !primerRegistro || primerRegistro.filter(esVacio).length > 3;

    if (registros.length === 0 || primerRegistroTieneExcesoDeVacios) {
        escribirFila(Array(5).fill('NA'));
    }

    registros.forEach(({ valores }) => {
        escribirFila(valores, itemCorrelativo++);
    });

    return filaActual;
}

module.exports = s2_1;