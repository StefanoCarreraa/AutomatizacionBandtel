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
    if (!fEquip) return filaActual;

    const estiloBold = { name: 'Arial', size: 11, bold: true };
    const bloquesEquipos = [
        { fijo: 'TM', datos: ['TM', 'TN', 'TO', 'TP'] },
        { fijo: 'TQ', datos: ['TQ', 'TR', 'TS', 'TT'] },
        { fijo: 'TU', datos: ['TU', 'TV', 'TW', 'TX'] },
        { fijo: 'TY', datos: ['TY', 'TZ', 'UA', 'UB'] }
    ];

    const FILA_BASE_PLANTILLA = filaActual;

    bloquesEquipos.forEach((bloque, index) => {
        const celdaNombre = fEquip.getCell(bloque.fijo);
        const celdaPrimerDato = fEquip.getCell(bloque.datos[0]);
        
        // --- BLINDAJE: Conversión segura a String ---
        const valorNombreRaw = celdaNombre.value ? celdaNombre.value.toString().trim() : "";
        const valorDatoRaw = celdaPrimerDato.value ? celdaPrimerDato.value.toString().trim() : "";

        // Si es NA, lo ignoramos totalmente
        if (valorNombreRaw.toUpperCase() === 'NA') return;

        // Si el nombre está vacío y no es el primer equipo, no saturamos el acta
        if (valorNombreRaw === "" && index > 0) return;

        // Insertar fila física
        if (filaActual > FILA_BASE_PLANTILLA) {
            hojaDestino.insertRow(filaActual, {});
        }

        const row = hojaDestino.getRow(filaActual);
        row.height = 25;
        limpiarMergesFila(hojaDestino, filaActual);
        aplicarEstiloCelda(row);

        // Lógica de datos inválidos (o vacíos)
        const esInvalido = valorDatoRaw === "" || 
                           valorDatoRaw === "-" || 
                           valorDatoRaw === "0" || 
                           valorDatoRaw.toUpperCase() === "NA";

        // Asignar Nombre del equipo (si está vacío, ponemos guion para el formato)
        row.getCell('B').value = valorNombreRaw || "-";
        row.getCell('B').font = estiloBold;
        row.getCell('B').alignment = { horizontal: 'left', vertical: 'middle' };

        if (esInvalido) {
            row.getCell('E').value = "-"; 
            row.getCell('G').value = "-"; 
            row.getCell('I').value = "-"; 
            row.getCell('K').value = "-"; 
        } else {
            row.getCell('E').value = celdaPrimerDato.value || "-"; 
            row.getCell('G').value = fEquip.getCell(bloque.datos[1]).value || "-"; 
            row.getCell('I').value = fEquip.getCell(bloque.datos[2]).value || "-"; 
            row.getCell('K').value = fEquip.getCell(bloque.datos[3]).value || "-"; 
        }

        try {
            hojaDestino.mergeCells(filaActual, 2, filaActual, 4);   
            hojaDestino.mergeCells(filaActual, 5, filaActual, 6);   
            hojaDestino.mergeCells(filaActual, 7, filaActual, 8);   
            hojaDestino.mergeCells(filaActual, 9, filaActual, 10);  
            hojaDestino.mergeCells(filaActual, 11, filaActual, 13); 
        } catch (e) { }

        filaActual++;
    });

    return filaActual;
}

module.exports = s2_2;