function procesarSeccion1(hojaDestino, fInfo, codigoTexto, proyecto) {
    if (!fInfo) return;

   // --- 1. Asignar el nombre del proyecto en C5 ---
    // Usamos el parámetro 'proyecto' directamente
    const nombreProyecto = proyecto ? proyecto.toUpperCase() : "SIN PROYECTO";

    // --- 2. Reemplazar '_______' por 'proyecto' en el título de B5 ---
    const celdaTitulo = hojaDestino.getCell('B5');
    let textoOriginal = celdaTitulo.value ? celdaTitulo.value.toString() : "";

    if (textoOriginal.includes('_______')) {
        // Reemplaza los guiones bajos por el nombre del proyecto
        celdaTitulo.value = textoOriginal.replace('_______', nombreProyecto);
    }

    // --- Lógica de Checkboxes (Tipos de Nodo) ---
    const esDistrital = fInfo.getCell('B').value?.toString().toUpperCase() === 'X';
    const esIntermedio = fInfo.getCell('C').value?.toString().toUpperCase() === 'X';
    const esTerminal = fInfo.getCell('D').value?.toString().toUpperCase() === 'X';
    const esOtro = fInfo.getCell('E').value?.toString().toUpperCase() === 'X';

    // Reset de celdas
    hojaDestino.getCell('B9').value = 'Nodo Distrital (     )';
    hojaDestino.getCell('E9').value = 'Nodo de Intermedio (     )';
    hojaDestino.getCell('H9').value = 'Nodo Terminal (     )';
    hojaDestino.getCell('K9').value = 'Otro (     )';

    if (esDistrital) hojaDestino.getCell('B9').value = 'Nodo Distrital (    X    )';
    else if (esIntermedio) hojaDestino.getCell('E9').value = 'Nodo de Intermedio (    X    )';
    else if (esTerminal) hojaDestino.getCell('H9').value = 'Nodo Terminal (    X    )';
    else if (esOtro) hojaDestino.getCell('K9').value = 'Otro (    X    )';

    // Formato de Checkboxes
    ['B9', 'E9', 'H9', 'K9'].forEach(celda => {
        const c = hojaDestino.getCell(celda);
        c.alignment = { horizontal: 'center', vertical: 'middle' };
        c.font = { name: 'Arial', size: 11, bold: true };
    });

    // --- Datos de Ubicación ---
    hojaDestino.getCell('D11').value = codigoTexto;
    hojaDestino.getCell('D10').value = fInfo.getCell('F').value;
    hojaDestino.getCell('K10').value = fInfo.getCell('M').value;
    hojaDestino.getCell('K11').value = fInfo.getCell('N').value;
    hojaDestino.getCell('D12').value = fInfo.getCell('H').value;
    hojaDestino.getCell('L12').value = fInfo.getCell('L').value;
    hojaDestino.getCell('D13').value = fInfo.getCell('I').value;
    hojaDestino.getCell('H13').value = fInfo.getCell('J').value;
    hojaDestino.getCell('L13').value = fInfo.getCell('K').value;

    // Sección 1.1
    hojaDestino.getCell('B16').value = fInfo.getCell('O').value;
}

module.exports = procesarSeccion1;