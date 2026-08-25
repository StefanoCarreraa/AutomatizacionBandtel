const ExcelJS = require('exceljs');
const path = require('path');

async function completarDatosYFecha(hojaDestino, codigoABuscar) {
    try {
        const rutaLibro1 = path.join(__dirname, '../..', 'input', 'Libro1.xlsx');
        const workbookOrigen = new ExcelJS.Workbook();
        await workbookOrigen.xlsx.readFile(rutaLibro1);
        const hojaOrigen = workbookOrigen.getWorksheet('MAPEO IP NODOS');

        if (!hojaOrigen) {
            console.error("❌ No se encontró la pestaña 'MAPEO IP NODOS'");
            return;
        }

        // 2. BUSCAR EL CÓDIGO EN LA COLUMNA I DEL LIBRO1
        let filaEncontrada = null;
        hojaOrigen.eachRow((row) => {
            const valorI = row.getCell('I').value?.toString().trim();
            if (valorI === codigoABuscar.toString().trim()) {
                filaEncontrada = row;
            }
        });

        if (!filaEncontrada) {
            console.warn(`⚠️ Código ${codigoABuscar} no encontrado en 'MAPEO IP NODOS' (Columna I).`);
            return;
        }

        // 3. EXTRAER DATOS (Columna F y Columna K)
        const valorF = filaEncontrada.getCell('F').value;
        const valorK = filaEncontrada.getCell('K').value;
        
        let dia, mesNombre, anio;

        if (valorK instanceof Date) {
            dia = valorK.getUTCDate();
            anio = valorK.getUTCFullYear();
            mesNombre = valorK.toLocaleString('es-ES', { month: 'long', timeZone: 'UTC' });
        } else {
            const partes = valorK.toString().split('/');
            dia = parseInt(partes[0], 10);
            const mesIndex = parseInt(partes[1], 10) - 1;
            anio = parseInt(partes[2], 10);
            
            const fechaTemporal = new Date(anio, mesIndex, dia);
            mesNombre = fechaTemporal.toLocaleString('es-ES', { month: 'long' });
        }

        // 4. POSICIONAR EN EL MODELO
        let filaBase = -1;
        hojaDestino.eachRow((row) => {
            const valorB = row.getCell('B').value;
            if (valorB && valorB.toString().includes("10. FIRMAS")) {
                filaBase = row.number;
            }
        });

        if (filaBase !== -1) {
            const filaObjetivo = filaBase + 9;
            
            // Pegar Columna F de Libro1 en Columna F del modelo
            hojaDestino.getCell(`F${filaObjetivo}`).value = valorF;

            // Pegar Fecha
            hojaDestino.getCell(`I${filaObjetivo}`).value = dia;
            hojaDestino.getCell(`K${filaObjetivo}`).value = mesNombre.charAt(0).toUpperCase() + mesNombre.slice(1);
            hojaDestino.getCell(`M${filaObjetivo}`).value = `del ${anio}`;
        }

    } catch (e) {
        console.error("❌ Error al procesar la búsqueda en Libro1:", e.message);
    }
}

module.exports = completarDatosYFecha;