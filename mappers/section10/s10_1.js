const fs = require('fs');

async function insertarFotoFirmas(workbook, hojaDestino, rutaImagen1, rutaImagen2) {
    try {
        let filaFirmas = -1;
        hojaDestino.eachRow((row, rowNumber) => {
            const valorB = row.getCell('B').value;
            if (valorB && valorB.toString().includes("10. FIRMAS")) {
                filaFirmas = rowNumber;
            }
        });

        if (filaFirmas === -1) return;

        // --- INSERTAR IMAGEN 1 (Izquierda) ---
        if (fs.existsSync(rutaImagen1)) {
            const imageId1 = workbook.addImage({
                filename: rutaImagen1,
                extension: 'jpeg',
            });

            hojaDestino.addImage(imageId1, {
                tl: { col: 9.1, row: filaFirmas + 1.2 },
                br: { col: 11.7, row: filaFirmas + 4.9 },
                editAs: 'oneCell'
            });
        }

        // --- INSERTAR IMAGEN 2 (Derecha - Al lado) ---
        if (fs.existsSync(rutaImagen2)) {
            const imageId2 = workbook.addImage({
                filename: rutaImagen2,
                extension: 'jpeg',
            });

            hojaDestino.addImage(imageId2, {
                // Empezamos en la Columna M (12) para que esté al lado
                tl: { col: 11.9, row: filaFirmas + 2.2 },

                // Terminamos en la Columna O (14.9)
                br: { col: 12.2, row: filaFirmas + 4.9 },

                editAs: 'oneCell'
            });
        }

    } catch (e) {
        console.error("❌ Error al insertar las firmas:", e.message);
    }
}

module.exports = insertarFotoFirmas;