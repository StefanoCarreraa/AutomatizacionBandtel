// utils/excelTools.js
module.exports = {
    aplicarEstiloCelda: (row, cols = [2, 13]) => {
        for (let c = cols[0]; c <= cols[1]; c++) {
            const cell = row.getCell(c);
            cell.border = {
                top: { style: 'thin' }, left: { style: 'thin' },
                bottom: { style: 'thin' }, right: { style: 'thin' }
            };
            cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
            cell.font = { name: 'Arial', size: 11 };
        }
    },
    limpiarMergesFila: (hoja, fila) => {
        if (hoja._merges) {
            Object.keys(hoja._merges).forEach(key => {
                const m = hoja._merges[key];
                if (m.top === fila) delete hoja._merges[key];
            });
        }
    }
};