const assert = require('node:assert/strict');
const { describe, it } = require('node:test');
const ExcelJS = require('exceljs');

const mapEnlacesBackhaul = require('../mappers/section2/s2_1_enlaces');
const mapEquiposRed = require('../mappers/section2/s2_2_equipos');
const mapDescripcionEnlaces = require('../mappers/section2/s2_3_detalle');
const mapEnlacesOpticos = require('../mappers/section2/s2_4_tarjetas');

const columnasEnlacesBackhaul = [
    { tipo: 'G', inicio: 8 },
    { tipo: 'AY', inicio: 52 },
    { tipo: 'CQ', inicio: 96 },
    { tipo: 'EJ', inicio: 139 },
    { tipo: 'GA', inicio: 184 },
    { tipo: 'HS', inicio: 228 },
    { tipo: 'JK', inicio: 272 },
    { tipo: 'LC', inicio: 316 },
    { tipo: 'MU', inicio: 360 },
    { tipo: 'OM', inicio: 404 },
    { tipo: 'QE', inicio: 448 },
    { tipo: 'RW', inicio: 491 }
];

const iniciosEnlacesDetalle = [
    'AM', 'CE', 'DW', 'FO', 'HG', 'IY', 'KQ', 'MI', 'OA', 'PS', 'RK', 'TC'
];

const iniciosEnlacesOpticos = ['UC', 'UJ', 'UQ', 'UX', 'VE', 'VL'];

function createTemplate(sectionTitle, row) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Acta');
    worksheet.getCell(`B${row}`).value = sectionTitle;
    return worksheet;
}

function createEquipmentRow(defaultValue = 'NA') {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Matriz');

    for (const { tipo, inicio } of columnasEnlacesBackhaul) {
        worksheet.getCell(`${tipo}5`).value = defaultValue;
        for (let enlace = 0; enlace < 7; enlace++) {
            for (let offset = 0; offset < 4; offset++) {
                worksheet.getRow(5).getCell(inicio + enlace * 4 + offset).value = defaultValue;
            }
        }
    }

    for (const inicio of iniciosEnlacesDetalle) {
        const base = worksheet.getCell(`${inicio}5`).col;
        for (let offset = 0; offset < 10; offset++) {
            worksheet.getRow(5).getCell(base + offset).value = defaultValue;
        }
    }

    for (const inicio of iniciosEnlacesOpticos) {
        const base = worksheet.getCell(`${inicio}5`).col;
        for (let offset = 0; offset < 7; offset++) {
            worksheet.getRow(5).getCell(base + offset).value = defaultValue;
        }
    }

    return worksheet.getRow(5);
}

function getValues(worksheet, row, columns) {
    return columns.map((column) => worksheet.getCell(`${column}${row}`).value);
}

describe('mapeadores de la sección 2', () => {
    describe('2.1 Radio Enlaces para el Backhaul', () => {
        it('escribe una sola fila NA si todos los campos están vacíos o son NA', () => {
            const template = createTemplate('2.1. Radio Enlaces para el Backhaul', 19);
            const row = createEquipmentRow('');

            mapEnlacesBackhaul(template, row);

            assert.deepEqual(getValues(template, 21, ['B', 'C', 'D', 'G', 'I', 'K', 'L']), [
                'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA'
            ]);
            assert.equal(template.getCell('B22').value, null);
        });

        it('coloca una fila NA antes de enlaces válidos si el primer enlace tiene más de tres campos vacíos', () => {
            const template = createTemplate('2.1. Radio Enlaces para el Backhaul', 19);
            const row = createEquipmentRow('NA');

            row.getCell('G').value = '1+0';
            row.getCell('L').value = 'RADIO';
            row.getCell('M').value = 'Huawei';
            row.getCell('N').value = 'Modelo-X';
            row.getCell('O').value = 'Serie-2';

            mapEnlacesBackhaul(template, row);

            assert.deepEqual(getValues(template, 21, ['B', 'C', 'D', 'G', 'I', 'K', 'L']), [
                'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA'
            ]);
            assert.deepEqual(getValues(template, 22, ['B', 'C', 'D', 'G', 'I', 'K', 'L']), [
                1, '1+0', 'RADIO', 'Huawei', 'Modelo-X', 'Serie-2', '-'
            ]);
        });

        it('acepta un enlace con exactamente tres campos vacíos', () => {
            const template = createTemplate('2.1. Radio Enlaces para el Backhaul', 19);
            const row = createEquipmentRow('NA');
            row.getCell('G').value = '1+0';
            row.getCell('H').value = 'RADIO';
            row.getCell('I').value = 'Huawei';
            row.getCell('J').value = 'NA';
            row.getCell('K').value = 'Serie-1';

            mapEnlacesBackhaul(template, row);

            assert.deepEqual(getValues(template, 21, ['B', 'C', 'D', 'G', 'I', 'K']), [
                1, '1+0', 'RADIO', 'Huawei', 'NA', 'Serie-1'
            ]);
        });
    });

    describe('2.2 Equipo de Red / Ópticos', () => {
        it('escribe una sola fila NA cuando no hay datos y centra los campos', () => {
            const template = createTemplate('2.2. Equipo de Red / Ópticos', 25);
            const row = createEquipmentRow('NA');

            mapEquiposRed(template, row, 22);

            assert.deepEqual(getValues(template, 27, ['B', 'E', 'G', 'I', 'K']), [
                'NA', 'NA', 'NA', 'NA', 'NA'
            ]);
            for (const column of ['B', 'E', 'G', 'I', 'K']) {
                assert.equal(template.getCell(`${column}27`).alignment.horizontal, 'center');
                assert.equal(template.getCell(`${column}27`).alignment.vertical, 'middle');
            }
            assert.equal(template.getCell('B28').value, null);
        });

        it('usa los nombres de la fila 3 y convierte campos vacíos o NA a NA', () => {
            const template = createTemplate('2.2. Equipo de Red / Ópticos', 25);
            const row = createEquipmentRow('NA');
            const worksheet = row.worksheet;
            worksheet.getCell('TM3').value = 'ROUTER';
            worksheet.getCell('TM5').value = 'HUAWEI';
            worksheet.getCell('TN5').value = '';
            worksheet.getCell('TO5').value = 'SERIE-1';
            worksheet.getCell('TP5').value = 'NA';

            mapEquiposRed(template, row, 22);

            assert.deepEqual(getValues(template, 27, ['B', 'E', 'G', 'I', 'K']), [
                'ROUTER', 'HUAWEI', 'NA', 'SERIE-1', 'NA'
            ]);
        });
    });

    describe('2.3 Descripción de Enlaces', () => {
        it('escribe una sola fila NA si toda la sección está vacía o es NA', () => {
            for (const initialValue of ['', 'NA']) {
                const template = createTemplate('2.3. Descripción de Enlaces', 25);
                const row = createEquipmentRow(initialValue);

                mapDescripcionEnlaces(template, row, 24);

                assert.deepEqual(
                    Array.from({ length: 12 }, (_, index) => template.getRow(27).getCell(index + 2).value),
                    Array(12).fill('NA')
                );
                assert.equal(template.getCell('B28').value, null);
            }
        });

        it('coloca NA al inicio y conserva bloques posteriores válidos', () => {
            const template = createTemplate('2.3. Descripción de Enlaces', 25);
            const row = createEquipmentRow('NA');
            const firstStart = row.getCell('AM').col;
            const nextStart = row.getCell('CE').col;

            for (let index = 0; index < 10; index++) {
                row.getCell(firstStart + index).value = index < 4 ? 'NA' : `DATO-${index}`;
                row.getCell(nextStart + index).value = `VALOR-${index}`;
            }

            mapDescripcionEnlaces(template, row, 24);

            assert.equal(template.getCell('B27').value, 'NA');
            assert.equal(template.getCell('B28').value, 1);
            assert.equal(template.getCell('D28').value, 'VALOR-0');
        });

        it('acepta tres campos vacíos y omite un bloque con cuatro', () => {
            const template = createTemplate('2.3. Descripción de Enlaces', 25);
            const row = createEquipmentRow('NA');
            const firstStart = row.getCell('AM').col;
            const nextStart = row.getCell('CE').col;

            for (let index = 0; index < 10; index++) {
                row.getCell(firstStart + index).value = index < 3 ? 'NA' : `DATO-${index}`;
                row.getCell(nextStart + index).value = index < 4 ? 'NA' : `DATO-${index}`;
            }

            mapDescripcionEnlaces(template, row, 24);

            assert.equal(template.getCell('B27').value, 1);
            assert.equal(template.getCell('D27').value, 'NA');
            assert.equal(template.getCell('B28').value, null);
        });
    });

    describe('2.4 Descripción de Enlaces Ópticos', () => {
        it('escribe una sola fila NA si toda la sección está vacía o es NA', () => {
            for (const initialValue of ['', 'NA']) {
                const template = createTemplate('2.4. Descripción de Enlaces (Ópticos)', 33);
                const row = createEquipmentRow(initialValue);

                mapEnlacesOpticos(template, row, 32);

                assert.deepEqual(getValues(template, 35, ['B', 'C', 'E', 'F', 'G', 'I', 'K', 'M']), [
                    'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA'
                ]);
                assert.equal(template.getCell('B36').value, null);
            }
        });

        it('coloca NA antes de bloques válidos si el primer bloque supera tres campos vacíos', () => {
            const template = createTemplate('2.4. Descripción de Enlaces (Ópticos)', 33);
            const row = createEquipmentRow('NA');
            const firstStart = row.getCell('UC').col;
            const nextStart = row.getCell('UJ').col;

            for (let index = 0; index < 7; index++) {
                row.getCell(firstStart + index).value = index < 4 ? 'NA' : `DATO-${index}`;
                row.getCell(nextStart + index).value = `VALOR-${index}`;
            }

            mapEnlacesOpticos(template, row, 32);

            assert.deepEqual(getValues(template, 35, ['B', 'C', 'E', 'F', 'G', 'I', 'K', 'M']), [
                'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA'
            ]);
            assert.equal(template.getCell('B36').value, 1);
            assert.equal(template.getCell('C36').value, 'VALOR-0');
        });

        it('acepta tres campos vacíos y omite los bloques posteriores que superan el límite', () => {
            const template = createTemplate('2.4. Descripción de Enlaces (Ópticos)', 33);
            const row = createEquipmentRow('NA');
            const firstStart = row.getCell('UC').col;
            const nextStart = row.getCell('UJ').col;

            for (let index = 0; index < 7; index++) {
                row.getCell(firstStart + index).value = index < 3 ? 'NA' : `DATO-${index}`;
                row.getCell(nextStart + index).value = index < 4 ? 'NA' : `DATO-${index}`;
            }

            mapEnlacesOpticos(template, row, 32);

            assert.equal(template.getCell('B35').value, 1);
            assert.equal(template.getCell('C35').value, 'NA');
            assert.equal(template.getCell('B36').value, null);
        });
    });
});
