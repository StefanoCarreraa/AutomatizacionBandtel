const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { describe, it } = require('node:test');
const ExcelJS = require('exceljs');

const mapSection1 = require('../mappers/section1/section1');
const mapSection3_1 = require('../mappers/section3/s3_1_radios');
const mapSection3_2 = require('../mappers/section3/s3_2_antenas');
const mapSection3_3 = require('../mappers/section3/s3_3_pop');
const mapSection3_4 = require('../mappers/section3/s3_4_instituciones');
const mapSection3_5 = require('../mappers/section3/s3_5_radiocobertura');
const mapSection4_1 = require('../mappers/section4/s4_1_rectificador');
const mapSection4_2 = require('../mappers/section4/s4_2_baterias');
const mapSection4_3 = require('../mappers/section4/s4_3_climatizacion');
const mapSection5 = require('../mappers/section5/s5_1_seguridad');
const mapSection6_1 = require('../mappers/section6/s6_1');
const mapSection6_2 = require('../mappers/section6/s6_2');
const mapSection6_3 = require('../mappers/section6/s6_3');
const mapSection6_4 = require('../mappers/section6/s6_4');
const mapSection6_5 = require('../mappers/section6/s6_5');
const mapSection7_1 = require('../mappers/section7/s7_1');
const mapSection7_2 = require('../mappers/section7/s7_2');
const insertSignatures = require('../mappers/section10/s10_1');
const completeSignatureData = require('../mappers/section10/s10_2');

function createWorkbookWithSheet() {
    const workbook = new ExcelJS.Workbook();
    return { workbook, worksheet: workbook.addWorksheet('Acta') };
}

function setSectionTitle(worksheet, row, title) {
    worksheet.getCell(`B${row}`).value = title;
}

describe('mapeadores de las secciones 1 y 3 a 7', () => {
    it('sección 1: asigna código, ubicación, proyecto y tipo de nodo', () => {
        const { worksheet } = createWorkbookWithSheet();
        worksheet.getCell('B5').value = 'ACTA DE _______';
        const sourceWorkbook = new ExcelJS.Workbook();
        const source = sourceWorkbook.addWorksheet('Info').getRow(5);
        source.getCell('B').value = 'X';
        source.getCell('F').value = 'Provincia';
        source.getCell('H').value = 'Distrito';
        source.getCell('I').value = 'Localidad';
        source.getCell('J').value = 'Latitud';
        source.getCell('K').value = 'Longitud';
        source.getCell('L').value = 'Altitud';
        source.getCell('M').value = 'Región';
        source.getCell('N').value = 'Departamento';
        source.getCell('O').value = 'Dirección';

        mapSection1(worksheet, source, 'PA-0001-AD', 'Pasco');

        assert.equal(worksheet.getCell('B5').value, 'ACTA DE PASCO');
        assert.match(worksheet.getCell('B9').value, /X/);
        assert.equal(worksheet.getCell('D10').value, 'Provincia');
        assert.equal(worksheet.getCell('D11').value, 'PA-0001-AD');
        assert.equal(worksheet.getCell('K10').value, 'Región');
        assert.equal(worksheet.getCell('B16').value, 'Dirección');
    });

    it('sección 1: no cambia la plantilla cuando no hay fila de origen', () => {
        const { worksheet } = createWorkbookWithSheet();
        worksheet.getCell('B5').value = 'ACTA DE _______';

        mapSection1(worksheet, null, 'PA-0001-AD', 'Pasco');

        assert.equal(worksheet.getCell('B5').value, 'ACTA DE _______');
    });

    it('sección 3.1: escribe las dos filas del primer bloque multipunto', () => {
        const { worksheet } = createWorkbookWithSheet();
        setSectionTitle(worksheet, 5, '3.1. Radio Enlaces Punto a Multipunto');
        const sourceWorkbook = new ExcelJS.Workbook();
        const source = sourceWorkbook.addWorksheet('POP').getRow(5);
        ['Equipo A', 2, 'Marca A', 'Modelo A', 'Serie A', 'Equipo B', 3, 'Marca B', 'Modelo B', 'Serie B', 'Detalle B']
            .forEach((value, index) => { source.getCell(16 + index).value = value; });

        const nextRow = mapSection3_1(worksheet, source);

        assert.equal(worksheet.getCell('B7').value, 'Equipo A');
        assert.equal(worksheet.getCell('E7').value, 2);
        assert.equal(worksheet.getCell('F7').value, 'Marca A');
        assert.equal(worksheet.getCell('B8').value, 'Equipo B');
        assert.equal(worksheet.getCell('K8').value, 'Detalle B');
        assert.equal(nextRow, 9);
    });

    it('sección 3.1: escribe una fila no aplica cuando no hay equipos', () => {
        const { worksheet } = createWorkbookWithSheet();
        const sourceWorkbook = new ExcelJS.Workbook();
        const source = sourceWorkbook.addWorksheet('POP').getRow(5);

        mapSection3_1(worksheet, source, 42);

        assert.equal(worksheet.getCell('B42').value, '-');
        assert.equal(worksheet.getCell('E42').value, '-');
    });

    it('sección 3.2: mapea el identificador y los detalles de antena', () => {
        const { worksheet } = createWorkbookWithSheet();
        setSectionTitle(worksheet, 5, '3.2. Descripción de Antenas Sectoriales');
        const sourceWorkbook = new ExcelJS.Workbook();
        const source = sourceWorkbook.addWorksheet('POP').getRow(5);
        source.getCell('N').value = 'ANT-1';
        source.getCell('O').value = 'Sector Norte';
        source.getCell('AA').value = 'Modelo';
        source.getCell('AB').value = 'Frecuencia';

        mapSection3_2(worksheet, source);

        assert.equal(worksheet.getCell('B7').value, 'ANT-1');
        assert.equal(worksheet.getCell('C7').value, 'Sector Norte');
        assert.equal(worksheet.getCell('E7').value, 'Modelo');
        assert.equal(worksheet.getCell('F7').value, 'Frecuencia');
    });

    it('sección 3.3: mapea localidad y campos geográficos en sus columnas', () => {
        const { worksheet } = createWorkbookWithSheet();
        setSectionTitle(worksheet, 5, '3.3. Localidades atendidas por el POP');
        const sourceWorkbook = new ExcelJS.Workbook();
        const source = sourceWorkbook.addWorksheet('POP').getRow(5);
        ['Tipo', 'Ubigeo', 'Localidad', 'Distrito', 'Provincia', 'SSID']
            .forEach((value, index) => { source.getCell(114 + index).value = value; });

        mapSection3_3(worksheet, source, 40);

        assert.equal(worksheet.getCell('B7').value, 1);
        assert.equal(worksheet.getCell('C7').value, 'Ubigeo');
        assert.equal(worksheet.getCell('E7').value, 'Localidad');
        assert.equal(worksheet.getCell('K7').value, 'Tipo');
        assert.equal(worksheet.getCell('M7').value, 'SSID');
    });

    it('sección 3.3: genera una fila no aplica cuando la primera localidad está vacía', () => {
        const { worksheet } = createWorkbookWithSheet();
        const sourceWorkbook = new ExcelJS.Workbook();
        const source = sourceWorkbook.addWorksheet('POP').getRow(5);

        mapSection3_3(worksheet, source, 40);

        assert.equal(worksheet.getCell('B43').value, '-');
        assert.equal(worksheet.getCell('M43').value, '-');
    });

    it('sección 3.4: mapea una institución cuando hay identificador', () => {
        const { worksheet } = createWorkbookWithSheet();
        setSectionTitle(worksheet, 5, '3.4. Instituciones Abonadas Obligatorias');
        const sourceWorkbook = new ExcelJS.Workbook();
        const source = sourceWorkbook.addWorksheet('POP').getRow(5);
        ['Tipo', 'ID-IAO', 'Localidad', 'Latitud', 'Longitud', 'RSS', 'SSID']
            .forEach((value, index) => { source.getCell(138 + index).value = value; });

        mapSection3_4(worksheet, source, 40);

        assert.equal(worksheet.getCell('B7').value, 1);
        assert.equal(worksheet.getCell('C7').value, 'ID-IAO');
        assert.equal(worksheet.getCell('E7').value, 'Localidad');
        assert.equal(worksheet.getCell('M7').value, 'SSID');
    });

    it('sección 3.4: genera una fila no aplica si el primer identificador está vacío', () => {
        const { worksheet } = createWorkbookWithSheet();
        const sourceWorkbook = new ExcelJS.Workbook();
        const source = sourceWorkbook.addWorksheet('POP').getRow(5);

        mapSection3_4(worksheet, source, 40);

        assert.equal(worksheet.getCell('B44').value, '-');
        assert.equal(worksheet.getCell('M44').value, '-');
    });

    it('sección 3.5: mapea un registro de cobertura del POP', () => {
        const { worksheet } = createWorkbookWithSheet();
        setSectionTitle(worksheet, 5, '3.5. Radio de cobertura del POP');
        const sourceWorkbook = new ExcelJS.Workbook();
        const source = sourceWorkbook.addWorksheet('POP').getRow(5);
        ['Punto', 'Localidad', 'Latitud', 'Longitud', 'RSS', 'SSID']
            .forEach((value, index) => { source.getCell(201 + index).value = value; });

        mapSection3_5(worksheet, source, 40);

        assert.equal(worksheet.getCell('B7').value, 1);
        assert.equal(worksheet.getCell('C7').value, 'Punto');
        assert.equal(worksheet.getCell('F7').value, 'Localidad');
        assert.equal(worksheet.getCell('M7').value, 'SSID');
    });

    it('sección 3.5: genera una fila no aplica cuando la cobertura está vacía', () => {
        const { worksheet } = createWorkbookWithSheet();
        const sourceWorkbook = new ExcelJS.Workbook();
        const source = sourceWorkbook.addWorksheet('POP').getRow(5);

        mapSection3_5(worksheet, source, 40);

        assert.equal(worksheet.getCell('B44').value, '-');
        assert.equal(worksheet.getCell('M44').value, '-');
    });

    it('sección 4.1: mapea equipos rectificadores y omite los NA', () => {
        const { worksheet } = createWorkbookWithSheet();
        setSectionTitle(worksheet, 5, '4.1. Sistema Rectificador');
        const sourceWorkbook = new ExcelJS.Workbook();
        const source = sourceWorkbook.addWorksheet('Energía').getRow(5);
        ['RECTIFICADOR', 'Marca', 'Modelo', 'Serie', 'Capacidad', 'Temperatura']
            .forEach((value, index) => { source.getCell(5 + index).value = value; });
        source.getCell('K').value = 'NA';

        const nextRow = mapSection4_1(worksheet, source, 1);

        assert.deepEqual(getValues(worksheet, 7, ['B', 'D', 'F', 'H', 'J', 'L']), [
            'RECTIFICADOR', 'Marca', 'Modelo', 'Serie', 'Capacidad', 'Temperatura'
        ]);
        assert.equal(nextRow, 8);
    });

    it('sección 4.2: mapea un banco de baterías', () => {
        const { worksheet } = createWorkbookWithSheet();
        setSectionTitle(worksheet, 5, '4.2. Banco de Baterias');
        const sourceWorkbook = new ExcelJS.Workbook();
        const source = sourceWorkbook.addWorksheet('Energía').getRow(5);
        ['BANCO 1', 'Marca', 'Modelo', 'Capacidad', 'Tensión', 'Temperatura']
            .forEach((value, index) => { source.getCell(83 + index).value = value; });

        const nextRow = mapSection4_2(worksheet, source, 1);

        assert.deepEqual(getValues(worksheet, 7, ['B', 'D', 'F', 'H', 'J', 'L']), [
            'BANCO 1', 'Marca', 'Modelo', 'Capacidad', 'Tensión', 'Temperatura'
        ]);
        assert.equal(nextRow, 8);
    });

    it('sección 4.3: mapea los equipos de climatización', () => {
        const { worksheet } = createWorkbookWithSheet();
        setSectionTitle(worksheet, 5, '4.3. Sistema de Climatización');
        const sourceWorkbook = new ExcelJS.Workbook();
        const source = sourceWorkbook.addWorksheet('Energía').getRow(5);
        ['AIRE ACONDICIONADO', 'Marca', 'Modelo', 'Serie', 'Capacidad', 'Seteo']
            .forEach((value, index) => { source.getCell(167 + index).value = value; });

        const nextRow = mapSection4_3(worksheet, source, 1);

        assert.deepEqual(getValues(worksheet, 7, ['B', 'D', 'F', 'H', 'J', 'L']), [
            'AIRE ACONDICIONADO', 'Marca', 'Modelo', 'Serie', 'Capacidad', 'Seteo'
        ]);
        assert.equal(nextRow, 8);
    });

    it('sección 5: mapea equipos de seguridad y omite bloques no instalados', () => {
        const { worksheet } = createWorkbookWithSheet();
        setSectionTitle(worksheet, 5, '5.1. Sistema de Seguridad');
        const sourceWorkbook = new ExcelJS.Workbook();
        const source = sourceWorkbook.addWorksheet('Seguridad').getRow(5);
        ['CÁMARA', 'Marca', 'Modelo', 'Serie', 'Comentario']
            .forEach((value, index) => { source.getCell(12 + index).value = value; });
        source.getCell('Q').value = 'NA';

        const nextRow = mapSection5(worksheet, source, 1);

        assert.deepEqual(getValues(worksheet, 7, ['B', 'E', 'G', 'I', 'K']), [
            'CÁMARA', 'Marca', 'Modelo', 'Serie', 'Comentario'
        ]);
        assert.equal(nextRow, 8);
    });

    it('sección 6.1: mapea el cerco y formatea las respuestas SI/NO', () => {
        const { worksheet } = createWorkbookWithSheet();
        setSectionTitle(worksheet, 5, '6.1. Cerco Perimétrico');
        const sourceWorkbook = new ExcelJS.Workbook();
        const source = sourceWorkbook.addWorksheet('Obras').getRow(5);
        ['Malla', 'Área', 'Altura', 'SI', 'NO', 'Dimensiones']
            .forEach((value, index) => { source.getCell(5 + index).value = value; });

        const result = mapSection6_1(worksheet, source, 1);

        assert.deepEqual(getValues(worksheet, 7, ['B', 'D', 'F', 'H', 'J', 'L']), [
            'Malla', 'Área', 'Altura', 'SI (  X  )  NO (     )',
            'SI (     )  NO (  X  )', 'Dimensiones'
        ]);
        assert.equal(result, 7);
    });

    it('sección 6.2: mapea la torre y sus verificaciones', () => {
        const { worksheet } = createWorkbookWithSheet();
        setSectionTitle(worksheet, 5, '6.2. Descripción de Torre');
        const sourceWorkbook = new ExcelJS.Workbook();
        const source = sourceWorkbook.addWorksheet('Obras').getRow(5);
        ['30 m', 'SI', 'NO', 'X', '1', 'VERDADERO']
            .forEach((value, index) => { source.getCell(11 + index).value = value; });

        mapSection6_2(worksheet, source, 1);

        assert.equal(worksheet.getCell('B7').value, '30 m');
        assert.equal(worksheet.getCell('D7').value, 'SI (  X  )  NO (     )');
        assert.equal(worksheet.getCell('F7').value, 'SI (     )  NO (  X  )');
        assert.equal(worksheet.getCell('L7').value, 'SI (  X  )  NO (     )');
    });

    it('sección 6.3: escribe las seis filas del patio y normaliza casillas', () => {
        const { worksheet } = createWorkbookWithSheet();
        setSectionTitle(worksheet, 5, '6.3. Patio');
        const sourceWorkbook = new ExcelJS.Workbook();
        const source = sourceWorkbook.addWorksheet('Obras').getRow(5);
        for (let index = 0; index < 6; index++) {
            const start = 17 + index * 4;
            source.getCell(start).value = index === 0 ? 'SI' : 'NO';
            source.getCell(start + 2).value = `Unidad ${index}`;
            source.getCell(start + 3).value = `Comentario ${index}`;
        }

        const nextRow = mapSection6_3(worksheet, source, 1);

        assert.equal(worksheet.getCell('E7').value, 'SI (  X  )  NO (     )');
        assert.equal(worksheet.getCell('J7').value, 'Unidad 0');
        assert.equal(worksheet.getCell('K7').value, 'Comentario 0');
        assert.equal(worksheet.getCell('E12').value, 'SI (     )  NO (  X  )');
        assert.equal(nextRow, 13);
    });

    it('sección 6.4: mapea el shelter y sus verificaciones', () => {
        const { worksheet } = createWorkbookWithSheet();
        setSectionTitle(worksheet, 5, '6.4. Shelter');
        const sourceWorkbook = new ExcelJS.Workbook();
        const source = sourceWorkbook.addWorksheet('Obras').getRow(5);
        ['Material', 'Dimensiones', 'SI', 'NO', 'X', '1', 'VERDADERO']
            .forEach((value, index) => { source.getCell(41 + index).value = value; });

        mapSection6_4(worksheet, source, 1);

        assert.equal(worksheet.getCell('D6').value, 'Material');
        assert.equal(worksheet.getCell('J6').value, 'Dimensiones');
        assert.equal(worksheet.getCell('B8').value, 'SI (  X  )  NO (     )');
        assert.equal(worksheet.getCell('D8').value, 'SI (     )  NO (  X  )');
        assert.equal(worksheet.getCell('L8').value, 'SI (  X  )  NO (     )');
    });

    it('sección 6.5: mapea suministro eléctrico y formatea casillas', () => {
        const { worksheet } = createWorkbookWithSheet();
        setSectionTitle(worksheet, 5, '6.5. Suministro de Energía');
        const sourceWorkbook = new ExcelJS.Workbook();
        const source = sourceWorkbook.addWorksheet('Obras').getRow(5);
        source.getCell('AV').value = 'SI';
        source.getCell('AW').value = 'NO';
        source.getCell('AX').value = 'Monofásico';
        source.getCell('AY').value = '220 V';

        const lastRow = mapSection6_5(worksheet, source, 1);

        assert.match(worksheet.getCell('H6').value, /^SI \(  X  \)/);
        assert.equal(worksheet.getCell('B8').value, 'SI (     )  NO (  X  )');
        assert.equal(worksheet.getCell('F8').value, 'Monofásico');
        assert.equal(worksheet.getCell('J8').value, '220 V');
        assert.equal(lastRow, 8);
    });

    it('sección 7.1: mapea cuatro respuestas SI/NO debajo del título', () => {
        const { worksheet } = createWorkbookWithSheet();
        setSectionTitle(worksheet, 5, '7.1.  Carácteristicas Generales');
        const sourceWorkbook = new ExcelJS.Workbook();
        const source = sourceWorkbook.addWorksheet('PAT').getRow(5);
        source.getCell('E').value = 'SI';
        source.getCell('F').value = 'NO';
        source.getCell('G').value = 'X';
        source.getCell('H').value = 'SI';

        mapSection7_1(worksheet, source);

        assert.equal(worksheet.getCell('F7').value, 'SI (  X  )  NO (     )');
        assert.equal(worksheet.getCell('L7').value, 'SI (     )  NO (  X  )');
        assert.equal(worksheet.getCell('F8').value, 'SI (  X  )  NO (     )');
        assert.equal(worksheet.getCell('L8').value, 'SI (  X  )  NO (     )');
    });

    it('sección 7.2: mapea tipo de electrodo y mediciones', () => {
        const { worksheet } = createWorkbookWithSheet();
        setSectionTitle(worksheet, 5, '7.2.  Mediciones (*)');
        const sourceWorkbook = new ExcelJS.Workbook();
        const source = sourceWorkbook.addWorksheet('PAT').getRow(5);
        source.getCell('I').value = 'SI';
        source.getCell('K').value = 12.1;
        source.getCell('L').value = 13.2;
        source.getCell('M').value = 14.3;
        source.getCell('N').value = 'NO';

        const lastRow = mapSection7_2(worksheet, source);

        assert.match(worksheet.getCell('B6').value, /Vertical \(  X  \)/);
        assert.equal(worksheet.getCell('D8').value, 12.1);
        assert.equal(worksheet.getCell('F8').value, 13.2);
        assert.equal(worksheet.getCell('H8').value, 14.3);
        assert.equal(worksheet.getCell('J8').value, 'SI (     )  NO (  X  )');
        assert.equal(lastRow, 8);
    });
});

describe('mapeadores de la sección 10', () => {
    it('10.1: inserta ambas imágenes en el libro y las ancla a la sección de firmas', async () => {
        const image1 = path.resolve(__dirname, '../input/Imagen1.png');
        const image2 = path.resolve(__dirname, '../input/Imagen2.png');
        assert.equal(fs.existsSync(image1), true);
        assert.equal(fs.existsSync(image2), true);

        const { workbook, worksheet } = createWorkbookWithSheet();
        setSectionTitle(worksheet, 10, '10. FIRMAS');

        await insertSignatures(workbook, worksheet, image1, image2);

        assert.equal(workbook.media.length, 2);
        assert.equal(worksheet.getImages().length, 2);
    });

    it('10.1: no inserta imágenes si falta el encabezado de firmas', async () => {
        const image1 = path.resolve(__dirname, '../input/Imagen1.png');
        const image2 = path.resolve(__dirname, '../input/Imagen2.png');
        const { workbook, worksheet } = createWorkbookWithSheet();

        await insertSignatures(workbook, worksheet, image1, image2);

        assert.equal(workbook.media.length, 0);
    });

    it('10.2: completa los datos de firma usando un código y fecha del libro de mapeo', async () => {
        const sourcePath = path.resolve(__dirname, '../input/Libro1.xlsx');
        assert.equal(fs.existsSync(sourcePath), true);

        const sourceWorkbook = new ExcelJS.Workbook();
        await sourceWorkbook.xlsx.readFile(sourcePath);
        const sourceSheet = sourceWorkbook.getWorksheet('MAPEO IP NODOS');
        assert.ok(sourceSheet, 'el libro de entrada debe incluir MAPEO IP NODOS');

        let sourceRow;
        sourceSheet.eachRow((row) => {
            if (!sourceRow && row.getCell('I').value && row.getCell('K').value instanceof Date) {
                sourceRow = row;
            }
        });
        assert.ok(sourceRow, 'el libro de entrada debe tener un código con fecha válida');

        const code = sourceRow.getCell('I').value.toString().trim();
        const date = sourceRow.getCell('K').value;
        const { worksheet } = createWorkbookWithSheet();
        setSectionTitle(worksheet, 10, '10. FIRMAS');

        await completeSignatureData(worksheet, code);

        const targetRow = 19;
        const month = date.toLocaleString('es-ES', { month: 'long', timeZone: 'UTC' });
        assert.equal(worksheet.getCell(`F${targetRow}`).value, sourceRow.getCell('F').value);
        assert.equal(worksheet.getCell(`I${targetRow}`).value, date.getUTCDate());
        assert.equal(worksheet.getCell(`K${targetRow}`).value, month.charAt(0).toUpperCase() + month.slice(1));
        assert.equal(worksheet.getCell(`M${targetRow}`).value, `del ${date.getUTCFullYear()}`);
    });
});

describe('secciones pendientes de implementación', () => {
    it.todo('sección 8: agregar pruebas funcionales cuando se implementen sus mapeadores');
    it.todo('sección 9: agregar pruebas funcionales cuando se implemente su mapeador');
});

function getValues(worksheet, row, columns) {
    return columns.map((column) => worksheet.getCell(`${column}${row}`).value);
}
