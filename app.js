const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');
const readline = require('readline/promises');
const { stdin: input, stdout: output } = require('process');

// --- IMPORTACIÓN DE MÓDULOS ---
const s1_info = require('./mappers/section1/section1');
const s2_1 = require('./mappers/section2/s2_1_enlaces');
const s2_2 = require('./mappers/section2/s2_2_equipos');
const s2_3 = require('./mappers/section2/s2_3_detalle');
const s2_4 = require('./mappers/section2/s2_4_tarjetas');
const s3_1 = require('./mappers/section3/s3_1_radios');
const s3_2 = require('./mappers/section3/s3_2_antenas');
const s3_3 = require('./mappers/section3/s3_3_pop');
const s3_4 = require('./mappers/section3/s3_4_instituciones');
const s3_5 = require('./mappers/section3/s3_5_radiocobertura');
const s4_1 = require('./mappers/section4/s4_1_rectificador');
const s4_2 = require('./mappers/section4/s4_2_baterias');
const s4_3 = require('./mappers/section4/s4_3_climatizacion');
const s5_1 = require('./mappers/section5/s5_1_seguridad');
const s6_1 = require('./mappers/section6/s6_1');
const s6_2 = require('./mappers/section6/s6_2');
const s6_3 = require('./mappers/section6/s6_3');
const s6_4 = require('./mappers/section6/s6_4');
const s6_5 = require('./mappers/section6/s6_5');
const s7_1 = require('./mappers/section7/s7_1');
const s7_2 = require('./mappers/section7/s7_2');
const insertarFotoFirmas = require('./mappers/section10/s10_1');
const completarDatosYFecha = require('./mappers/section10/s10_2');

async function ejecutarAutomatizacion() {
    const rl = readline.createInterface({ input, output });

    try {
        console.log('--- Iniciando App - AutomatizacionBandtel ---');

        // --- PREGUNTA INICIAL ---
        const opcion = await rl.question(
            '\n¿Qué regiones deseas procesar?\n' +
            '  [1] Solo Pasco\n' +
            '  [2] Solo Huánuco\n' +
            '  [3] Ambos\n' +
            'Selecciona una opción (1, 2 o 3): '
        );
        rl.close();

        const opcionTrim = opcion.trim();
        if (!['1', '2', '3'].includes(opcionTrim)) {
            console.error('❌ Opción inválida. Ejecuta el script nuevamente y selecciona 1, 2 o 3.');
            return;
        }

        const libroLista = new ExcelJS.Workbook();
        const libroMatrizPasco = new ExcelJS.Workbook();
        const libroMatrizHuanuco = new ExcelJS.Workbook();

        await libroLista.xlsx.readFile('input/Lista.xlsx');

        if (opcionTrim === '1' || opcionTrim === '3') {
            await libroMatrizPasco.xlsx.readFile('input/Tabulado Acta de Nodo_PR-Pasco8.xlsx');
        }
        if (opcionTrim === '2' || opcionTrim === '3') {
            await libroMatrizHuanuco.xlsx.readFile('input/Tabulado Acta de Nodo_PR-Huánuco5.xlsx');
        }

        const hojaLista = libroLista.getWorksheet('Hoja1');

        // --- 1. DEFINICIÓN DE PESTAÑAS Y CREACIÓN DE ÍNDICES (MAPAS) ---
        const nombresHojas = {
            info: '1. INFORMACIÓN DEL NODO',
            equip: '2. EQUIPAMIENTO DE TRANSMISIÓN',
            pop: '3. PUNTO DE PRESENCIA - POP',
            energia: '4. SISTEMAS DE ENERGÍA & CLIMAT',
            seguridad: '5. SEGURIDAD FISICA',
            obras: '6. OBRAS CIVILES',
            puestaTierra: '7. SISTEMA PUESTA A TIERRA',
            fotos: '8. REPORTE FOTOGRÁFICO',
            comentarios: '9. COMENTARIOS GENERALES',
            firmas: '10. FIRMAS'
        };

        const indices = {
            info: new Map(), equip: new Map(), pop: new Map(), energia: new Map(),
            seguridad: new Map(), obras: new Map(), puestaTierra: new Map(),
            fotos: new Map(), comentarios: new Map(), firmas: new Map()
        };

        const mapeoColumnas = {
            equip: 'C', pop: 'C', energia: 'C', seguridad: 'B',
            obras: 'C', puestaTierra: 'C', fotos: 'C', comentarios: 'C', firmas: 'C'
        };

        // Función para indexar un libro específico
        const indexarLibro = (libro, regionLabel) => {
            console.log(`⏳ Indexando matriz de ${regionLabel}...`);

            // 1. Indexar Info
            const hojaInfo = libro.getWorksheet(nombresHojas.info);
            if (hojaInfo) {
                hojaInfo.eachRow(row => {
                    const cod = row.getCell('G').value?.toString().trim();
                    if (cod) indices.info.set(cod, row);
                });
            }

            // 2. Indexar el resto de pestañas
            for (const key in mapeoColumnas) {
                const hoja = libro.getWorksheet(nombresHojas[key]);
                if (hoja) {
                    hoja.eachRow(row => {
                        const cod = row.getCell(mapeoColumnas[key]).value?.toString().trim();
                        if (cod) indices[key].set(cod, row);
                    });
                }
            }
        };

        // Ejecutar indexación según la selección
        if (opcionTrim === '1' || opcionTrim === '3') {
            indexarLibro(libroMatrizPasco, 'Pasco');
        }
        if (opcionTrim === '2' || opcionTrim === '3') {
            indexarLibro(libroMatrizHuanuco, 'Huánuco');
        }

        const { default: pLimit } = await import('p-limit');
        const limite = pLimit(100);
        const tareas = [];
        let activos = 0;

        // --- 2. PROCESAMIENTO DE LA LISTA ---
        for (let i = 1; ; i++) {
            const rutaImagen1 = 'C:\\DEV\\Github\\AutomatizacionBandtel\\input\\Imagen1.png';
            const rutaImagen2 = 'C:\\DEV\\Github\\AutomatizacionBandtel\\input\\Imagen2.png';
            const codigoABuscar = hojaLista.getCell(`A${i}`).value;

            if (!codigoABuscar) break;

            const codigoTexto = codigoABuscar.toString().trim();

            let proyecto = "";
            if (codigoTexto.startsWith("PA")) {
                proyecto = "PASCO";
            } else if (codigoTexto.startsWith("HC")) {
                proyecto = "HUANUCO";
            } else {
                console.warn(`⚠️ Código ${codigoTexto} no tiene un prefijo reconocido (PA/HC).`);
                continue;
            }

            // Filtrar según la opción elegida por el usuario
            if (opcionTrim === '1' && proyecto !== "PASCO") continue;
            if (opcionTrim === '2' && proyecto !== "HUANUCO") continue;

            const fInfo = indices.info.get(codigoTexto);
            const fEquip = indices.equip.get(codigoTexto);
            const fPop = indices.pop.get(codigoTexto);
            const fEnergia = indices.energia.get(codigoTexto);
            const fSeguridad = indices.seguridad.get(codigoTexto);
            const fObras = indices.obras.get(codigoTexto);
            const fPuestaTierra = indices.puestaTierra.get(codigoTexto);

            tareas.push(limite(async () => {
                activos++;
                console.log(`--- Activos actualmente: ${activos} ---`);
                try {
                    console.log(`🔍 Procesando: ${codigoTexto} | Región: ${proyecto}`);
                    if (fInfo) {
                        const libroModelo = new ExcelJS.Workbook();
                        await libroModelo.xlsx.readFile('input/Modelo Acta de Nodo_Rev_BANDTEL.xlsx');
                        const hoja = libroModelo.getWorksheet(1);

                        // --- LLAMADAS MODULARES ---
                        s1_info(hoja, indices.info.get(codigoTexto), codigoTexto, proyecto);
                        s2_1(hoja, indices.equip.get(codigoTexto));
                        s2_2(hoja, fEquip);
                        s2_3(hoja, fEquip);
                        s2_4(hoja, fEquip);
                        // sección 3
                        s3_1(hoja, fPop);
                        s3_2(hoja, fPop);
                        s3_3(hoja, fPop);
                        s3_4(hoja, fPop);
                        s3_5(hoja, fPop);
                        // sección 4
                        s4_1(hoja, fEnergia);
                        s4_2(hoja, fEnergia);
                        s4_3(hoja, fEnergia);
                        // sección 5
                        s5_1(hoja, fSeguridad);
                        // sección 6
                        s6_1(hoja, fObras);
                        s6_2(hoja, fObras);
                        s6_3(hoja, fObras);
                        s6_4(hoja, fObras);
                        s6_5(hoja, fObras);
                        // sección 7
                        s7_1(hoja, fPuestaTierra);
                        s7_2(hoja, fPuestaTierra);

                        await new Promise(resolve => setTimeout(resolve, 500));
                        await insertarFotoFirmas(libroModelo, hoja, rutaImagen1, rutaImagen2);
                        await completarDatosYFecha(hoja, codigoTexto);

                        // --- GESTIÓN DE CARPETAS Y GUARDADO ---
                        const carpetaProyecto = path.join(__dirname, 'output', proyecto);

                        if (!fs.existsSync(carpetaProyecto)) {
                            fs.mkdirSync(carpetaProyecto, { recursive: true });
                        }

                        const nombreBase = `Acta_${codigoTexto}`;
                        const rutaExcel = path.join(carpetaProyecto, `${nombreBase}.xlsx`);

                        await libroModelo.xlsx.writeFile(rutaExcel);
                        console.log(`   ✅ ${codigoTexto} Excel generado en [${proyecto}]: ${nombreBase}.xlsx`);
                    }
                } catch (err) {
                    console.error(`   ❌ Error en ${codigoTexto}: ${err.message}`);
                }
                finally {
                    activos--;
                }
            }));
        }

        console.log('⏳ Esperando a que finalicen todas las tareas paralelas...');
        await Promise.all(tareas);
        console.log('--- ✅ Todas las actas fueron generadas correctamente ---');
    } catch (error) {
        console.error('\n❌ Error:', error.message);
    }
}

ejecutarAutomatizacion();