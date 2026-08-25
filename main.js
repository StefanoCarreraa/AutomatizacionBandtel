// const ExcelJS = require('exceljs');
// const path = require('path');

// async function ejecutarAutomatizacion() {
//     const libroLista = new ExcelJS.Workbook();
//     const libroMatriz = new ExcelJS.Workbook();

//     try {
//         console.log('--- Iniciando App - AutomatizacionBandtel ---');

//         await libroLista.xlsx.readFile('input/Lista.xlsx');
//         await libroMatriz.xlsx.readFile('input/Tabulado Acta de Nodo V2.xlsx');

//         const hojaLista = libroLista.getWorksheet('Hoja1');

//         // --- 1. DEFINICIÓN DE PESTAÑAS Y CREACIÓN DE ÍNDICES (MAPAS) ---
//         const nombresHojas = {
//             info: '1. INFORMACIÓN DEL NODO',
//             equip: '2. EQUIPAMIENTO DE TRANSMISIÓN',
//             pop: '3. PUNTO DE PRESENCIA - POP',
//             energia: '4. SISTEMAS DE ENERGÍA & CLIMAT',
//             seguridad: '5. SEGURIDAD FISICA',
//             obras: '6. OBRAS CIVILES',
//             puestaTierra: '7. SISTEMA PUESTA A TIERRA',
//             fotos: '8. REPORTE FOTOGRÁFICO',
//             comentarios: '9. COMENTARIOS GENERALES',
//             firmas: '10. FIRMAS'
//         };

//         const indices = {
//             info: new Map(), equip: new Map(), pop: new Map(), energia: new Map(),
//             seguridad: new Map(), obras: new Map(), puestaTierra: new Map(),
//             fotos: new Map(), comentarios: new Map(), firmas: new Map()
//         };

//         console.log('⚡ Indexando todas las pestañas de la matriz...');

//         libroMatriz.getWorksheet(nombresHojas.info).eachRow(row => {
//             const cod = row.getCell('G').value?.toString().trim();
//             if (cod) indices.info.set(cod, row);
//         });

//         const mapeoColumnas = {
//             equip: 'C', pop: 'C', energia: 'C', seguridad: 'B',
//             obras: 'C', puestaTierra: 'C', fotos: 'C', comentarios: 'C', firmas: 'C'
//         };

//         for (const key in mapeoColumnas) {
//             const hoja = libroMatriz.getWorksheet(nombresHojas[key]);
//             if (hoja) {
//                 hoja.eachRow(row => {
//                     const cod = row.getCell(mapeoColumnas[key]).value?.toString().trim();
//                     if (cod) indices[key].set(cod, row);
//                 });
//             }
//         }

//         // --- 2. PROCESAMIENTO DE LA LISTA ---
//         for (let i = 1; ; i++) {
//             const codigoABuscar = hojaLista.getCell(`A${i}`).value;
//             if (!codigoABuscar) break;

//             const codigoTexto = codigoABuscar.toString().trim();
//             console.log(`\n🔍 Procesando: ${codigoTexto}`);

//             const fInfo = indices.info.get(codigoTexto);
//             const fEquip = indices.equip.get(codigoTexto);
//             const fPop = indices.pop.get(codigoTexto);
//             const fEnergia = indices.energia.get(codigoTexto);
//             const fSeguridad = indices.seguridad.get(codigoTexto);
//             const fObras = indices.obras.get(codigoTexto);
//             const fPuestaTierra = indices.puestaTierra.get(codigoTexto);
//             const fFotos = indices.fotos.get(codigoTexto);
//             const fComentarios = indices.comentarios.get(codigoTexto);
//             const fFirmas = indices.firmas.get(codigoTexto);

//             if (fInfo) {
//                 const libroModelo = new ExcelJS.Workbook();
//                 await libroModelo.xlsx.readFile('input/Modelo Acta de Nodo_Rev_BANDTEL.xlsx');
//                 const hojaDestino = libroModelo.getWorksheet(1);

//                 // --- MAPEADO: HOJA 1 (INFO NODO) ---
//                 const region = fInfo.getCell('A').value || '';
//                 const celdaTitulo = hojaDestino.getCell('B5');
//                 const textoOriginal = celdaTitulo.value ? celdaTitulo.value.toString() : "";
//                 if (textoOriginal.includes('_______')) {
//                     celdaTitulo.value = textoOriginal.replace('_______', region.toString().toUpperCase());
//                 }
//                 // Seccion 1 - Información General
//                 const esDistrital = fInfo.getCell('B').value?.toString().toUpperCase() === 'X';
//                 const esIntermedio = fInfo.getCell('C').value?.toString().toUpperCase() === 'X';
//                 const esTerminal = fInfo.getCell('D').value?.toString().toUpperCase() === 'X';
//                 const esOtro = fInfo.getCell('E').value?.toString().toUpperCase() === 'X';

//                 hojaDestino.getCell('B9').value = 'Nodo Distrital (           )';
//                 hojaDestino.getCell('E9').value = 'Nodo de Intermedio (           )';
//                 hojaDestino.getCell('H9').value = 'Nodo Terminal (           )';
//                 hojaDestino.getCell('K9').value = 'Otro (           )';

//                 if (esDistrital) {
//                     hojaDestino.getCell('B9').value = 'Nodo Distrital (     X     )';
//                 } else if (esIntermedio) {
//                     hojaDestino.getCell('E9').value = 'Nodo de Intermedio (     X     )';
//                 } else if (esTerminal) {
//                     hojaDestino.getCell('H9').value = 'Nodo Terminal (     X     )';
//                 } else if (esOtro) {
//                     hojaDestino.getCell('K9').value = 'Otro (     X     )';
//                 }

//                 ['B9', 'E9', 'H9', 'K9'].forEach(celda => {
//                     hojaDestino.getCell(celda).alignment = { horizontal: 'center', vertical: 'middle' };
//                     hojaDestino.getCell(celda).font = {
//                         name: 'Arial',
//                         size: 11,
//                         bold: true
//                     };
//                 });
//                 hojaDestino.getCell('D11').value = codigoTexto;
//                 hojaDestino.getCell('D10').value = fInfo.getCell('F').value;
//                 hojaDestino.getCell('K10').value = fInfo.getCell('M').value;
//                 hojaDestino.getCell('K11').value = fInfo.getCell('N').value;
//                 hojaDestino.getCell('D12').value = fInfo.getCell('H').value;
//                 hojaDestino.getCell('L12').value = fInfo.getCell('L').value;
//                 hojaDestino.getCell('D13').value = fInfo.getCell('I').value;
//                 hojaDestino.getCell('H13').value = fInfo.getCell('J').value;
//                 hojaDestino.getCell('L13').value = fInfo.getCell('K').value;

//                 // Seccion 1.1 - Información de Ubicacion
//                 hojaDestino.getCell('B16').value = fInfo.getCell('O').value;

//                 // --- MAPEADO: HOJA 2 (EQUIPAMIENTO)
//                 //
//                 // Seccion 2.1 - Radio Enlaces
//                 let filaActualDestino = 21;
//                 if (fEquip) {
//                     const configuracionEnlaces = [
//                         { nombre: 'Enlace 1', origenC: 'F', colInicioMatriz: 7 },
//                         { nombre: 'Enlace 2', origenC: 'AW', colInicioMatriz: 50 },
//                         { nombre: 'Enlace 3', origenC: 'CN', colInicioMatriz: 93 },
//                         { nombre: 'Enlace 4', origenC: 'EE', colInicioMatriz: 136 },
//                         { nombre: 'Enlace 5', origenC: 'FV', colInicioMatriz: 179 },
//                         { nombre: 'Enlace 6', origenC: 'HM', colInicioMatriz: 222 },
//                         { nombre: 'Enlace 7', origenC: 'JD', colInicioMatriz: 265 },
//                         { nombre: 'Enlace 8', origenC: 'KU', colInicioMatriz: 308 },
//                         { nombre: 'Enlace 9', origenC: 'ML', colInicioMatriz: 351 },
//                         { nombre: 'Enlace 10', origenC: 'OC', colInicioMatriz: 394 },
//                         { nombre: 'Enlace 11', origenC: 'PT', colInicioMatriz: 437 },
//                         { nombre: 'Enlace 12', origenC: 'RK', colInicioMatriz: 480 }
//                     ];

//                     // IMPORTANTE: Definimos el rango exacto de la tabla de Radio Enlaces
//                     const FILA_INICIO_TABLA = 21;
//                     const FILA_FIN_TABLA_ORIGINAL = 21;

//                     configuracionEnlaces.forEach((enlace, index) => {
//                         let colActual = enlace.colInicioMatriz;
//                         const numeroEnlace = index + 1;

//                         for (let i = 0; i < 7; i++) {
//                             const valorD = fEquip.getCell(colActual).value;
//                             if (!valorD || valorD.toString().trim().toUpperCase() === 'NA') {
//                                 colActual += 4;
//                                 continue;
//                             }

//                             // 1. GESTIÓN DE FILAS: Solo insertamos si superamos el espacio original
//                             if (filaActualDestino > FILA_FIN_TABLA_ORIGINAL) {
//                                 hojaDestino.insertRow(filaActualDestino, {});
//                             }

//                             const row = hojaDestino.getRow(filaActualDestino);
//                             row.height = 30;

//                             // 2. LIMPIEZA QUIRÚRGICA DE MERGES (Solo de la fila actual)
//                             if (hojaDestino._merges) {
//                                 Object.keys(hojaDestino._merges).forEach(key => {
//                                     const m = hojaDestino._merges[key];
//                                     if (m.top === filaActualDestino) {
//                                         delete hojaDestino._merges[key];
//                                     }
//                                 });
//                             }

//                             // 3. APLICAR FORMATO CELDA POR CELDA (B hasta M)
//                             for (let c = 2; c <= 13; c++) {
//                                 const cell = row.getCell(c);
//                                 cell.value = null;
//                                 cell.border = {
//                                     top: { style: 'thin' },
//                                     left: { style: 'thin' },
//                                     bottom: { style: 'thin' },
//                                     right: { style: 'thin' }
//                                 };
//                                 cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
//                                 cell.font = { name: 'Arial', size: 11 };
//                             }

//                             // 4. UNIR CELDAS (Sin conflicto ahora que borramos el merge de esta fila)
//                             try {
//                                 hojaDestino.mergeCells(filaActualDestino, 4, filaActualDestino, 6);   // D-F
//                                 hojaDestino.mergeCells(filaActualDestino, 7, filaActualDestino, 8);   // G-H
//                                 hojaDestino.mergeCells(filaActualDestino, 9, filaActualDestino, 10);  // I-J
//                                 hojaDestino.mergeCells(filaActualDestino, 12, filaActualDestino, 13); // L-M
//                             } catch (e) { }

//                             // 5. ASIGNACIÓN DE VALORES
//                             row.getCell('B').value = numeroEnlace;
//                             row.getCell('C').value = fEquip.getCell(enlace.origenC).value;
//                             row.getCell('D').value = valorD;

//                             let posMatriz = colActual + 1;
//                             row.getCell('G').value = fEquip.getCell(posMatriz++).value;
//                             row.getCell('I').value = fEquip.getCell(posMatriz++).value;
//                             row.getCell('K').value = fEquip.getCell(posMatriz++).value;
//                             row.getCell('L').value = "-";

//                             filaActualDestino++;
//                             colActual += 4;
//                         }
//                     });

//                     // 6. LIMPIEZA DE FILAS SOBRANTES (Si hubo menos enlaces que los originales)
//                     if (filaActualDestino <= FILA_FIN_TABLA_ORIGINAL) {
//                         for (let r = filaActualDestino; r <= FILA_FIN_TABLA_ORIGINAL; r++) {
//                             const row = hojaDestino.getRow(r);
//                             row.eachCell((cell) => cell.value = null);
//                         }
//                     }
//                 }

//                 // --- PRUEBA SECCIÓN 2.2 ---
//                 let filaSeccion2_2 = filaActualDestino + 5;
//                 // --- PROCESAMIENTO DINÁMICO SECCIÓN 2.2 (EQUIPOS DE RED) ---
//                 if (fEquip) {
//                     const hojaEquipMatriz = libroMatriz.getWorksheet(nombresHojas.equip);
//                     const estiloBold = { name: 'Arial', size: 11, bold: true };

//                     const bloquesEquipos = [
//                         { fijo: 'TA3', datos: ['TA', 'TB', 'TC', 'TD'] },
//                         { fijo: 'TE3', datos: ['TE', 'TF', 'TG', 'TH'] },
//                         { fijo: 'TI3', datos: ['TI', 'TJ', 'TK', 'TL'] },
//                         { fijo: 'TM3', datos: ['TM', 'TN', 'TO', 'TP'] }
//                     ];

//                     let filaActualSeccion2 = filaActualDestino + 5;
//                     let primerBloqueEscrito = false;

//                     bloquesEquipos.forEach(bloque => {
//                         const primerValor = fEquip.getCell(bloque.datos[0]).value;

//                         if (!primerValor || primerValor.toString().trim().toUpperCase() === 'NA') {
//                             return;
//                         }

//                         if (primerBloqueEscrito) {
//                             hojaDestino.insertRow(filaActualSeccion2, {});
//                         }

//                         const row = hojaDestino.getRow(filaActualSeccion2);
//                         row.height = 25;

//                         // 1. Limpieza de merges previa en la fila
//                         if (hojaDestino._merges) {
//                             Object.keys(hojaDestino._merges).forEach(key => {
//                                 const m = hojaDestino._merges[key];
//                                 if (m.top === filaActualSeccion2) {
//                                     delete hojaDestino._merges[key];
//                                 }
//                             });
//                         }

//                         // 2. Aplicar bordes y formato base (Columnas B a M)
//                         for (let c = 2; c <= 13; c++) {
//                             const cell = row.getCell(c);
//                             cell.border = {
//                                 top: { style: 'thin' }, left: { style: 'thin' },
//                                 bottom: { style: 'thin' }, right: { style: 'thin' }
//                             };
//                             cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
//                             cell.font = { name: 'Arial', size: 11 };
//                         }

//                         // 3. ASIGNACIÓN DE VALORES
//                         // Título en B (irá en el merge B-D)
//                         row.getCell('B').value = hojaEquipMatriz.getCell(bloque.fijo).value;
//                         row.getCell('B').font = estiloBold;
//                         row.getCell('B').alignment = { horizontal: 'left', vertical: 'middle' };

//                         // Datos en E, G, I, K
//                         row.getCell('E').value = fEquip.getCell(bloque.datos[0]).value; // Marca/Modelo
//                         row.getCell('G').value = fEquip.getCell(bloque.datos[1]).value; // Serie
//                         row.getCell('I').value = fEquip.getCell(bloque.datos[2]).value; // Versión
//                         row.getCell('K').value = fEquip.getCell(bloque.datos[3]).value; // Estado/Otro

//                         // 4. APLICAR MERGES SOLICITADOS
//                         try {
//                             hojaDestino.mergeCells(filaActualSeccion2, 2, filaActualSeccion2, 4);   // B, C, D (Título)
//                             hojaDestino.mergeCells(filaActualSeccion2, 5, filaActualSeccion2, 6);   // E, F
//                             hojaDestino.mergeCells(filaActualSeccion2, 7, filaActualSeccion2, 8);   // G, H
//                             hojaDestino.mergeCells(filaActualSeccion2, 9, filaActualSeccion2, 10);  // I, J
//                             hojaDestino.mergeCells(filaActualSeccion2, 11, filaActualSeccion2, 13); // K, L, M
//                         } catch (e) {
//                             console.log(`⚠️ Error al unir celdas en fila ${filaActualSeccion2}:`, e.message);
//                         }

//                         filaActualSeccion2++;
//                         primerBloqueEscrito = true;
//                     });

//                     filaActualDestino = filaActualSeccion2;
//                 }

//                 // --- PROCESAMIENTO DINÁMICO SECCIÓN 2.3 ---
//                 let filaActualSeccion2_3 = filaActualDestino + 3;
//                 // --- PROCESAMIENTO SECCIÓN 2.3 ---
//                 if (fEquip) {
//                     let primerBloque2_3Escrito = false;

//                     const bloquesDetalle = [
//                         { inicio: 'AJ', c: ['AJ', 'AK'], d_m: ['AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU'] },
//                         { inicio: 'CA', c: ['CA', 'CB'], d_m: ['CC', 'CD', 'CE', 'CF', 'CG', 'CH', 'CI', 'CJ', 'CK', 'CL'] },
//                         { inicio: 'DR', c: ['DR', 'DS'], d_m: ['DT', 'DU', 'DV', 'DW', 'DX', 'DY', 'DZ', 'EA', 'EB', 'EC'] },
//                         { inicio: 'FI', c: ['FI', 'FJ'], d_m: ['FK', 'FL', 'FM', 'FN', 'FO', 'FP', 'FQ', 'FR', 'FS', 'FT'] },
//                         { inicio: 'GZ', c: ['GZ', 'HA'], d_m: ['HB', 'HC', 'HD', 'HE', 'HF', 'HG', 'HH', 'HI', 'HJ', 'HK'] },
//                         { inicio: 'IQ', c: ['IQ', 'IR'], d_m: ['IS', 'IT', 'IU', 'IV', 'IW', 'IX', 'IY', 'IZ', 'JA', 'JB'] },
//                         { inicio: 'KH', c: ['KH', 'KI'], d_m: ['KJ', 'KK', 'KL', 'KM', 'KN', 'KO', 'KP', 'KQ', 'KR', 'KS'] },
//                         { inicio: 'LY', c: ['LY', 'LZ'], d_m: ['MA', 'MB', 'MC', 'MD', 'ME', 'MF', 'MG', 'MH', 'MI', 'MJ'] },
//                         { inicio: 'NP', c: ['NP', 'NQ'], d_m: ['NR', 'NS', 'NT', 'NU', 'NV', 'NW', 'NX', 'NY', 'NZ', 'OA'] },
//                         { inicio: 'PG', c: ['PG', 'PH'], d_m: ['PI', 'PJ', 'PK', 'PL', 'PM', 'PN', 'PO', 'PP', 'PQ', 'PR'] },
//                         { inicio: 'QX', c: ['QX', 'QY'], d_m: ['QZ', 'RA', 'RB', 'RC', 'RD', 'RE', 'RF', 'RG', 'RH', 'RI'] },
//                         { inicio: 'SO', c: ['SO', 'SP'], d_m: ['SQ', 'SR', 'SS', 'ST', 'SU', 'SV', 'SW', 'SX', 'SY', 'SZ'] }
//                     ];

//                     bloquesDetalle.forEach((bloque, index) => {
//                         // Validamos NA usando la primera columna del bloque de datos (d_m[0])
//                         const valorRef = fEquip.getCell(bloque.d_m[0]).value;
//                         if (!valorRef || valorRef.toString().trim().toUpperCase() === 'NA') return;

//                         // Inserción de fila si no es la primera
//                         if (primerBloque2_3Escrito) {
//                             hojaDestino.insertRow(filaActualSeccion2_3, {});
//                         }

//                         const row = hojaDestino.getRow(filaActualSeccion2_3);

//                         // --- APLICACIÓN DE ESTILOS FILA POR FILA ---
//                         for (let c = 2; c <= 13; c++) {
//                             const cell = row.getCell(c);

//                             // 1. Bordes delgados en los 4 lados
//                             cell.border = {
//                                 top: { style: 'thin' },
//                                 left: { style: 'thin' },
//                                 bottom: { style: 'thin' },
//                                 right: { style: 'thin' }
//                             };

//                             // 2. Alineación y Ajuste de texto (Crucial para que crezca hacia abajo)
//                             cell.alignment = {
//                                 vertical: 'middle',
//                                 horizontal: 'center',
//                                 wrapText: true // <--- Esto hace que el texto no se salga de la celda
//                             };

//                             cell.font = { name: 'Arial', size: 11 };
//                         }

//                         // --- ASIGNACIÓN DE VALORES ---
//                         row.getCell('B').value = index + 1;

//                         const v1 = fEquip.getCell(bloque.c[0]).value || '';
//                         const v2 = fEquip.getCell(bloque.c[1]).value || '';
//                         row.getCell('C').value = `${v1} / ${v2}`;

//                         const columnasDestino = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'];
//                         bloque.d_m.forEach((colMatriz, i) => {
//                             row.getCell(columnasDestino[i]).value = fEquip.getCell(colMatriz).value;
//                         });

//                         // Forzar que la altura sea automática para que el wrapText funcione visualmente
//                         row.height = undefined;

//                         filaActualSeccion2_3++;
//                         primerBloque2_3Escrito = true;
//                     });

//                     filaActualDestino = filaActualSeccion2_3;
//                 }

//                 // --- PROCESAMIENTO DINÁMICO SECCIÓN 2.4 ---
//                 let filaActualSeccion2_4 = filaActualDestino + 3;

//                 if (fEquip) {
//                     let primerBloque2_4Escrito = false;
//                     const iniciosColumnas = ['TQ', 'TX', 'UE', 'UL', 'US', 'UZ'];

//                     iniciosColumnas.forEach((letraInicio, index) => {
//                         const celdaMatrizBase = fEquip.getCell(letraInicio);
//                         const valorRef = celdaMatrizBase.value;

//                         if (!valorRef || valorRef.toString().trim().toUpperCase() === 'NA') return;

//                         if (primerBloque2_4Escrito) {
//                             hojaDestino.insertRow(filaActualSeccion2_4, {});
//                         }

//                         const rowDestino = hojaDestino.getRow(filaActualSeccion2_4);
//                         const rNum = filaActualSeccion2_4;

//                         // Limpiar merges antiguos en esta fila específica
//                         if (hojaDestino._merges) {
//                             Object.keys(hojaDestino._merges).forEach(key => {
//                                 const m = hojaDestino._merges[key];
//                                 if (m.top === rNum) delete hojaDestino._merges[key];
//                             });
//                         }

//                         // Formato
//                         for (let c = 2; c <= 13; c++) {
//                             const cell = rowDestino.getCell(c);
//                             cell.border = {
//                                 top: { style: 'thin' }, left: { style: 'thin' },
//                                 bottom: { style: 'thin' }, right: { style: 'thin' }
//                             };
//                             cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
//                             cell.font = { name: 'Arial', size: 11 };
//                         }

//                         const colIdx = celdaMatrizBase.col;

//                         // Mapeo corregido para celdas combinadas
//                         rowDestino.getCell('B').value = index + 1;
//                         rowDestino.getCell('C').value = fEquip.getCell(colIdx).value;     // TQ (C-D)
//                         rowDestino.getCell('E').value = fEquip.getCell(colIdx + 1).value; // TR (E)
//                         rowDestino.getCell('F').value = fEquip.getCell(colIdx + 2).value; // TS (F)
//                         rowDestino.getCell('G').value = fEquip.getCell(colIdx + 3).value; // TT (G-H)
//                         rowDestino.getCell('I').value = fEquip.getCell(colIdx + 4).value; // TU (I-J)
//                         rowDestino.getCell('K').value = fEquip.getCell(colIdx + 5).value; // TV (K-L)
//                         rowDestino.getCell('M').value = fEquip.getCell(colIdx + 6).value; // TW (M)

//                         try {
//                             hojaDestino.mergeCells(rNum, 3, rNum, 4);   // C-D
//                             hojaDestino.mergeCells(rNum, 7, rNum, 8);   // G-H
//                             hojaDestino.mergeCells(rNum, 9, rNum, 10);  // I-J
//                             hojaDestino.mergeCells(rNum, 11, rNum, 12); // K-L
//                         } catch (e) { }

//                         rowDestino.height = undefined;
//                         filaActualSeccion2_4++;
//                         primerBloque2_4Escrito = true;
//                     });
//                     filaActualDestino = filaActualSeccion2_4;
//                 }

//                 // --- MAPEADO: HOJA 4 (ENERGÍA) - EJEMPLO ---
//                 if (fEnergia) {
//                     // hojaDestino.getCell('B40').value = fEnergia.getCell('E').value;
//                 }

//                 // --- GUARDADO ---
//                 const nombreBase = `Acta_${codigoTexto}`;
//                 const rutaExcel = path.join(__dirname, 'output', `${nombreBase}.xlsx`);
//                 await libroModelo.xlsx.writeFile(rutaExcel);
//                 console.log(`   ✅ Excel generado correctamente.`);

//             } else {
//                 console.log(`   ⚠️ Código ${codigoTexto} no encontrado en la pestaña de Información.`);
//             }
//         }
//         console.log('\n--- PROCESO FINALIZADO ---');
//     } catch (error) {
//         console.error('\n❌ Error:', error.message);
//     }
// }

// ejecutarAutomatizacion();