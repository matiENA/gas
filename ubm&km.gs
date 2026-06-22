/**
 * ============================================================================
 * EXTRACCIÓN MASIVA: KILOMETRAJE + VIAJES CAMPO + HOJAS DE RUTA
 * (Arquitectura de Diccionario Puro para evitar Datos Fantasma)
 * ============================================================================
 */

function procesarKilometrosYViajesCore(minDate, maxDate, hacerMerge = false) {
  const ssMaestro = SpreadsheetApp.getActiveSpreadsheet();
  const ID_SHEET_KILOMETROS = '1Wr-_P4mDvldif_cAx08sp7yT8uTUrajI2HQAJF6tnGM';

  let mapaKms = {};
  let viajesAgrupados = {};
  let viajesDetalleObj = {}; // 👉 NUEVO: Ahora es un Objeto, no un Array
  
  let cacheSheet = ssMaestro.getSheetByName("API_CACHE_BASICO");
  
  // --- 1. LECTURA Y PRESERVACIÓN DE CACHÉ EXISTENTE (Merge Inteligente) ---
  if (hacerMerge && cacheSheet) {
    // A) Preservar api_km clásico
    try {
      const hojaKm = ssMaestro.getSheetByName('api_km');
      if (hojaKm && hojaKm.getLastRow() > 0) {
        let kmStr = "";
        hojaKm.getDataRange().getValues().forEach(row => {
          row.forEach(cell => { if (cell) kmStr += String(cell).replace(/^'/, ""); });
        });
        if (kmStr) {
          let cacheVieja = JSON.parse(kmStr);
          for (let chofer in cacheVieja) {
            mapaKms[chofer] = cacheVieja[chofer].filter(registro => {
              let partes = registro.fechaCorta.split('/');
              let fRegistro = new Date("20" + partes[2], partes[1] - 1, partes[0]);
              return fRegistro < minDate || fRegistro > maxDate; 
            });
          }
        }
      }
    } catch(e) { console.warn("No se pudo procesar merge de api_km."); }

    // B) PRESERVAR EL HISTORIAL DE LA FILA 12 (Purga Temporal)
    try {
      if (cacheSheet.getMaxRows() < 12) cacheSheet.insertRowsAfter(cacheSheet.getMaxRows(), 12 - cacheSheet.getMaxRows());
      let lastCol = cacheSheet.getLastColumn() || 1;
      let dataFila12 = cacheSheet.getRange(12, 1, 1, lastCol).getValues()[0];
      let jsonNuevaSeccionRaw = dataFila12.filter(String).map(c => String(c || "").replace(/^'/, "")).join("");
      
      if (jsonNuevaSeccionRaw) {
        let registrosAnteriores = JSON.parse(jsonNuevaSeccionRaw);
        
        // Verificamos que ya tenga el formato de objeto nuevo
        if (registrosAnteriores && typeof registrosAnteriores === 'object' && !Array.isArray(registrosAnteriores)) {
            viajesDetalleObj = registrosAnteriores;
            
            // 👉 MAGIA PURA: Borramos de la memoria todo lo que caiga en la ventana temporal actual.
            for (let chofer in viajesDetalleObj) {
                for (let fecha in viajesDetalleObj[chofer]) {
                    let fReg = new Date(fecha + "T12:00:00");
                    if (fReg >= minDate && fReg <= maxDate) {
                        delete viajesDetalleObj[chofer][fecha]; 
                    }
                }
                // Limpieza de choferes que se quedaron sin viajes en el historial
                if (Object.keys(viajesDetalleObj[chofer]).length === 0) {
                    delete viajesDetalleObj[chofer];
                }
            }
        }
      }
    } catch(e) { console.warn("No se pudo realizar merge de la fila 12 de viajes detallados."); }
  }

  // --- 2. LECTURA DE LA HOJA DE KILÓMETROS MAESTRA ---
  let dataKm;
  try {
    const ssKm = SpreadsheetApp.openById(ID_SHEET_KILOMETROS);
    const sheetKm = ssKm.getSheetByName('KM') || ssKm.getSheets()[0]; 
    dataKm = sheetKm.getDataRange().getValues();
  } catch(e) { 
    console.error("Error leyendo planilla de KM:", e); 
    return; 
  }

  const parseNumberSafe = (val) => {
      if (typeof val === 'number') return val;
      let limpio = String(val || '').replace(/,/g, '.').replace(/[^0-9.-]/g, '');
      return parseFloat(limpio) || 0;
  };
  
  const parseSafeDateLocal = (rawDate) => {
     if (rawDate instanceof Date) return rawDate;
     let strDate = String(rawDate).split('-')[0].trim();
     let parts = strDate.split(/[\/\-]/);
     if (parts.length >= 3) {
        if(parts[0].length <= 2 && parts[2].length >= 2) { 
          let anio = parts[2].length === 2 ? "20" + parts[2] : parts[2];
          return new Date(anio, parseInt(parts[1], 10) - 1, parts[0]);
        } else if (parts[0].length === 4) { 
          return new Date(parts[0], parseInt(parts[1], 10) - 1, parts[2]);
        }
     }
     return new Date(strDate); 
  };

  const normalizarNombre = (n) => String(n).trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, ' ');
  const mesesAbrev = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];

  let contadorExtraidos = 0;

  // --- 3. BUCLE MAESTRO ---
  for (let i = 1; i < dataKm.length; i++) {
    let row = dataKm[i];

    let dominioRaw  = row[0];   // Col A
    let fechaRaw    = row[1];   // Col B
    let nombreRaw   = row[2];   // Col C
    let livianoRaw  = row[3];   // Col D
    let euroRaw     = row[4];   // Col E
    let campoRaw    = row[5];   // Col F
    let infiniaDRaw = row[7];   // Col H
    let kmBackupRaw = row[8];   // Col I 
    let kmBaseRaw   = row[16];  // Col Q 
    let hojaRutaRaw = row[19];  // Col T

    let nOriginal = String(nombreRaw).trim().toLowerCase(); 
    if (!nOriginal) continue;
    
    let nombreNorm = normalizarNombre(nOriginal); 
    let dObj = parseSafeDateLocal(fechaRaw);
    
    if (!dObj || isNaN(dObj.getTime())) continue;

    // Solo extraemos si la fecha entra en la ventana
    if (dObj >= minDate && dObj <= maxDate) {
      
      // A) KILÓMETROS GLOBALES
      let km = parseFloat(kmBaseRaw);
      if (isNaN(km) || km === 0) km = parseFloat(kmBackupRaw) || 0;        
      
      if (km > 0) {
        let dd = String(dObj.getDate()).padStart(2, '0');
        let mm = String(dObj.getMonth() + 1).padStart(2, '0');
        let yy = String(dObj.getFullYear()).slice(-2);
        let datePart = `${dd}/${mm}/${yy}`;
        
        if (!mapaKms[nOriginal]) mapaKms[nOriginal] = [];
        mapaKms[nOriginal].push({ fechaCorta: datePart, km: km }); 
      }

      // B) VIAJES CAMPO Y HOJAS DE RUTA DETALLADOS
      let livianoNum  = parseNumberSafe(livianoRaw);
      let euroNum     = parseNumberSafe(euroRaw);
      let campoNum    = parseNumberSafe(campoRaw);
      let infiniaDNum = parseNumberSafe(infiniaDRaw);
      let hojaStr     = String(hojaRutaRaw || "").trim();

      if (campoNum > 0 || livianoNum > 0 || euroNum > 0 || infiniaDNum > 0 || hojaStr !== "") {
          
          let mesAnio = `${mesesAbrev[dObj.getMonth()]}-${String(dObj.getFullYear()).slice(-2)}`;
          if (!viajesAgrupados[nombreNorm]) viajesAgrupados[nombreNorm] = {};
          if (!viajesAgrupados[nombreNorm][mesAnio]) {
              viajesAgrupados[nombreNorm][mesAnio] = { km: 0, liviano: 0, euro: 0, infiniaD: 0 };
          }
          viajesAgrupados[nombreNorm][mesAnio].km += campoNum;
          viajesAgrupados[nombreNorm][mesAnio].liviano += livianoNum;
          viajesAgrupados[nombreNorm][mesAnio].euro += euroNum;
          viajesAgrupados[nombreNorm][mesAnio].infiniaD += infiniaDNum;

          let isoStr = dObj.toISOString().split('T')[0];

          // 👉 NUEVA SECCIÓN: Asignación limpia al Diccionario
          if (!viajesDetalleObj[nombreNorm]) viajesDetalleObj[nombreNorm] = {};
          
          // Si hay múltiples filas en la planilla para un chofer el mismo día, se fusionan
          if (!viajesDetalleObj[nombreNorm][isoStr]) {
              viajesDetalleObj[nombreNorm][isoStr] = {
                  dominio: String(dominioRaw || "").trim(),
                  liviano: 0, euro: 0, campo: 0, infiniaD: 0,
                  hoja_ruta: []
              };
          }
          
          let target = viajesDetalleObj[nombreNorm][isoStr];
          target.liviano += livianoNum;
          target.euro += euroNum;
          target.campo += campoNum;
          target.infiniaD += infiniaDNum;

          if (hojaStr !== "") {
              let arrHojas = hojaStr.split(',').map(s => s.trim()).filter(Boolean);
              arrHojas.forEach(h => {
                  if (!target.hoja_ruta.includes(h)) target.hoja_ruta.push(h);
              });
          }
          contadorExtraidos++;
      }
    }
  }

  // --- 4. ESCRITURA ATÓMICA EN CACHÉ ---
  let hojaKm = ssMaestro.getSheetByName('api_km');
  if (!hojaKm) { hojaKm = ssMaestro.insertSheet('api_km'); hojaKm.hideSheet(); }
  hojaKm.clearContents(); 
  
  let kmChunks = [];
  let kmStr = JSON.stringify(mapaKms);
  for (let i = 0; i < kmStr.length; i += 40000) { kmChunks.push(["'" + kmStr.substring(i, i + 40000)]); }
  if (kmChunks.length > 0) hojaKm.getRange(1, 1, kmChunks.length, 1).setValues(kmChunks);

  if (cacheSheet && typeof escribirChunksEnFila === 'function') {
    escribirChunksEnFila(cacheSheet, 7, JSON.stringify(viajesAgrupados));
    escribirChunksEnFila(cacheSheet, 12, JSON.stringify(viajesDetalleObj)); // 👉 Guardamos el Objeto Limpio
  }
  
  ssMaestro.toast(`Proceso OK. ${contadorExtraidos} viajes detallados cacheados.`, "✅ Éxito");
}

function generarJSONKilometros_Frecuente() {
  const hoy = new Date(); hoy.setHours(23, 59, 59, 999);
  const hace60Dias = new Date(hoy); hace60Dias.setDate(hoy.getDate() - 60); hace60Dias.setHours(0, 0, 0, 0);
  procesarKilometrosYViajesCore(hace60Dias, hoy, true); 
}

function generarJSONKilometros_Completo() {
  const hoy = new Date(); hoy.setHours(23, 59, 59, 999);
  const hace1Ano = new Date(hoy); hace1Ano.setFullYear(hoy.getFullYear() - 1); hace1Ano.setHours(0, 0, 0, 0);
  procesarKilometrosYViajesCore(hace1Ano, hoy, false);
}

function ubmkm() {
  generarJSONKilometros_Completo();
}

/**
 * ============================================================================
 * EXTRACCIÓN QUIRÚRGICA: onEdit PARA HOJA DE RUTA
 * ============================================================================
 */
function alEditarKilometros(e) {
  if (!e || !e.range) return;

  const fila = e.range.getRow();
  const columna = e.range.getColumn();
  const sheet = e.source.getActiveSheet();
  const nombreHoja = sheet.getName();

  if (fila >= 2 && columna === 20 && nombreHoja.toUpperCase() === 'KM') {
    try {
      const ID_PLANILLA_MAESTRA = '1eQ9Y5diL5fwxYTxvseNgZJFbX-lSUQ13axbp3cLiqPc'; 
      
      let ssMaestro;
      try { ssMaestro = SpreadsheetApp.openById(ID_PLANILLA_MAESTRA); } 
      catch (err) { ssMaestro = SpreadsheetApp.getActiveSpreadsheet(); }

      const cacheSheet = ssMaestro.getSheetByName("API_CACHE_BASICO");
      if (!cacheSheet) return;

      const rangoFila = sheet.getRange(fila, 1, 1, 20).getValues()[0];
      const dominioRaw  = rangoFila[0];   
      const fechaRaw    = rangoFila[1];   
      const nombreRaw   = rangoFila[2];   
      const livianoNum  = parseFloat(String(rangoFila[3] || '').replace(/,/g, '.').replace(/[^0-9.-]/g, '')) || 0;
      const euroNum     = parseFloat(String(rangoFila[4] || '').replace(/,/g, '.').replace(/[^0-9.-]/g, '')) || 0;
      const campoNum    = parseFloat(String(rangoFila[5] || '').replace(/,/g, '.').replace(/[^0-9.-]/g, '')) || 0;
      const infiniaDNum = parseFloat(String(rangoFila[7] || '').replace(/,/g, '.').replace(/[^0-9.-]/g, '')) || 0;
      const nuevoValorHR = String(rangoFila[19] || "").trim();

      if (!nombreRaw || !fechaRaw) return;

      const normalizar = (n) => String(n).trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, ' ');
      const choferNorm = normalizar(nombreRaw);

      let fechaIso = "";
      if (fechaRaw instanceof Date) {
        let tempDate = new Date(fechaRaw.getTime() - (fechaRaw.getTimezoneOffset() * 60000));
        fechaIso = tempDate.toISOString().split('T')[0];
      } else {
        let parts = String(fechaRaw).split('-')[0].trim().split(/[\/\-]/);
        if (parts.length >= 3) {
          let aa = parts[2].length === 2 ? "20" + parts[2] : parts[2];
          fechaIso = `${aa}-${String(parts[1]).padStart(2,'0')}-${String(parts[0]).padStart(2,'0')}`;
        }
      }
      if (!fechaIso) return;

      if (cacheSheet.getMaxRows() < 12) cacheSheet.insertRowsAfter(cacheSheet.getMaxRows(), 12 - cacheSheet.getMaxRows());
      let lastCol = cacheSheet.getLastColumn() || 1;
      let dataFila12 = cacheSheet.getRange(12, 1, 1, lastCol).getValues()[0];
      let jsonRaw = dataFila12.filter(String).map(c => String(c || "").replace(/^'/, "")).join("");
      
      let viajesDetalleObj = {};
      if (jsonRaw) {
        try { 
            let parsed = JSON.parse(jsonRaw); 
            if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) viajesDetalleObj = parsed;
        } catch(err) {}
      }

      let hojasParseadas = nuevoValorHR !== "" ? nuevoValorHR.split(',').map(s => s.trim()).filter(Boolean) : [];

      // 👉 Inyectamos limpiamente como Diccionario
      if (!viajesDetalleObj[choferNorm]) viajesDetalleObj[choferNorm] = {};
      
      if (!viajesDetalleObj[choferNorm][fechaIso]) {
          viajesDetalleObj[choferNorm][fechaIso] = {
              dominio: String(dominioRaw || "").trim(),
              liviano: livianoNum,
              euro: euroNum,
              campo: campoNum,
              infiniaD: infiniaDNum,
              hoja_ruta: []
          };
      }
      viajesDetalleObj[choferNorm][fechaIso].hoja_ruta = hojasParseadas;

      if (typeof escribirChunksEnFila === 'function') {
        escribirChunksEnFila(cacheSheet, 12, JSON.stringify(viajesDetalleObj));
        sheet.getParent().toast(`Hoja Ruta [${hojasParseadas.join(',')}] sincronizada`, "⚡ Caché Actualizado");

        // 👇 EL GATILLO SEGURO Y BIEN UBICADO 👇
        if (typeof notificarBackendNode === 'function') {
            notificarBackendNode('KM');
        }
      }

    } catch (error) {
      console.error("Error crítico en inyección onEdit Fila 12:", error);
    }
  }
}

// ====================================================================
// LECTURA DIRECTA: EVITA CACHÉS Y FANTASMAS AL ELIMINAR ROWS
// ====================================================================
function obtenerViajesYHRDirecto() {
    try {
      const ID_SHEET_KILOMETROS = '1Wr-_P4mDvldif_cAx08sp7yT8uTUrajI2HQAJF6tnGM';
      const ssKm = SpreadsheetApp.openById(ID_SHEET_KILOMETROS);
      const sheetKm = ssKm.getSheetByName('KM') || ssKm.getSheets()[0];
      
      const dataKm = sheetKm.getDataRange().getValues();
      let viajesDetalleObj = {};
      
      const limiteDate = new Date();
      limiteDate.setDate(limiteDate.getDate() - 60);
      
      const normalizar = (n) => String(n).trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, ' ');
      
      for (let i = 1; i < dataKm.length; i++) {
        let row = dataKm[i];
        let fechaRaw = row[1];
        let nombreRaw = row[2];
        
        if (!fechaRaw || !nombreRaw) continue;
        
        let dObj;
        if (fechaRaw instanceof Date) { dObj = fechaRaw; } 
        else {
          let parts = String(fechaRaw).split('-')[0].trim().split(/[\/\-]/);
          if(parts.length >= 3) {
            let aa = parts[2].length === 2 ? "20" + parts[2] : parts[2];
            dObj = new Date(aa, parseInt(parts[1], 10) - 1, parts[0]);
          } else { dObj = new Date(fechaRaw); }
        }
        
        if (isNaN(dObj.getTime()) || dObj < limiteDate) continue;
        
        let choferNorm = normalizar(nombreRaw);
        let dominioRaw  = row[0];
        let livianoNum  = parseFloat(String(row[3] || '').replace(/,/g, '.').replace(/[^0-9.-]/g, '')) || 0;
        let euroNum     = parseFloat(String(row[4] || '').replace(/,/g, '.').replace(/[^0-9.-]/g, '')) || 0;
        let campoNum    = parseFloat(String(row[5] || '').replace(/,/g, '.').replace(/[^0-9.-]/g, '')) || 0;
        let infiniaDNum = parseFloat(String(row[7] || '').replace(/,/g, '.').replace(/[^0-9.-]/g, '')) || 0;
        let hojaStr     = String(row[19] || "").trim();
        
        if (campoNum > 0 || livianoNum > 0 || euroNum > 0 || infiniaDNum > 0 || hojaStr !== "") {
          let tempDate = new Date(dObj.getTime() - (dObj.getTimezoneOffset() * 60000));
          let fechaIso = tempDate.toISOString().split('T')[0];
          
          if (!viajesDetalleObj[choferNorm]) viajesDetalleObj[choferNorm] = {};
          if (!viajesDetalleObj[choferNorm][fechaIso]) {
            viajesDetalleObj[choferNorm][fechaIso] = {
              dominio: String(dominioRaw || "").trim(),
              liviano: 0, euro: 0, campo: 0, infiniaD: 0,
              hoja_ruta: []
            };
          }
          
          let target = viajesDetalleObj[choferNorm][fechaIso];
          target.liviano += livianoNum;
          target.euro += euroNum;
          target.campo += campoNum;
          target.infiniaD += infiniaDNum;
          
          if (hojaStr !== "") {
            let arrHojas = hojaStr.split(',').map(s => s.trim()).filter(Boolean);
            arrHojas.forEach(h => {
              if (!target.hoja_ruta.includes(h)) target.hoja_ruta.push(h);
            });
          }
        }
      }
      
      return JSON.stringify(viajesDetalleObj);
    } catch(e) {
      console.error("Error Lectura Directa HR:", e);
      return JSON.stringify({});
    }
}/**
 * ============================================================================
 * EXTRACCIÓN MASIVA: KILOMETRAJE + VIAJES CAMPO + HOJAS DE RUTA
 * (Arquitectura de Diccionario Puro para evitar Datos Fantasma)
 * ============================================================================
 */

function procesarKilometrosYViajesCore(minDate, maxDate, hacerMerge = false) {
  const ssMaestro = SpreadsheetApp.getActiveSpreadsheet();
  const ID_SHEET_KILOMETROS = '1Wr-_P4mDvldif_cAx08sp7yT8uTUrajI2HQAJF6tnGM';

  let mapaKms = {};
  let viajesAgrupados = {};
  let viajesDetalleObj = {}; // 👉 NUEVO: Ahora es un Objeto, no un Array
  
  let cacheSheet = ssMaestro.getSheetByName("API_CACHE_BASICO");
  
  // --- 1. LECTURA Y PRESERVACIÓN DE CACHÉ EXISTENTE (Merge Inteligente) ---
  if (hacerMerge && cacheSheet) {
    // A) Preservar api_km clásico
    try {
      const hojaKm = ssMaestro.getSheetByName('api_km');
      if (hojaKm && hojaKm.getLastRow() > 0) {
        let kmStr = "";
        hojaKm.getDataRange().getValues().forEach(row => {
          row.forEach(cell => { if (cell) kmStr += String(cell).replace(/^'/, ""); });
        });
        if (kmStr) {
          let cacheVieja = JSON.parse(kmStr);
          for (let chofer in cacheVieja) {
            mapaKms[chofer] = cacheVieja[chofer].filter(registro => {
              let partes = registro.fechaCorta.split('/');
              let fRegistro = new Date("20" + partes[2], partes[1] - 1, partes[0]);
              return fRegistro < minDate || fRegistro > maxDate; 
            });
          }
        }
      }
    } catch(e) { console.warn("No se pudo procesar merge de api_km."); }

    // B) PRESERVAR EL HISTORIAL DE LA FILA 12 (Purga Temporal)
    try {
      if (cacheSheet.getMaxRows() < 12) cacheSheet.insertRowsAfter(cacheSheet.getMaxRows(), 12 - cacheSheet.getMaxRows());
      let lastCol = cacheSheet.getLastColumn() || 1;
      let dataFila12 = cacheSheet.getRange(12, 1, 1, lastCol).getValues()[0];
      let jsonNuevaSeccionRaw = dataFila12.filter(String).map(c => String(c || "").replace(/^'/, "")).join("");
      
      if (jsonNuevaSeccionRaw) {
        let registrosAnteriores = JSON.parse(jsonNuevaSeccionRaw);
        
        // Verificamos que ya tenga el formato de objeto nuevo
        if (registrosAnteriores && typeof registrosAnteriores === 'object' && !Array.isArray(registrosAnteriores)) {
            viajesDetalleObj = registrosAnteriores;
            
            // 👉 MAGIA PURA: Borramos de la memoria todo lo que caiga en la ventana temporal actual.
            for (let chofer in viajesDetalleObj) {
                for (let fecha in viajesDetalleObj[chofer]) {
                    let fReg = new Date(fecha + "T12:00:00");
                    if (fReg >= minDate && fReg <= maxDate) {
                        delete viajesDetalleObj[chofer][fecha]; 
                    }
                }
                // Limpieza de choferes que se quedaron sin viajes en el historial
                if (Object.keys(viajesDetalleObj[chofer]).length === 0) {
                    delete viajesDetalleObj[chofer];
                }
            }
        }
      }
    } catch(e) { console.warn("No se pudo realizar merge de la fila 12 de viajes detallados."); }
  }

  // --- 2. LECTURA DE LA HOJA DE KILÓMETROS MAESTRA ---
  let dataKm;
  try {
    const ssKm = SpreadsheetApp.openById(ID_SHEET_KILOMETROS);
    const sheetKm = ssKm.getSheetByName('KM') || ssKm.getSheets()[0]; 
    dataKm = sheetKm.getDataRange().getValues();
  } catch(e) { 
    console.error("Error leyendo planilla de KM:", e); 
    return; 
  }

  const parseNumberSafe = (val) => {
      if (typeof val === 'number') return val;
      let limpio = String(val || '').replace(/,/g, '.').replace(/[^0-9.-]/g, '');
      return parseFloat(limpio) || 0;
  };
  
  const parseSafeDateLocal = (rawDate) => {
     if (rawDate instanceof Date) return rawDate;
     let strDate = String(rawDate).split('-')[0].trim();
     let parts = strDate.split(/[\/\-]/);
     if (parts.length >= 3) {
        if(parts[0].length <= 2 && parts[2].length >= 2) { 
          let anio = parts[2].length === 2 ? "20" + parts[2] : parts[2];
          return new Date(anio, parseInt(parts[1], 10) - 1, parts[0]);
        } else if (parts[0].length === 4) { 
          return new Date(parts[0], parseInt(parts[1], 10) - 1, parts[2]);
        }
     }
     return new Date(strDate); 
  };

  const normalizarNombre = (n) => String(n).trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, ' ');
  const mesesAbrev = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];

  let contadorExtraidos = 0;

  // --- 3. BUCLE MAESTRO ---
  for (let i = 1; i < dataKm.length; i++) {
    let row = dataKm[i];

    let dominioRaw  = row[0];   // Col A
    let fechaRaw    = row[1];   // Col B
    let nombreRaw   = row[2];   // Col C
    let livianoRaw  = row[3];   // Col D
    let euroRaw     = row[4];   // Col E
    let campoRaw    = row[5];   // Col F
    let infiniaDRaw = row[7];   // Col H
    let kmBackupRaw = row[8];   // Col I 
    let kmBaseRaw   = row[16];  // Col Q 
    let hojaRutaRaw = row[19];  // Col T

    let nOriginal = String(nombreRaw).trim().toLowerCase(); 
    if (!nOriginal) continue;
    
    let nombreNorm = normalizarNombre(nOriginal); 
    let dObj = parseSafeDateLocal(fechaRaw);
    
    if (!dObj || isNaN(dObj.getTime())) continue;

    // Solo extraemos si la fecha entra en la ventana
    if (dObj >= minDate && dObj <= maxDate) {
      
      // A) KILÓMETROS GLOBALES
      let km = parseFloat(kmBaseRaw);
      if (isNaN(km) || km === 0) km = parseFloat(kmBackupRaw) || 0;        
      
      if (km > 0) {
        let dd = String(dObj.getDate()).padStart(2, '0');
        let mm = String(dObj.getMonth() + 1).padStart(2, '0');
        let yy = String(dObj.getFullYear()).slice(-2);
        let datePart = `${dd}/${mm}/${yy}`;
        
        if (!mapaKms[nOriginal]) mapaKms[nOriginal] = [];
        mapaKms[nOriginal].push({ fechaCorta: datePart, km: km }); 
      }

      // B) VIAJES CAMPO Y HOJAS DE RUTA DETALLADOS
      let livianoNum  = parseNumberSafe(livianoRaw);
      let euroNum     = parseNumberSafe(euroRaw);
      let campoNum    = parseNumberSafe(campoRaw);
      let infiniaDNum = parseNumberSafe(infiniaDRaw);
      let hojaStr     = String(hojaRutaRaw || "").trim();

      if (campoNum > 0 || livianoNum > 0 || euroNum > 0 || infiniaDNum > 0 || hojaStr !== "") {
          
          let mesAnio = `${mesesAbrev[dObj.getMonth()]}-${String(dObj.getFullYear()).slice(-2)}`;
          if (!viajesAgrupados[nombreNorm]) viajesAgrupados[nombreNorm] = {};
          if (!viajesAgrupados[nombreNorm][mesAnio]) {
              viajesAgrupados[nombreNorm][mesAnio] = { km: 0, liviano: 0, euro: 0, infiniaD: 0 };
          }
          viajesAgrupados[nombreNorm][mesAnio].km += campoNum;
          viajesAgrupados[nombreNorm][mesAnio].liviano += livianoNum;
          viajesAgrupados[nombreNorm][mesAnio].euro += euroNum;
          viajesAgrupados[nombreNorm][mesAnio].infiniaD += infiniaDNum;

          let isoStr = dObj.toISOString().split('T')[0];

          // 👉 NUEVA SECCIÓN: Asignación limpia al Diccionario
          if (!viajesDetalleObj[nombreNorm]) viajesDetalleObj[nombreNorm] = {};
          
          // Si hay múltiples filas en la planilla para un chofer el mismo día, se fusionan
          if (!viajesDetalleObj[nombreNorm][isoStr]) {
              viajesDetalleObj[nombreNorm][isoStr] = {
                  dominio: String(dominioRaw || "").trim(),
                  liviano: 0, euro: 0, campo: 0, infiniaD: 0,
                  hoja_ruta: []
              };
          }
          
          let target = viajesDetalleObj[nombreNorm][isoStr];
          target.liviano += livianoNum;
          target.euro += euroNum;
          target.campo += campoNum;
          target.infiniaD += infiniaDNum;

          if (hojaStr !== "") {
              let arrHojas = hojaStr.split(',').map(s => s.trim()).filter(Boolean);
              arrHojas.forEach(h => {
                  if (!target.hoja_ruta.includes(h)) target.hoja_ruta.push(h);
              });
          }
          contadorExtraidos++;
      }
    }
  }

  // --- 4. ESCRITURA ATÓMICA EN CACHÉ ---
  let hojaKm = ssMaestro.getSheetByName('api_km');
  if (!hojaKm) { hojaKm = ssMaestro.insertSheet('api_km'); hojaKm.hideSheet(); }
  hojaKm.clearContents(); 
  
  let kmChunks = [];
  let kmStr = JSON.stringify(mapaKms);
  for (let i = 0; i < kmStr.length; i += 40000) { kmChunks.push(["'" + kmStr.substring(i, i + 40000)]); }
  if (kmChunks.length > 0) hojaKm.getRange(1, 1, kmChunks.length, 1).setValues(kmChunks);

  if (cacheSheet && typeof escribirChunksEnFila === 'function') {
    escribirChunksEnFila(cacheSheet, 7, JSON.stringify(viajesAgrupados));
    escribirChunksEnFila(cacheSheet, 12, JSON.stringify(viajesDetalleObj)); // 👉 Guardamos el Objeto Limpio
  }
  
  ssMaestro.toast(`Proceso OK. ${contadorExtraidos} viajes detallados cacheados.`, "✅ Éxito");
}

function generarJSONKilometros_Frecuente() {
  const hoy = new Date(); hoy.setHours(23, 59, 59, 999);
  const hace60Dias = new Date(hoy); hace60Dias.setDate(hoy.getDate() - 60); hace60Dias.setHours(0, 0, 0, 0);
  procesarKilometrosYViajesCore(hace60Dias, hoy, true); 
}

function generarJSONKilometros_Completo() {
  const hoy = new Date(); hoy.setHours(23, 59, 59, 999);
  const hace1Ano = new Date(hoy); hace1Ano.setFullYear(hoy.getFullYear() - 1); hace1Ano.setHours(0, 0, 0, 0);
  procesarKilometrosYViajesCore(hace1Ano, hoy, false);
}

function ubmkm() {
  generarJSONKilometros_Completo();
}

/**
 * ============================================================================
 * EXTRACCIÓN QUIRÚRGICA: onEdit PARA HOJA DE RUTA
 * ============================================================================
 */
function alEditarKilometros(e) {
  if (!e || !e.range) return;

  const fila = e.range.getRow();
  const columna = e.range.getColumn();
  const sheet = e.source.getActiveSheet();
  const nombreHoja = sheet.getName();

  if (fila >= 2 && columna === 20 && nombreHoja.toUpperCase() === 'KM') {
    try {
      const ID_PLANILLA_MAESTRA = '1eQ9Y5diL5fwxYTxvseNgZJFbX-lSUQ13axbp3cLiqPc'; 
      
      let ssMaestro;
      try { ssMaestro = SpreadsheetApp.openById(ID_PLANILLA_MAESTRA); } 
      catch (err) { ssMaestro = SpreadsheetApp.getActiveSpreadsheet(); }

      const cacheSheet = ssMaestro.getSheetByName("API_CACHE_BASICO");
      if (!cacheSheet) return;

      const rangoFila = sheet.getRange(fila, 1, 1, 20).getValues()[0];
      const dominioRaw  = rangoFila[0];   
      const fechaRaw    = rangoFila[1];   
      const nombreRaw   = rangoFila[2];   
      const livianoNum  = parseFloat(String(rangoFila[3] || '').replace(/,/g, '.').replace(/[^0-9.-]/g, '')) || 0;
      const euroNum     = parseFloat(String(rangoFila[4] || '').replace(/,/g, '.').replace(/[^0-9.-]/g, '')) || 0;
      const campoNum    = parseFloat(String(rangoFila[5] || '').replace(/,/g, '.').replace(/[^0-9.-]/g, '')) || 0;
      const infiniaDNum = parseFloat(String(rangoFila[7] || '').replace(/,/g, '.').replace(/[^0-9.-]/g, '')) || 0;
      const nuevoValorHR = String(rangoFila[19] || "").trim();

      if (!nombreRaw || !fechaRaw) return;

      const normalizar = (n) => String(n).trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, ' ');
      const choferNorm = normalizar(nombreRaw);

      let fechaIso = "";
      if (fechaRaw instanceof Date) {
        let tempDate = new Date(fechaRaw.getTime() - (fechaRaw.getTimezoneOffset() * 60000));
        fechaIso = tempDate.toISOString().split('T')[0];
      } else {
        let parts = String(fechaRaw).split('-')[0].trim().split(/[\/\-]/);
        if (parts.length >= 3) {
          let aa = parts[2].length === 2 ? "20" + parts[2] : parts[2];
          fechaIso = `${aa}-${String(parts[1]).padStart(2,'0')}-${String(parts[0]).padStart(2,'0')}`;
        }
      }
      if (!fechaIso) return;

      if (cacheSheet.getMaxRows() < 12) cacheSheet.insertRowsAfter(cacheSheet.getMaxRows(), 12 - cacheSheet.getMaxRows());
      let lastCol = cacheSheet.getLastColumn() || 1;
      let dataFila12 = cacheSheet.getRange(12, 1, 1, lastCol).getValues()[0];
      let jsonRaw = dataFila12.filter(String).map(c => String(c || "").replace(/^'/, "")).join("");
      
      let viajesDetalleObj = {};
      if (jsonRaw) {
        try { 
            let parsed = JSON.parse(jsonRaw); 
            if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) viajesDetalleObj = parsed;
        } catch(err) {}
      }

      let hojasParseadas = nuevoValorHR !== "" ? nuevoValorHR.split(',').map(s => s.trim()).filter(Boolean) : [];

      // 👉 Inyectamos limpiamente como Diccionario
      if (!viajesDetalleObj[choferNorm]) viajesDetalleObj[choferNorm] = {};
      
      if (!viajesDetalleObj[choferNorm][fechaIso]) {
          viajesDetalleObj[choferNorm][fechaIso] = {
              dominio: String(dominioRaw || "").trim(),
              liviano: livianoNum,
              euro: euroNum,
              campo: campoNum,
              infiniaD: infiniaDNum,
              hoja_ruta: []
          };
      }
      viajesDetalleObj[choferNorm][fechaIso].hoja_ruta = hojasParseadas;

      if (typeof escribirChunksEnFila === 'function') {
        escribirChunksEnFila(cacheSheet, 12, JSON.stringify(viajesDetalleObj));
        sheet.getParent().toast(`Hoja Ruta [${hojasParseadas.join(',')}] sincronizada`, "⚡ Caché Actualizado");

        // 👇 EL GATILLO SEGURO Y BIEN UBICADO 👇
        if (typeof notificarBackendNode === 'function') {
            notificarBackendNode('KM');
        }
      }

    } catch (error) {
      console.error("Error crítico en inyección onEdit Fila 12:", error);
    }
  }
}

// ====================================================================
// LECTURA DIRECTA: EVITA CACHÉS Y FANTASMAS AL ELIMINAR ROWS
// ====================================================================
function obtenerViajesYHRDirecto() {
    try {
      const ID_SHEET_KILOMETROS = '1Wr-_P4mDvldif_cAx08sp7yT8uTUrajI2HQAJF6tnGM';
      const ssKm = SpreadsheetApp.openById(ID_SHEET_KILOMETROS);
      const sheetKm = ssKm.getSheetByName('KM') || ssKm.getSheets()[0];
      
      const dataKm = sheetKm.getDataRange().getValues();
      let viajesDetalleObj = {};
      
      const limiteDate = new Date();
      limiteDate.setDate(limiteDate.getDate() - 60);
      
      const normalizar = (n) => String(n).trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, ' ');
      
      for (let i = 1; i < dataKm.length; i++) {
        let row = dataKm[i];
        let fechaRaw = row[1];
        let nombreRaw = row[2];
        
        if (!fechaRaw || !nombreRaw) continue;
        
        let dObj;
        if (fechaRaw instanceof Date) { dObj = fechaRaw; } 
        else {
          let parts = String(fechaRaw).split('-')[0].trim().split(/[\/\-]/);
          if(parts.length >= 3) {
            let aa = parts[2].length === 2 ? "20" + parts[2] : parts[2];
            dObj = new Date(aa, parseInt(parts[1], 10) - 1, parts[0]);
          } else { dObj = new Date(fechaRaw); }
        }
        
        if (isNaN(dObj.getTime()) || dObj < limiteDate) continue;
        
        let choferNorm = normalizar(nombreRaw);
        let dominioRaw  = row[0];
        let livianoNum  = parseFloat(String(row[3] || '').replace(/,/g, '.').replace(/[^0-9.-]/g, '')) || 0;
        let euroNum     = parseFloat(String(row[4] || '').replace(/,/g, '.').replace(/[^0-9.-]/g, '')) || 0;
        let campoNum    = parseFloat(String(row[5] || '').replace(/,/g, '.').replace(/[^0-9.-]/g, '')) || 0;
        let infiniaDNum = parseFloat(String(row[7] || '').replace(/,/g, '.').replace(/[^0-9.-]/g, '')) || 0;
        let hojaStr     = String(row[19] || "").trim();
        
        if (campoNum > 0 || livianoNum > 0 || euroNum > 0 || infiniaDNum > 0 || hojaStr !== "") {
          let tempDate = new Date(dObj.getTime() - (dObj.getTimezoneOffset() * 60000));
          let fechaIso = tempDate.toISOString().split('T')[0];
          
          if (!viajesDetalleObj[choferNorm]) viajesDetalleObj[choferNorm] = {};
          if (!viajesDetalleObj[choferNorm][fechaIso]) {
            viajesDetalleObj[choferNorm][fechaIso] = {
              dominio: String(dominioRaw || "").trim(),
              liviano: 0, euro: 0, campo: 0, infiniaD: 0,
              hoja_ruta: []
            };
          }
          
          let target = viajesDetalleObj[choferNorm][fechaIso];
          target.liviano += livianoNum;
          target.euro += euroNum;
          target.campo += campoNum;
          target.infiniaD += infiniaDNum;
          
          if (hojaStr !== "") {
            let arrHojas = hojaStr.split(',').map(s => s.trim()).filter(Boolean);
            arrHojas.forEach(h => {
              if (!target.hoja_ruta.includes(h)) target.hoja_ruta.push(h);
            });
          }
        }
      }
      
      return JSON.stringify(viajesDetalleObj);
    } catch(e) {
      console.error("Error Lectura Directa HR:", e);
      return JSON.stringify({});
    }
}
