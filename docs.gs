function actualizarCacheG2_Vencimientos() {
  const ssMaestro = SpreadsheetApp.getActiveSpreadsheet();
  let hojaCache = ssMaestro.getSheetByName('API_CACHE_BASICO');
  if (!hojaCache) return;
  
  const hoy = new Date();
  let mapaDocs = {}, mapaHabs = {}, mapaCerts = {};

  // 👉 1. Crear mapa interno [DNI -> Nombre] EXCLUSIVAMENTE DESDE TAB 'dni'
  let mapaDniANombre = {};
  
  try {
    // ENRUTAMIENTO DIRECTO AL ID PROPORCIONADO
    const sheetDni = SpreadsheetApp.openById('1eQ9Y5diL5fwxYTxvseNgZJFbX-lSUQ13axbp3cLiqPc').getSheetByName('dni');
    
    if (sheetDni) {
      const dataDni = sheetDni.getDataRange().getValues();
      
      for (let i = 0; i < dataDni.length; i++) {
        let fila = dataDni[i];
        
        // --- LECTURA DEL BLOQUE IZQUIERDO (Columna A y C) ---
        let nIzq = fila[0] ? String(fila[0]).trim().toLowerCase() : ""; 
        let rawDniIzq = fila[2] ? String(fila[2]).replace(/\D/g, '') : "";
        if (nIzq && rawDniIzq) mapaDniANombre[String(parseInt(rawDniIzq, 10))] = nIzq;

        // --- LECTURA DEL BLOQUE DERECHO (Columna F y H - Índices 5 y 7) ---
        let nDer = fila[5] ? String(fila[5]).trim().toLowerCase() : "";
        let rawDniDer = fila[7] ? String(fila[7]).replace(/\D/g, '') : "";
        if (nDer && rawDniDer) mapaDniANombre[String(parseInt(rawDniDer, 10))] = nDer;
      }
    }
  } catch(e) { console.error("Error G2 tab DNI:", e); } 

  // 👉 Helper para extraer DNI del CUIL
  function extraerDniDeCuil(cuil) {
    let limpio = String(cuil).replace(/\D/g, '');
    if (!limpio) return "";
    if (limpio.length === 11) return String(parseInt(limpio.substring(2, 10), 10));
    if (limpio.length === 10) return String(parseInt(limpio.substring(2, 9), 10));
    return String(parseInt(limpio, 10)); 
  }

  // 👉 2. Módulo de DOCUMENTOS 
  try {
    const dataDocs = SpreadsheetApp.openById(ID_SHEET_DOCUMENTOS).getSheetByName('PERIODICOS').getDataRange().getValues();
    for (let i = 1; i < dataDocs.length; i++) {
      let nombreSheet = String(dataDocs[i][1]).trim().toLowerCase();
      if (!nombreSheet) continue;
      
      let cuil = dataDocs[i][4]; // Columna E
      let dniExtracted = extraerDniDeCuil(cuil);
      let n = mapaDniANombre[dniExtracted] || nombreSheet; // Enrutador
      
      let vVen = parseSafeDate(dataDocs[i][8]);
      let estado = 'OK';
      if (vVen) {
        let diff = Math.ceil((vVen - hoy) / (1000 * 60 * 60 * 24));
        estado = diff < 0 ? 'VENCIDO' : (diff <= 30 ? 'POR_VENCER' : 'VIGENTE');
      }
      mapaDocs[n] = { ven: toISODate(vVen), estado: estado };
    }
  } catch(e) { console.error("Error G2 Docs:", e); }

  // 👉 3. Módulo de HABILITACIONES 
  try {
    const dataHab = SpreadsheetApp.openById(ID_SHEET_HABILITACIONES).getSheetByName('VENCIMIENTOS').getDataRange().getValues();
    for (let i = 1; i < dataHab.length; i++) {
      let nombreSheet = String(dataHab[i][1]).trim().toLowerCase();
      if (!nombreSheet) continue;
      
      let rawDni = String(dataHab[i][2]).replace(/\D/g, ''); // Columna C
      let dniLimpio = rawDni ? String(parseInt(rawDni, 10)) : "";
      let n = mapaDniANombre[dniLimpio] || nombreSheet; // Enrutador
      
      let vLic = parseSafeDate(dataHab[i][4]); 
      let estadoLic = vLic ? (Math.ceil((vLic - hoy) / 86400000) < 0 ? 'VENCIDO' : (Math.ceil((vLic - hoy) / 86400000) <= 30 ? 'POR_VENCER' : 'VIGENTE')) : 'OK';
      let vCert = parseSafeDate(dataHab[i][3]);
      let estadoCert = vCert ? (Math.ceil((vCert - hoy) / 86400000) < 0 ? 'VENCIDO' : (Math.ceil((vCert - hoy) / 86400000) <= 30 ? 'POR_VENCER' : 'VIGENTE')) : 'OK';

      mapaHabs[n] = { ven: toISODate(vLic), estado: estadoLic };
      mapaCerts[n] = { ven: toISODate(vCert), estado: estadoCert };
    }
  } catch(e) { console.error("Error G2 Habs/Certs:", e); }

  escribirChunksEnFila(hojaCache, 2, JSON.stringify(mapaDocs));
  escribirChunksEnFila(hojaCache, 3, JSON.stringify(mapaHabs));
  escribirChunksEnFila(hojaCache, 5, JSON.stringify(mapaCerts));
  ssMaestro.toast("Vencimientos actualizados (Desde tab DNI)", "G2 OK");
}
function generarJSONAptosMedicos() {
  const ssMaestro = SpreadsheetApp.getActiveSpreadsheet();
  let mapaAptos = {};
  
  try {
    const ssAptos = SpreadsheetApp.openById(ID_SHEET_APTOS_MEDICOS);
    const sheetAptos = ssAptos.getSheetByName('Seguimiento Avalados Mensual');
    if (!sheetAptos) throw new Error("No se encontró la pestaña 'Seguimiento Avalados Mensual'");
    
    // Leemos toda la hoja
    const dataAptos = sheetAptos.getDataRange().getDisplayValues(); 
    if (dataAptos.length === 0) return;

    const headers = dataAptos[0];

    // 👉 1. BÚSQUEDA INTELIGENTE DE LA COLUMNA DE HOY
    const hoy = new Date();
    const d = String(hoy.getDate()).padStart(2, '0');
    const m = String(hoy.getMonth() + 1).padStart(2, '0');
    const y = hoy.getFullYear();
    const mesesLargo = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];
    
    // Array con los posibles formatos que pueden haber escrito en el encabezado
    const formatosHoy = [
      `${hoy.getDate()}/${mesesLargo[hoy.getMonth()]}/${y}`.toLowerCase(), // 21/abril/2026
      `${d}/${m}/${y}`, // 21/04/2026
      `${d}/${m}`,      // 21/04
      String(hoy.getDate()) // 21
    ];

    let colDiaria = -1;
    for (let c = 12; c < headers.length; c++) {
      let headLimpio = String(headers[c]).trim().toLowerCase();
      if (formatosHoy.includes(headLimpio)) {
        colDiaria = c;
        break;
      }
    }

    // Fallback: Si no encuentra la columna (ej: error de tipeo), toma la última columna con encabezado
    if (colDiaria === -1) {
      for (let c = headers.length - 1; c >= 12; c--) {
         if (String(headers[c]).trim() !== "") {
             colDiaria = c;
             break;
         }
      }
    }

    // 👉 Helper interno para extraer DNI desde CUIL
    function extraerDniDeCuil(cuil) {
        let limpio = String(cuil).replace(/\D/g, '');
        if (!limpio) return "";
        if (limpio.length === 11) return String(parseInt(limpio.substring(2, 10), 10));
        if (limpio.length === 10) return String(parseInt(limpio.substring(2, 9), 10));
        return String(parseInt(limpio, 10));
    }

    // 👉 2. MAPEO ANCLADO AL DNI CON BÚSQUEDA DEL ÚLTIMO ESTADO VÁLIDO
    for (let i = 1; i < dataAptos.length; i++) {
      let fila = dataAptos[i];
      let nombreRaw = String(fila[0]).trim(); // Columna A
      if (!nombreRaw || nombreRaw.toLowerCase() === "nombre completo") continue;
      
      let cuil = String(fila[1]).trim(); // Columna B
      let dniLimpio = extraerDniDeCuil(cuil);

      let responsable = String(fila[5]).trim();                // Col F (Índice 5)
      let observaciones = String(fila[10]).trim();             // Col K (Índice 10)
      let observaciones_salud = String(fila[11]).trim();       // Col L (Índice 11)
      
      // ----------------------------------------------------------------------
      // CORRECCIÓN: ESCÁNER HACIA ATRÁS (Evita los guiones o días en blanco)
      // ----------------------------------------------------------------------
      let estadoDiario = "-";
      let limiteBusqueda = colDiaria > -1 ? colDiaria : fila.length - 1;
      
      // Bucle que lee desde "Hoy" hacia el primer día del mes (Columna M / 12)
      for (let c = limiteBusqueda; c >= 12; c--) {
          let val = String(fila[c] || "").trim();
          // Si encuentra texto que no sea vacío ni guion, captura el estado y corta el bucle
          if (val !== "" && val !== "-") {
              estadoDiario = val; 
              break;
          }
      }

      // Normalizamos aplicando la misma regla estricta antifallos (sin tildes, sin espacios extra)
      let nombreNormalizado = nombreRaw.replace(/,/g, '').toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, ' ');

      let objApto = {
        dni: dniLimpio,
        cuil: cuil,
        estado: estadoDiario,  // Ahora garantiza devolver el último estado válido registrado
        responsable: responsable,
        observaciones: observaciones,
        observaciones_sector_salud: observaciones_salud
      };

      // Anclaje Estricto a DNI
      if (dniLimpio) {
          mapaAptos[dniLimpio] = objApto;
      }
      
      // Anclaje Secundario a Nombre
      if (nombreNormalizado) {
          mapaAptos[nombreNormalizado] = objApto;
      }
    }
    
    let hojaObsCache = ssMaestro.getSheetByName('OBSERVACIONES');
    if (!hojaObsCache) { hojaObsCache = ssMaestro.insertSheet('OBSERVACIONES'); hojaObsCache.hideSheet(); }
    
    // Limpia solo la fila 2 (Dedicada a Aptos Médicos)
    escribirChunksEnFila(hojaObsCache, 2, JSON.stringify(mapaAptos));
    ssMaestro.toast("Aptos Médicos diarios actualizados.", "Aptos OK");
    
  } catch (e) {
    console.error("Error generando JSON Aptos Médicos:", e);
    ssMaestro.toast("Error al generar Aptos Médicos", "Error");
  }
}
function generarJSONObservacionesGlobal() {
  const ssMaestro = SpreadsheetApp.getActiveSpreadsheet();
  let mapaObservaciones = {};
  
  try {
    const ssObs = SpreadsheetApp.openById(ID_SHEET_OBSERVACIONES);
    const sheetMov = ssObs.getSheetByName('Movimientos');
    if (!sheetMov) return;
    
    const data = sheetMov.getDataRange().getDisplayValues();
    for (let i = 4; i < data.length; i++) {
      let fila = data[i];
      if (!fila || fila.length < 8) continue;
      
      let nombre = String(fila[1]).trim().toLowerCase();
      if (!nombre) continue;
      
      if (!mapaObservaciones[nombre]) mapaObservaciones[nombre] = [];
      
      mapaObservaciones[nombre].push({
        admin: fila[0] || "-", fecha: fila[2] || "-", unidad: fila[3] || "-",
        evento: fila[4] || "-", obsEvento: fila[5] || "", estado: fila[6] || "-", obsEstado: fila[7] || ""
      });
    }
    
    let hojaObsCache = ssMaestro.getSheetByName('OBSERVACIONES');
    if (!hojaObsCache) { hojaObsCache = ssMaestro.insertSheet('OBSERVACIONES'); hojaObsCache.hideSheet(); }
    
    // Limpia solo la fila 1
    escribirChunksEnFila(hojaObsCache, 1, JSON.stringify(mapaObservaciones));
    
  } catch(e) { console.error("Error generando JSON de Observaciones:", e); }
}
function guardarDocumentos(nombre, exVen, licVen, certVen) {
  try {
    let nBuscado = String(nombre).trim().toLowerCase();
    
    // 1. Identificamos el DNI real EXCLUSIVAMENTE desde la tab 'dni'
    let dniBuscado = "";
    try {
        const sheetDni = SpreadsheetApp.openById('1eQ9Y5diL5fwxYTxvseNgZJFbX-lSUQ13axbp3cLiqPc').getSheetByName('dni');
        
        if (sheetDni) {
            const dataDni = sheetDni.getDataRange().getValues();
            for(let i = 0; i < dataDni.length; i++) {
                let fila = dataDni[i];
                // Revisar lado izquierdo
                if(fila[0] && String(fila[0]).trim().toLowerCase() === nBuscado) {
                    let raw = String(fila[2]).replace(/\D/g, '');
                    if(raw) { dniBuscado = String(parseInt(raw, 10)); break; }
                }
                // Revisar lado derecho
                if(fila[5] && String(fila[5]).trim().toLowerCase() === nBuscado) {
                    let raw = String(fila[7]).replace(/\D/g, '');
                    if(raw) { dniBuscado = String(parseInt(raw, 10)); break; }
                }
            }
        }
    } catch(e) { console.error("Error leyendo DNI para guardar:", e); }

    function extraerDniDeCuil(cuil) {
        let limpio = String(cuil).replace(/\D/g, '');
        if (!limpio) return "";
        if (limpio.length === 11) return String(parseInt(limpio.substring(2, 10), 10));
        if (limpio.length === 10) return String(parseInt(limpio.substring(2, 9), 10));
        return String(parseInt(limpio, 10));
    }

    // 2. Buscar y guardar en HABILITACIONES
    const ssHab = SpreadsheetApp.openById(ID_SHEET_HABILITACIONES);
    const sheetHab = ssHab.getSheetByName('VENCIMIENTOS'); 
    const dataHab = sheetHab.getRange("B:C").getValues();
    let filaHab = -1;
    for(let i=0; i<dataHab.length; i++) {
        let nSheet = String(dataHab[i][0]).trim().toLowerCase();
        let dniSheet = String(dataHab[i][1]).replace(/\D/g, '');
        dniSheet = dniSheet ? String(parseInt(dniSheet, 10)) : "";
        
        if( (dniBuscado && dniSheet === dniBuscado) || (nSheet === nBuscado) ) { 
            filaHab = i + 1; 
            break; 
        }
    }
    
    if(filaHab !== -1) {
        if(licVen !== undefined && licVen !== null) sheetHab.getRange(filaHab, 5).setValue(licVen); 
        if(certVen !== undefined && certVen !== null) sheetHab.getRange(filaHab, 4).setValue(certVen); 
    }

    // 3. Buscar y guardar en DOCUMENTOS
    const ssDoc = SpreadsheetApp.openById(ID_SHEET_DOCUMENTOS);
    const sheetDoc = ssDoc.getSheetByName('PERIODICOS'); 
    const dataDoc = sheetDoc.getRange("B:E").getValues(); 
    let filaDoc = -1;
    for(let i=0; i<dataDoc.length; i++) {
        let nSheet = String(dataDoc[i][0]).trim().toLowerCase();
        let cuilSheet = dataDoc[i][3]; 
        let dniSheet = extraerDniDeCuil(cuilSheet);
        
        if( (dniBuscado && dniSheet === dniBuscado) || (nSheet === nBuscado) ) { 
            filaDoc = i + 1; 
            break; 
        }
    }
    
    if(filaDoc !== -1) {
        if(exVen !== undefined && exVen !== null) sheetDoc.getRange(filaDoc, 9).setValue(exVen); 
    }

    actualizarCacheG2_Vencimientos(); 
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function actualizarCacheG4_Fotos() {
  const ssMaestro = SpreadsheetApp.getActiveSpreadsheet();
  const hojaFotos = ssMaestro.getSheetByName('fotos'); 
  let hojaCache = ssMaestro.getSheetByName('API_CACHE_BASICO');
  
  if (!hojaFotos || !hojaCache) return;

  // Mismo helper que en G3 para extraer el DNI exacto a partir del CUIL
  function _parseDniExacto(val) {
    let limpio = String(val).replace(/\D/g, '');
    if (!limpio) return "";
    if (limpio.length === 11) return String(parseInt(limpio.substring(2, 10), 10));
    if (limpio.length === 10) return String(parseInt(limpio.substring(2, 9), 10));
    return String(parseInt(limpio, 10));
  }

  let mapaFotos = {};
  const data = hojaFotos.getDataRange().getValues();
  
  for (let i = 0; i < data.length; i++) {
    let dniExacto = _parseDniExacto(data[i][0]);
    let urlImgur = String(data[i][1]).trim();
    
    // Solo guardamos si tenemos un DNI válido y un link real
    if (dniExacto && urlImgur && urlImgur.includes('http')) {
      mapaFotos[dniExacto] = urlImgur;
    }
  }

  // Guardamos en la Fila 10 (Índice 9)
  escribirChunksEnFila(hojaCache, 10, JSON.stringify(mapaFotos));
  ssMaestro.toast("Caché de Fotos actualizado correctamente.", "G4 OK");
}
function updatedocs() {
    actualizarCacheG2_Vencimientos();
    actualizarCacheG3_Estaticos();
    generarJSONObservacionesGlobal();
    generarJSONAptosMedicos();
    actualizarCacheG4_Fotos()
  }

