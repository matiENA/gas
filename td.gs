function updatetds() {

  actualizarCacheTDsCampo();
  actualizarCacheTDsInfiniaD();
  actualizarCacheTDsLivianoYEuro();
}

function actualizarCacheTDsCampo() {
  const ssPrincipal = SpreadsheetApp.getActiveSpreadsheet();
  const ID_SHEET_PROFORMAS = '1AWpQPA9q-rUYY0rI6Vn_SZ7vEANzBW5Pc6URQuoAk5g';
  const mesesLargos = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"];
  const hoy = new Date();
  const nombreMesActual = mesesLargos[hoy.getMonth()];
  const nombrePestana = "CAMPO G2 " + nombreMesActual;
  
  let mapaCampos = {};
  
  try {
    const ssProformas = SpreadsheetApp.openById(ID_SHEET_PROFORMAS);
    const sheetCampo = ssProformas.getSheetByName(nombrePestana);
    
    if (sheetCampo) {
      const data = sheetCampo.getDataRange().getValues();
      const normalizar = (str) => {
        if (!str) return "";
        return String(str).trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, ' ');
      };
      
      for (let i = 1; i < data.length; i++) {
        let fila = data[i];
        let rawFecha = fila[0];
        let rawNombre = fila[2];
        let rawTractor = fila[3];
        let td = fila[13];
        let avisoVacio = fila[17];
        
        if (!rawNombre || !td) continue; 
        
        let nKey = normalizar(rawNombre);
        let fechaStr = (rawFecha instanceof Date) 
          ? Utilities.formatDate(rawFecha, Session.getScriptTimeZone(), "dd/MM/yyyy") 
          : String(rawFecha || "").trim();
        
        if (!mapaCampos[nKey]) mapaCampos[nKey] = [];
        
        mapaCampos[nKey].push({
          fecha: fechaStr,
          tractor: String(rawTractor || "").trim(),
          td: String(td || "").trim(),
          avisoVacio: String(avisoVacio || "").trim()
        });
      }
    } else {
      ssPrincipal.toast(`⚠️ Pestaña ${nombrePestana} no encontrada`, "Aviso");
      return;
    }
    
    // --- 2. IMPRESIÓN Y CHUNKING (Fila 1) ---
    let jsonStr = JSON.stringify(mapaCampos);
    let chunks = [];
    const chunkSize = 45000; 
    
    for (let i = 0; i < jsonStr.length; i += chunkSize) {
      chunks.push("'" + jsonStr.substring(i, i + chunkSize));
    }
    
    let hojaCacheTD = ssPrincipal.getSheetByName('td');
    if (!hojaCacheTD) {
      hojaCacheTD = ssPrincipal.insertSheet('td');
      hojaCacheTD.hideSheet();
    }
    
    // CORRECCIÓN: Limpiar solo la Fila 1 en lugar de toda la hoja
    hojaCacheTD.getRange("1:1").clearContent(); 
    
    if (chunks.length > 0) {
      hojaCacheTD.getRange(1, 1, 1, chunks.length).setValues([chunks]);
    }
    
    ssPrincipal.toast(`Caché TDs Campo actualizada`, "✅ Éxito");
    
  } catch (error) {
    console.error("Error en actualizarCacheTDsCampo:", error);
    ssPrincipal.toast("Error: " + error.message, "❌ Error", -1);
  }
}

function actualizarCacheTDsInfiniaD() {
  const ssPrincipal = SpreadsheetApp.getActiveSpreadsheet();
  const ID_SHEET_PROFORMAS = '1AWpQPA9q-rUYY0rI6Vn_SZ7vEANzBW5Pc6URQuoAk5g';
  
  // Determinamos el nombre dinámico de la pestaña (Ej: "ABRIL INFINIA DIESEL")
  const mesesLargos = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"];
  const hoy = new Date();
  const nombreMesActual = mesesLargos[hoy.getMonth()];
  const nombrePestana = nombreMesActual + " INFINIA DIESEL"; 
  
  let mapaInfinia = {};
  
  try {
    const ssProformas = SpreadsheetApp.openById(ID_SHEET_PROFORMAS);
    const sheetInfinia = ssProformas.getSheetByName(nombrePestana);
    
    if (sheetInfinia) {
      const data = sheetInfinia.getDataRange().getValues();
      
      // Iteramos saltando la fila 0 (cabeceras)
      for (let i = 1; i < data.length; i++) {
        let fila = data[i];
        
        // Mapeo de columnas según tu especificación
        let rawFecha = fila[0];     // COL A: FECHA
        let rawDni = fila[1];       // COL B: DNI
        let td = fila[12];          // COL M: TD
        
        // Si no hay DNI o no hay TD, saltamos la fila
        if (!rawDni || !td) continue; 
        
        // Limpiamos el DNI (quitamos espacios extra)
        let dniLimpio = String(rawDni).trim();
        
        // Formatear Fecha (para que se guarde como string ligero "dd/MM/yyyy" y no como objeto Date)
        let fechaStr = "";
        if (rawFecha instanceof Date) {
          fechaStr = Utilities.formatDate(rawFecha, Session.getScriptTimeZone(), "dd/MM/yyyy");
        } else {
          fechaStr = String(rawFecha || "").trim();
        }
        
        // Inicializamos el array del chofer si no existe
        if (!mapaInfinia[dniLimpio]) {
          mapaInfinia[dniLimpio] = [];
        }
        
        // Agregamos el registro de Infinia al DNI correspondiente
        mapaInfinia[dniLimpio].push({
          fecha: fechaStr,
          td: String(td).trim()
        });
      }
    } else {
      console.warn(`Pestaña no encontrada: ${nombrePestana}`);
      ssPrincipal.toast(`⚠️ Pestaña ${nombrePestana} no encontrada`, "Aviso");
      return;
    }
    
    // --- 2. IMPRESIÓN Y CHUNKING (Pestaña "td" - Horizontal Fila 2) ---
    let jsonStr = JSON.stringify(mapaInfinia);
    let chunks = [];
    const chunkSize = 45000; // Límite seguro de caracteres
    
    for (let i = 0; i < jsonStr.length; i += chunkSize) {
      // Comilla simple para forzar formato texto en Google Sheets
      chunks.push("'" + jsonStr.substring(i, i + chunkSize));
    }
    
    // Buscamos o creamos la pestaña 'td'
    let hojaCacheTD = ssPrincipal.getSheetByName('td');
    if (!hojaCacheTD) {
      hojaCacheTD = ssPrincipal.insertSheet('td');
      hojaCacheTD.hideSheet();
    }
    
    // Limpiamos SOLO la FILA 2 (para no borrar los datos de "Campo" en la Fila 1)
    hojaCacheTD.getRange("2:2").clearContent();
    
    // Imprimimos horizontalmente empezando en A2
    if (chunks.length > 0) {
      let array2D = [chunks]; 
      hojaCacheTD.getRange(2, 1, 1, chunks.length).setValues(array2D);
    }
    
    ssPrincipal.toast(`Caché Infinia D actualizada (${nombrePestana})`, "✅ Éxito");
    
  } catch (error) {
    console.error("Error en actualizarCacheTDsInfiniaD:", error);
    ssPrincipal.toast("Error: " + error.message, "❌ Error", -1);
  }
}

function actualizarCacheTDsLivianoYEuro() {
  const ssPrincipal = SpreadsheetApp.getActiveSpreadsheet();
  const ID_SHEET_LIVEURO = '1AUw6FY5xOV9WXGyf1i-sgAPK52cWwJQx1S2yJYsfAxI';
  
  // Nombre de la pestaña (Si en el futuro le agregan el mes como a las otras, 
  // puedes cambiarlo a: const nombreMesActual + " Liviano & EURO")
  const nombrePestana = "Liviano & EURO";
  
  let mapaLiviano = {};
  let mapaEuro = {};
  
  try {
    const ssProformas = SpreadsheetApp.openById(ID_SHEET_LIVEURO);
    const sheetDatos = ssProformas.getSheetByName(nombrePestana);
    
    if (sheetDatos) {
      const data = sheetDatos.getDataRange().getValues();
      
      // Helper para normalizar nombres de choferes (igual que en Campos)
      const normalizar = (str) => {
        if (!str) return "";
        return String(str).trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, ' ');
      };
      
      // Iteramos saltando la fila 0 (cabeceras)
      for (let i = 1; i < data.length; i++) {
        let fila = data[i];
        
        let rawFecha = fila[0];     // COL A: FECHA
        let td = fila[4];           // COL E: TD
        let producto = fila[7];     // COL H: PRODUCTO
        let rawNombre = fila[9];    // COL J: CHOFER (Índice 9 en CSV)
        let avisoVacio = fila[10];  // COL K: AVISO DE VACIO
        
        // Si no hay Chofer o no hay TD, saltamos
        if (!rawNombre || !td) continue; 
        
        let nKey = normalizar(rawNombre);
        
        // Formatear Fecha
        let fechaStr = "";
        if (rawFecha instanceof Date) {
          fechaStr = Utilities.formatDate(rawFecha, Session.getScriptTimeZone(), "dd/MM/yyyy");
        } else {
          fechaStr = String(rawFecha || "").trim();
        }
        
        // Crear el objeto del viaje
        let viajeObj = {
          fecha: fechaStr,
          td: String(td || "").trim(),
          avisoVacio: String(avisoVacio || "").trim()
        };
        
        // Verificamos si es INFINIA DIESEL comprobando la columna H
        let esInfiniaDiesel = String(producto || "").toUpperCase().includes("405200 - INFINIA DIESEL");
        
        if (esInfiniaDiesel) {
          // --- LOGICA EURO ---
          if (!mapaEuro[nKey]) mapaEuro[nKey] = [];
          mapaEuro[nKey].push(viajeObj);
        } else {
          // --- LOGICA LIVIANO ---
          if (!mapaLiviano[nKey]) mapaLiviano[nKey] = [];
          mapaLiviano[nKey].push(viajeObj);
        }
      }
    } else {
      console.warn(`Pestaña no encontrada: ${nombrePestana}`);
      ssPrincipal.toast(`⚠️ Pestaña ${nombrePestana} no encontrada`, "Aviso");
      return;
    }
    
    // ====================================================================
    // --- IMPRESIÓN Y CHUNKING (Pestaña "td" - Horizontal Filas 3 y 4) ---
    // ====================================================================
    
    let hojaCacheTD = ssPrincipal.getSheetByName('td');
    if (!hojaCacheTD) {
      hojaCacheTD = ssPrincipal.insertSheet('td');
      hojaCacheTD.hideSheet();
    }
    
    const chunkSize = 45000; // Límite seguro de caracteres
    
    // --- 1. GUARDAR LIVIANO (FILA 3) ---
    hojaCacheTD.getRange("3:3").clearContent(); // Limpiar solo la fila 3
    let jsonLiviano = JSON.stringify(mapaLiviano);
    let chunksLiviano = [];
    for (let i = 0; i < jsonLiviano.length; i += chunkSize) {
      chunksLiviano.push("'" + jsonLiviano.substring(i, i + chunkSize));
    }
    if (chunksLiviano.length > 0) {
      hojaCacheTD.getRange(3, 1, 1, chunksLiviano.length).setValues([chunksLiviano]);
    }
    
    // --- 2. GUARDAR EURO (FILA 4) ---
    hojaCacheTD.getRange("4:4").clearContent(); // Limpiar solo la fila 4
    let jsonEuro = JSON.stringify(mapaEuro);
    let chunksEuro = [];
    for (let i = 0; i < jsonEuro.length; i += chunkSize) {
      chunksEuro.push("'" + jsonEuro.substring(i, i + chunkSize));
    }
    if (chunksEuro.length > 0) {
      hojaCacheTD.getRange(4, 1, 1, chunksEuro.length).setValues([chunksEuro]);
    }
    
    ssPrincipal.toast(`Caché Liviano y Euro actualizados con éxito.`, "✅ Éxito");
    
  } catch (error) {
    console.error("Error en actualizarCacheTDsLivianoYEuro:", error);
    ssPrincipal.toast("Error: " + error.message, "❌ Error", -1);
  }
}

function obtenerDatosTDParaFront() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaTD = ss.getSheetByName('td');
  if (!hojaTD) return { campo:{}, infinia:{}, liviano:{}, euro:{}, estados:{}, codigosExtra:{} };

  // 👉 LECTURA SEGURA: Evita el SyntaxError uniendo los chunks celda por celda
  const leerFila = (numFila) => {
    let values = hojaTD.getRange(numFila + ":" + numFila).getValues()[0];
    let jsonStr = values.filter(String).map(c => String(c).replace(/^'/, "")).join("");
    // Protegemos el código: Si el JSON está roto, devuelve vacío en lugar de colapsar la app
    try { return jsonStr ? JSON.parse(jsonStr) : {}; } catch(e) { return {}; }
  };

  return {
    campo: leerFila(1),
    infinia: leerFila(2),
    liviano: leerFila(3),
    euro: leerFila(4),
    estados: JSON.parse(hojaTD.getRange("A5").getValue() || "{}"),      // Checkboxes
    codigosExtra: JSON.parse(hojaTD.getRange("A6").getValue() || "{}") // Códigos 6 dígitos
  };
}

function guardarEstadoCheckboxTD(tdId, estado, codigosExtra, usuarioResponsable) {
      try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const hoja = ss.getSheetByName('td');
        if (!hoja) throw new Error("No se encontró la pestaña 'td'");

        // 1. Manejar Estados (Fila A5) - Ahora guarda el usuario y la fecha
        let estados = JSON.parse(hoja.getRange("A5").getValue() || "{}");
        estados[tdId] = {
          estado: estado,
          usuario: usuarioResponsable || "Desconocido",
          fecha: new Date().toISOString()
        };
        hoja.getRange("A5").setValue(JSON.stringify(estados));

        // 2. Manejar Códigos (Fila A6)
        let todosLosCodigos = JSON.parse(hoja.getRange("A6").getValue() || "{}");
        
        if (estado === true && codigosExtra && codigosExtra.length > 0) {
          todosLosCodigos[tdId] = codigosExtra;
        } else {
          delete todosLosCodigos[tdId];
        }
        hoja.getRange("A6").setValue(JSON.stringify(todosLosCodigos));

        // Registro en DB_Logs
        registrarLog(usuarioResponsable, "MODIFICO_ESTADO_TD", {
          tdId: tdId,
          checkbox_estado: estado ? "MARCADO" : "DESMARCADO",
          codigos_extra: codigosExtra || []
        });

        return { success: true };
      } catch (e) {
        return { success: false, error: e.message };
      }
}


function guardarDatosExtraBackend(tdId, valores, usuarioResponsable) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName('td');
    if (!hoja) throw new Error("No se encontró la pestaña 'td'");

    // 1. Leer estado ACTUAL del checkbox desde A5 (Soporta formato viejo y nuevo)
    let estados = JSON.parse(hoja.getRange("A5").getValue() || "{}");
    let estadoActualCheckbox = "DESMARCADO";
    if (estados[tdId]) {
      if (typeof estados[tdId] === 'object') {
        estadoActualCheckbox = estados[tdId].estado ? "MARCADO" : "DESMARCADO";
      } else {
        estadoActualCheckbox = estados[tdId] ? "MARCADO" : "DESMARCADO"; // Legacy (Boolean puro)
      }
    }

    // 2. Manejar Códigos (Fila A6)
    let todosLosCodigos = JSON.parse(hoja.getRange("A6").getValue() || "{}");
    if (valores && valores.length > 0) {
      todosLosCodigos[tdId] = valores;
    } else {
      delete todosLosCodigos[tdId];
    }
    hoja.getRange("A6").setValue(JSON.stringify(todosLosCodigos));

    // Registro en DB_Logs
    registrarLog(usuarioResponsable, "MODIFICO_EXTRA_TD", {
      tdId: tdId,
      checkbox_estado: estadoActualCheckbox,
      codigos_extra: valores || []
    });

    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}
