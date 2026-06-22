/**
 * ============================================================================
 * DIAGRAMAS EOR - BACKEND (ARQUITECTURA MODULAR G1, G2, G3, G4)
 * ============================================================================
 */

const ID_SHEET_MOVIMIENTOS = '1hhJKwp9xOOHL_zZSJMbrJh5fwfsIPre155UTWhKWI44'; 
const NOMBRE_PESTANA_MOVIMIENTOS = 'JUNIO 2026- Mov.Unidades y Choferes'; 
const ID_SPREADSHEET_DIAGRAMAS = '1mhfXpFCF6upMlnRnZjDdBVS_wqTx5q8v0qQArNCnNAU'; 
const ID_SHEET_OBSERVACIONES = '1VwCNK89ecaac7IDlMWWCLHRqZoch9HB6vop5AfQEaA0';
const ID_SHEET_KILOMETROS = '1Wr-_P4mDvldif_cAx08sp7yT8uTUrajI2HQAJF6tnGM';
const ID_SHEET_DOCUMENTOS = '1pnYXKDSv70Vq78Rchxus5FHMKdgXdbfltVsEg6vArjo';
const ID_SHEET_HABILITACIONES = '1hPDno09tMBtKh7aIdsvzEYcyOY7leYj2B6XnniD0aXg';
const ID_SHEET_APTOS_MEDICOS = '1oJmN8hurfHfNnGBYUFcBdlrIj2VUzeIyq0ZTWxTpYNI';
const ID_SHEET_TELEFONOS = '1xJ31YzNbX8BwO3t3zLSm3GdRyM_E7_TImKpD7plBNoU';


function doGet(e) {
  // 👉 1. Endpoint para Diagramas Básicos
  if (e && e.parameter && e.parameter.action === 'obtenerDiagramasCacheados') {
    let jsonString = obtenerDiagramasCacheados(); 
    return ContentService.createTextOutput(jsonString).setMimeType(ContentService.MimeType.JSON);
  }
  
  // 👉 2. Endpoint para descargar los TDs (Actualizado para incluir códigos)
  if (e && e.parameter && e.parameter.action === 'obtenerTDs') {
    let data = obtenerDatosTDParaFront();
    return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
  }

  // Agregar esto junto a los otros "if" de tu doGet:
  if (e && e.parameter && e.parameter.action === 'obtenerViajesYHRDirecto') {
    return ContentService.createTextOutput(obtenerViajesYHRDirecto()).setMimeType(ContentService.MimeType.JSON);
  }

// dni extractor (alt)
  if (e && e.parameter && e.parameter.action === 'obtenerNombresMesActual') {
    return ContentService.createTextOutput(obtenerNombresMesActual()).setMimeType(ContentService.MimeType.JSON);
  }
  
  // 👉 Fallback clásico
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Gestor de Diagramas EOR')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  let respuesta = { success: false, error: "Acción no reconocida o payload vacío" };
  
  try {
    let payload = JSON.parse(e.postData.contents);
    
    if (payload.action === 'actualizarEstado') {
      respuesta = actualizarEstado(payload.nombre, payload.startIso, payload.endIso, payload.est, payload.usuario);
      
    } else if (payload.action === 'guardarDocumentos') {
      respuesta = guardarDocumentos(payload.nombre, payload.exVen, payload.licVen, payload.certVen);
      
    // 👉 NUEVO: Endpoint Bidireccional de Hojas de Ruta
    } else if (payload.action === 'guardarHojaRutaPlanilla') { 
      respuesta = procesarGuardadoHojaRuta(payload.nombre, payload.fecha, payload.hojas, payload.usuario, payload.tractor); 

    } else if (payload.action === 'sincronizarTotal') {
      respuesta = ejecutarSincronizacionTotal();
      
    } else if (payload.action === 'login') {
      respuesta = validarLogin(payload.usuario, payload.password);
    }

  } catch (error) {
    respuesta = { success: false, error: error.message };
  }
  
  return ContentService.createTextOutput(JSON.stringify(respuesta)).setMimeType(ContentService.MimeType.JSON);
}

function validarLogin(usuarioLimpio, passwordLimpio) {
  // Conexión al archivo maestro donde viven las bases de datos
  const ID_MASTER = '1eQ9Y5diL5fwxYTxvseNgZJFbX-lSUQ13axbp3cLiqPc';
  
  try {
    const ssMaster = SpreadsheetApp.openById(ID_MASTER);
    const hojaUsuarios = ssMaster.getSheetByName('DB_Usuarios');
    
    if (!hojaUsuarios) {
      return { success: false, error: "Error de configuración: No existe DB_Usuarios." };
    }

    const data = hojaUsuarios.getDataRange().getValues();
    
    // Ignoramos la fila 1 (Cabeceras)
    for (let i = 1; i < data.length; i++) {
      let userDb = String(data[i][0]).trim();
      let passDb = String(data[i][1]).trim();
      let rolDb = String(data[i][2]).trim();

      if (userDb === usuarioLimpio && passDb === passwordLimpio) {
        // Generamos un "token" muy básico y opaco para el frontend
        let token = Utilities.base64Encode(userDb + "_" + new Date().getTime());
        
        return { 
          success: true, 
          token: token, 
          usuario: userDb,
          rol: rolDb 
        };
      }
    }
    
    return { success: false, error: "Credenciales inválidas." };
  } catch (error) {
    return { success: false, error: "Error interno del servidor." };
  }
}

function obtenerDiagramasCacheados() {
  const ssMaestro = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. GESTALT: PRÄGNANZ & PROXIMITY
  // We define the final unified object upfront. This acts as our single source of truth 
  // and establishes default fallbacks (empty arrays/objects) preventing UI crashes.
  const payload = {
    diagramas: [],
    documentos: {},
    habilitaciones: {},
    dnis: {},
    certificados: {},
    telefonos: {},
    viajesCampo: {},
    infiniaProformas: {},
    entregasInfinia: {},
    fotosImgur: {},
    vencimientosObj: {}, // <-- Integrated Expirations Pipeline (Row 11)
    nuevaSeccionViajes: [], // <-- NUEVO: Array para el detalle de viajes (Fila 12)
    kilometros: {},
    observaciones: {},
    aptosMedicos: {}
  };

  // 2. Caché Base (Filas de la 1 a la 12)
  try {
    const hojaCache = ssMaestro.getSheetByName('API_CACHE_BASICO');
    if (hojaCache && hojaCache.getLastRow() > 0) {
      const data = hojaCache.getDataRange().getValues();
      
      // Función auxiliar simplificada: Retorna un objeto parseado o null
      const parseSeguro = (filaIndex) => {
        if (filaIndex >= data.length) return null;
        const str = data[filaIndex].map(c => String(c||"").replace(/^'/,"")).join("");
        if (!str) return null;
        try { return JSON.parse(str); } catch(e) { return null; }
      };

      // Mapeo directo y conciso. Si parseSeguro falla, mantiene el fallback por defecto del payload.
      payload.diagramas        = parseSeguro(0)  || payload.diagramas;
      payload.documentos       = parseSeguro(1)  || payload.documentos;
      payload.habilitaciones   = parseSeguro(2)  || payload.habilitaciones;
      payload.dnis             = parseSeguro(3)  || payload.dnis;
      payload.certificados     = parseSeguro(4)  || payload.certificados;
      payload.telefonos        = parseSeguro(5)  || payload.telefonos;
      payload.viajesCampo      = parseSeguro(6)  || payload.viajesCampo;
      payload.infiniaProformas = parseSeguro(7)  || payload.infiniaProformas;
      payload.entregasInfinia  = parseSeguro(8)  || payload.entregasInfinia;
      payload.fotosImgur       = parseSeguro(9)  || payload.fotosImgur;
      payload.vencimientosObj  = parseSeguro(10) || payload.vencimientosObj; // <-- Fila 11
      payload.nuevaSeccionViajes = parseSeguro(11) || payload.nuevaSeccionViajes; // <-- NUEVO: Fila 12
    }
  } catch (e) { 
    console.error("Error leyendo caché base:", e); 
  }
  
  // 3. Caché de Kilómetros
  try {
    const hojaKm = ssMaestro.getSheetByName('api_km');
    if (hojaKm && hojaKm.getLastRow() > 0) {
      let kmStr = "";
      hojaKm.getDataRange().getValues().forEach(row => {
        row.forEach(cell => { if (cell) kmStr += String(cell).replace(/^'/, ""); });
      });
      if (kmStr) payload.kilometros = JSON.parse(kmStr);
    }
  } catch (e) { 
    console.error("Error leyendo caché KMs:", e); 
  }

  // 4. Caché de Observaciones y Aptos Médicos
  try {
    const hojaObs = ssMaestro.getSheetByName('OBSERVACIONES'); 
    if (hojaObs && hojaObs.getLastRow() > 0) {
      const dataObs = hojaObs.getDataRange().getValues();
      
      if (dataObs.length > 0 && dataObs[0][0]) {
        const obsStr = dataObs[0].map(c => String(c||"").replace(/^'/,"")).join("");
        if (obsStr) payload.observaciones = JSON.parse(obsStr);
      }
      
      if (dataObs.length > 1 && dataObs[1][0]) {
        const aptosStr = dataObs[1].map(c => String(c||"").replace(/^'/,"")).join("");
        if (aptosStr) payload.aptosMedicos = JSON.parse(aptosStr);
      }
    }
  } catch (e) { 
    console.error("Error leyendo caché Observaciones/Aptos:", e); 
  }

  // 5. Retornar el payload unificado
  return JSON.stringify(payload);
}
function parseSafeDate(val) {
  if (!val) return null;
  if (val instanceof Date) return val;
  let s = String(val).trim();
  if (s.match(/^\d{1,2}\/\d{1,2}\/\d{4}/)) {
    let parts = s.split('/');
    return new Date(parts[2], parts[1] - 1, parts[0]);
  }
  if (s.match(/^\d{4}-\d{2}-\d{2}/)) {
    let parts = s.split('-');
    return new Date(parts[0], parts[1] - 1, parts[2].substring(0,2));
  }
  let d = new Date(s);
  if (!isNaN(d.getTime())) return d;
  return null;
}
function toISODate(d) {
  if (!d || !(d instanceof Date) || isNaN(d)) return "";
  let tzOffset = d.getTimezoneOffset() * 60000;
  return new Date(d.getTime() - tzOffset).toISOString().split('T')[0];
}

function obtenerHojasObjetivo() {
  const hojas = [];
  const mesesAbrev = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];
  let hoy = new Date();
  for (let i = -1; i <= 3; i++) {
    let d = new Date(hoy.getFullYear(), hoy.getMonth() + i, 1);
    let yearSuffix = String(d.getFullYear()).slice(-2);
    let nombreHoja = mesesAbrev[d.getMonth()] + "-" + yearSuffix;
    hojas.push({ mes: d.getMonth(), anio: d.getFullYear(), nombre: nombreHoja, fechaBase: d });
  }
  return hojas;
}

function escribirChunksEnFila(hoja, fila, stringData) {
  let chunkSize = 45000; // <--- Ajustado a 45k
  let chunks = [];
  for (let i = 0; i < stringData.length; i += chunkSize) {
    chunks.push("'" + stringData.substring(i, i + chunkSize));
  }
  
  // LIMPIA SOLO ESTA FILA antes de escribir
  hoja.getRange(fila + ":" + fila).clearContent();
  
  if (chunks.length > 0) {
    hoja.getRange(fila, 1, 1, chunks.length).setValues([chunks]);
  }
}

function generarJSONBase_Frecuente() {
  const hoy = new Date();
  const haceUnaSemana = new Date(hoy);
  haceUnaSemana.setDate(hoy.getDate() - 7);
  let offsets = [];
  if (haceUnaSemana.getMonth() !== hoy.getMonth()) { offsets.push(-1); }
  offsets.push(0, 1, 2);
  procesarVentanaDiagramas(offsets, true);
}

function generarJSONBase_Completo() {
  const hoy = new Date();
  const mesActual = hoy.getMonth(); 
  let offsetsAnuales = [];
  for (let i = -mesActual; i <= (11 - mesActual); i++) { offsetsAnuales.push(i); }
  procesarVentanaDiagramas(offsetsAnuales, false);
}

function generarJSONBase_Diario() {
  let ventanaAnual = [];
  for(let i = -6; i <= 6; i++) ventanaAnual.push(i);
  procesarVentanaDiagramas(ventanaAnual);
}

function procesarVentanaDiagramas(offsetsMeses, hacerMerge = false) {
  // Enrutamiento duro: Siempre apuntará a tu base de datos principal, sin importar desde dónde se dispare el trigger
  const ssMaestro = SpreadsheetApp.openById('1eQ9Y5diL5fwxYTxvseNgZJFbX-lSUQ13axbp3cLiqPc'); 
  const hoy = new Date();
  
  // ... (el resto del código queda exactamente igual)
  
  let hojaCache = ssMaestro.getSheetByName('API_CACHE_BASICO');
  if (!hojaCache) { hojaCache = ssMaestro.insertSheet('API_CACHE_BASICO'); hojaCache.hideSheet(); }
  
  let mapaChoferes = {}; 
  
  // 👉 1. Helper para Normalizar Nombres estricto (Evita bugs por espacios/tildes)
  const normalizarNombre = (n) => String(n).trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, ' ');

// 👉 2. Leer JSON de DNIs desde la FILA 4 (API_CACHE_BASICO) soportando chunks infinitos
  let mapaDni = {};
  try {
    // Leemos toda la fila 4 de una vez
    let dataFila4 = hojaCache.getRange("4:4").getValues()[0];
    
    // Filtramos las celdas vacías, quitamos la comilla inicial y unimos todos los pedazos
    let jsonDniRaw = dataFila4.filter(String).map(c => String(c || "").replace(/^'/, "")).join("");
    
    if (jsonDniRaw) {
      let dataDniObj = JSON.parse(jsonDniRaw);
      // Normalizamos las claves del diccionario DNI para que crucen perfecto
      for (let keyDni in dataDniObj) {
        mapaDni[normalizarNombre(keyDni)] = dataDniObj[keyDni];
      }
    }
  } catch (e) {
    console.warn("No se pudo leer/parsear el mapa de DNIs en Fila 4:", e);
  }
  const ssDiag = SpreadsheetApp.openById(ID_SPREADSHEET_DIAGRAMAS);
  const mesesAbrev = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];
  const nombreHojaActual = mesesAbrev[hoy.getMonth()] + "-" + String(hoy.getFullYear()).slice(-2);

  let choferesMesActual = new Set();
  try {
    const hojaActual = ssDiag.getSheetByName(nombreHojaActual);
    if (hojaActual) {
      let datosActual = hojaActual.getDataRange().getValues();
      // Leemos Nombres en Columna B (Índice 1)
      for (let i = 5; i < datosActual.length; i++) {
        let n = String(datosActual[i][1]).trim();
        if (n && n !== "APELLIDO Y NOMBRE" && n !== "Personal Activo") {
          choferesMesActual.add(normalizarNombre(n));
        }
      }
    }
  } catch(e) { console.error("Error leyendo hoja del mes actual:", e); }

  if (hacerMerge) {
    try {
      let dataCache = hojaCache.getRange(1, 1, 1, hojaCache.getLastColumn()).getValues();
      if (dataCache[0] && dataCache[0].length > 0) {
        let jsonString = dataCache[0].map(c => String(c||"").replace(/^'/,"")).join("");
        let arrCache = JSON.parse(jsonString);
        arrCache.forEach(ch => { 
          let key = normalizarNombre(ch.nom);
          if (choferesMesActual.has(key)) { mapaChoferes[key] = ch; }
        });
      }
    } catch(e) { console.warn("Iniciando caché G1 en blanco."); }
  }

  // B1. Leer JSON de Flota (Interno)
  let mapaFlota = {};
  try {
    const hojaUnidades = ssMaestro.getSheetByName('choferes y unidades');
    if (hojaUnidades) {
      let jsonStringFlota = hojaUnidades.getRange("H1").getValue();
      if (jsonStringFlota) {
        let flotaData = JSON.parse(jsonStringFlota);
        flotaData.forEach(ch => {
          let keyFlota = normalizarNombre(ch.nombre);
          // 💡 Agregamos la captura de los colores (soportando ambas nomenclaturas por las dudas)
          mapaFlota[keyFlota] = { 
            tractor: ch.tractor || "-", 
            semi: ch.semi || "-", 
            n_ute: ch.n_ute || "-", 
            td: ch.td || "-",
            hex1: ch.hex1 || ch.hex_1 || "", // <-- NUEVO
            hex2: ch.hex2 || ch.hex_2 || ""  // <-- NUEVO
          };
        });
      }
    }
  } catch(e) {}


  // B2. Leer Semáforo de Observaciones (Externo)
  let mapaAlertas = {};
  try {
    const ssObs = SpreadsheetApp.openById(ID_SHEET_OBSERVACIONES);
    const dataObs = ssObs.getSheetByName('Movimientos').getDataRange().getDisplayValues();
    for (let i = 4; i < dataObs.length; i++) {
      let nombreObs = normalizarNombre(dataObs[i][1]);
      let estadoObs = String(dataObs[i][6]).trim();
      if (nombreObs && estadoObs) {
        if (estadoObs !== "SITUACION COMUNICADA Y ACLARADA" && estadoObs !== "INFORMATIVO") {
          mapaAlertas[nombreObs] = estadoObs;
        } else {
          delete mapaAlertas[nombreObs]; 
        }
      }
    }
  } catch(e) {}

  // MERGE CON LA VENTANA DE DIAGRAMAS
  const hojasObjetivo = [];
  for (let i of offsetsMeses) { 
    let d = new Date(hoy.getFullYear(), hoy.getMonth() + i, 1);
    let nombreHoja = mesesAbrev[d.getMonth()] + "-" + String(d.getFullYear()).slice(-2);
    hojasObjetivo.push(nombreHoja);
  }

  for (let nombreHoja of hojasObjetivo) {
    let hojaDiag = ssDiag.getSheetByName(nombreHoja);
    if (!hojaDiag) continue;
    
    let datosDiag = hojaDiag.getDataRange().getValues();
    for (let i = 5; i < datosDiag.length; i++) {
      
      // ELIMINADO LEGAJO (Ya no leemos datosDiag[i][0])
      let nombreRaw = String(datosDiag[i][1]).trim();
      if (!nombreRaw || nombreRaw === "APELLIDO Y NOMBRE" || nombreRaw === "Personal Activo") continue;
      
      let key = normalizarNombre(nombreRaw);
      if (!choferesMesActual.has(key)) continue;
      
      let servicio = String(datosDiag[i][2]).trim();
      let diagrama = String(datosDiag[i][3]).trim();
      let tiraDias = "";
      for(let d = 4; d <= 34; d++) {
         tiraDias += (datosDiag[i][d] ? String(datosDiag[i][d]).trim() : "-") + ",";
      }
      tiraDias = tiraDias.slice(0, -1); 
      
      if (!mapaChoferes[key]) {
        mapaChoferes[key] = {
          "dni": mapaDni[key] || "", // EXTRAYENDO DNI DESDE EL MAPA DE A4
          "nom": nombreRaw,          // Conserva mayúsculas/minúsculas originales
          "srv": servicio, 
          "diag": diagrama,
          "tractor": "", "semi": "", "n_ute": "", "alerta_obs": "",
          "dias": {} 
        };
      }
      
      mapaChoferes[key].dias[nombreHoja] = tiraDias; 
      
      if (nombreHoja === nombreHojaActual) {
        mapaChoferes[key].srv = servicio;
        mapaChoferes[key].diag = diagrama;
      }
    }
  }

  for (let key in mapaChoferes) {
     // 💡 Sumamos hex1 y hex2 al fallback por defecto
     let infoFlota = mapaFlota[key] || { tractor: '-', semi: '-', n_ute: '-', td: '-', hex1: '', hex2: '' };
     mapaChoferes[key].tractor = infoFlota.tractor;
     mapaChoferes[key].semi = infoFlota.semi;
     mapaChoferes[key].n_ute = infoFlota.n_ute;
     mapaChoferes[key].td = infoFlota.td;
     
     // 💡 Traspasamos los colores al caché que lee el HTML
     mapaChoferes[key].hex1 = infoFlota.hex1; // <-- NUEVO
     mapaChoferes[key].hex2 = infoFlota.hex2; // <-- NUEVO
     
     mapaChoferes[key].alerta_obs = mapaAlertas[key] || "";
  }
  
  // GUARDAR SOLO EN FILA 1
  let jsonSalida = Object.values(mapaChoferes);
  escribirChunksEnFila(hojaCache, 1, JSON.stringify(jsonSalida));
  
  ssMaestro.toast(`Diagramas OK. (Sin Legajos, Usando Nombres y DNI)`, "G1 OK");
}

/**
 * Actualiza el estado en el diagrama y registra el log.
 * El parámetro usuarioActivo debe ir al final para evitar desplazar los datos originales.
 */
function actualizarEstado(nombreChofer, fechaInicioIso, fechaFinIso, nuevoEstado, usuarioActivo) {
  try {
    const ssDiag = SpreadsheetApp.openById(ID_SPREADSHEET_DIAGRAMAS);
    const mesesAbrev = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];
    
    let fIni = new Date(fechaInicioIso + "T12:00:00");
    let fFin = new Date(fechaFinIso + "T12:00:00");
    if (fIni > fFin) { let temp = fIni; fIni = fFin; fFin = temp; }

    let updatesPorHoja = {};
    let currentDate = new Date(fIni);
    let diaIndex = 0;

    // 1. Mapeo de días
    while (currentDate <= fFin) {
      let nombreHoja = mesesAbrev[currentDate.getMonth()] + "-" + String(currentDate.getFullYear()).slice(-2);
      let colEscribir = currentDate.getDate() + 4; 

      let valorBruto = Array.isArray(nuevoEstado) ? nuevoEstado[diaIndex] : nuevoEstado;
      let valorAEscribir = (valorBruto === 'BORRAR' || valorBruto === 'OPERATIVO') ? "" : valorBruto;

      if (!updatesPorHoja[nombreHoja]) {
        updatesPorHoja[nombreHoja] = { startCol: colEscribir, values: [] };
      }
      updatesPorHoja[nombreHoja].values.push(valorAEscribir);

      currentDate.setDate(currentDate.getDate() + 1);
      diaIndex++;
    }

    // 2. Normalización de seguridad
    const normalizarNombre = (n) => String(n).trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, ' ');
    let nombreBuscado = normalizarNombre(nombreChofer);

    // 3. Escritura en ID_SPREADSHEET_DIAGRAMAS
    for (let nombreHoja in updatesPorHoja) {
      let hoja = ssDiag.getSheetByName(nombreHoja);
      if (!hoja) continue; 
      
      let dataNombres = hoja.getRange("B6:B" + hoja.getLastRow()).getValues();
      let rowTarget = -1;
      
      for (let i = 0; i < dataNombres.length; i++) {
        if (normalizarNombre(dataNombres[i][0]) === nombreBuscado) { 
          rowTarget = i + 6; 
          break; 
        }
      }
      
      if (rowTarget !== -1) {
        let chunk = updatesPorHoja[nombreHoja];
        hoja.getRange(rowTarget, chunk.startCol, 1, chunk.values.length).setValues([chunk.values]);
      }
    }

    // 4. Registro de Auditoría (Asegurando que el usuario no sea nulo)
    let operador = usuarioActivo ? usuarioActivo : "Usuario No Enviado";
    registrarLog(operador, "CAMBIO_ESTADO", {
      chofer: nombreChofer,
      rango: `${fechaInicioIso} / ${fechaFinIso}`,
      estado: Array.isArray(nuevoEstado) ? "SECUENCIA_MULTIPLE" : nuevoEstado
    });

    generarJSONBase_Frecuente();
    return { success: true };
  } catch (e) { 
    return { success: false, error: e.toString() }; 
  }
}

function actualizarCacheG3_Estaticos() {
  const ssMaestro = SpreadsheetApp.getActiveSpreadsheet();
  
  // Agrupación (Ley de Proximidad): Todas las dependencias de IDs juntas.
  const ID_CENTRAL = '1eQ9Y5diL5fwxYTxvseNgZJFbX-lSUQ13axbp3cLiqPc';
  const ID_LEGAJOS = '19_UPtQYtu7l9zeZPK_glqonxD5jnxXyD8msyy_1lydg';
  // Nota: Asegúrate de que ID_SHEET_TELEFONOS esté declarada de forma global en tus scripts.
  
  let hojaCache = ssMaestro.getSheetByName('API_CACHE_BASICO');
  if (!hojaCache) return;
  
  let mapaDnis = {}; 
  let mapaInfo = {}; 

  // 👉 Helper interno: Parseo de DNI (Ley de Pregnancia: Una función, un propósito claro)
  function _parseDni(val) {
    let limpio = String(val).replace(/\D/g, '');
    if (!limpio) return "";
    if (limpio.length === 11) return String(parseInt(limpio.substring(2, 10), 10));
    if (limpio.length === 10) return String(parseInt(limpio.substring(2, 9), 10));
    return String(parseInt(limpio, 10));
  }

  // --- 1. & 2. FUENTE CENTRAL: DNI y Vacaciones ---
  try {
    const ssCentral = SpreadsheetApp.openById(ID_CENTRAL);
    
    // Nombres y DNIs
    const sheetDni = ssCentral.getSheetByName('dni');
    if (sheetDni) {
      const dataDni = sheetDni.getDataRange().getValues();
      for (let i = 0; i < dataDni.length; i++) {
        let fila = dataDni[i];
        [[0, 2], [5, 7]].forEach(pos => {
          let nom = String(fila[pos[0]] || "").trim().toLowerCase();
          let dni = _parseDni(fila[pos[1]]);
          if (nom && dni && dni !== "NaN") {
            let obj = { dni: dni, vac: 0 };
            mapaDnis[nom] = obj;
            mapaDnis[dni] = obj;
          }
        });
      }
    }

  // --- 2. LECTURA DE VACACIONES (CON FILTRO DE POLÍTICA) ---
    const sheetVac = ssCentral.getSheetByName('vacaciones');
    if (sheetVac) {
      const dataVac = sheetVac.getDataRange().getValues();
      for (let i = 1; i < dataVac.length; i++) {
        
        // 💡 EL SALVAVIDAS: Solo procesamos la fila si es de Vacaciones
        let tipoLicencia = String(dataVac[i][2] || "").toLowerCase();
        if (!tipoLicencia.includes("vacaciones")) continue; 
        
        let nomVac = String(dataVac[i][0] || "").trim().toLowerCase(); 
        let cuilRaw = dataVac[i][1]; 
        
        let dniDesdeCuil = _parseDni(cuilRaw); 
        
        if (dniDesdeCuil) {
          if (!mapaDnis[dniDesdeCuil]) {
            mapaDnis[dniDesdeCuil] = { dni: dniDesdeCuil, vac: 0 };
          }
          
          let disponible = parseFloat(dataVac[i][4]); 
          let saldo = parseFloat(dataVac[i][6]);
          let diasAsignar = !isNaN(saldo) ? saldo : (!isNaN(disponible) ? disponible : 0);
          
          // Usamos += por si el empleado tiene vacaciones de años distintos en filas separadas
          mapaDnis[dniDesdeCuil].vac = (mapaDnis[dniDesdeCuil].vac || 0) + diasAsignar;
          
          if (nomVac) {
             if (!mapaDnis[nomVac]) {
                 mapaDnis[nomVac] = { dni: dniDesdeCuil, vac: mapaDnis[dniDesdeCuil].vac };
             } else {
                 mapaDnis[nomVac].vac = mapaDnis[dniDesdeCuil].vac;
             }
          }
        }
      }
    }      } catch(e) { 
    console.error("Error G3 (DNI/Vacaciones):", e); 
  }

// --- 3. FUENTE LEGAJOS: Datos de Perfil Expandido ---
  try {
    const ssLegajos = SpreadsheetApp.openById(ID_LEGAJOS);
    const sheetLegajos = ssLegajos.getSheetByName('INFORMACION CONDUCTORES');
    if (sheetLegajos) {
      const dataLeg = sheetLegajos.getDataRange().getValues();
      
      for (let i = 1; i < dataLeg.length; i++) {
        let legajo = String(dataLeg[i][0] || "").trim();               // Columna A (Nº)
        
        // 💡 LEY DE PRÄGNANZ (Filtro simple): Si la columna A no es un número, es una cabecera. La ignoramos.
        if (!legajo || isNaN(parseInt(legajo, 10))) continue;

        let nombreRaw = String(dataLeg[i][1] || "").trim();            // Columna B
        let rawDni = String(dataLeg[i][2] || "");                      // Columna C
        let dni = rawDni ? _parseDni(rawDni) : "";
        
        // 💡 EXTRACCIÓN ESTRICTA: El Regex \D elimina palabras como "(INTERNO)" o "(ESPOSA)"
        let telefonoSec = String(dataLeg[i][3] || "").replace(/\D/g, ''); 
        let email = String(dataLeg[i][4] || "").trim();
        let fechaAltaRaw = dataLeg[i][10];
        
        let fechaAltaFormateada = fechaAltaRaw instanceof Date ? 
                                  Utilities.formatDate(fechaAltaRaw, Session.getScriptTimeZone(), "dd/MM/yyyy") : 
                                  String(fechaAltaRaw || "").trim();

        // Normalizamos nombre para el fallback
        let nombreNorm = nombreRaw.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, ' ');

        if (nombreNorm && (dni || legajo)) {
          let infoObj = {
            legajo: legajo,
            dni: dni,
            telefono: telefonoSec,
            email: email,
            fechaAlta: fechaAltaFormateada
          };
          
          // 👉 GUARDADO DOBLE: Guardamos usando DNI como llave maestra, y Nombre como respaldo
          if (dni) mapaInfo[dni] = infoObj;
          mapaInfo[nombreNorm] = infoObj; 
        }
      }
    }
  } catch(e) { 
    console.error("Error G3 Info Legajos:", e); 
  }

  // --- 4. FUENTE TELÉFONOS: Prioridad Principal ---
  try {
    const sheetTel = SpreadsheetApp.openById(ID_SHEET_TELEFONOS).getSheetByName('Telefonos');
    if (sheetTel) {
      const dataTel = sheetTel.getDataRange().getValues();
      for (let i = 1; i < dataTel.length; i++) {
        let nombreNorm = String(dataTel[i][0]).trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, ' ');
        let telRaw = String(dataTel[i][2] || "").trim(); 
        
        // 👉 EXTRACCIÓN ESTRICTA
        let telNumerico = telRaw.replace(/\D/g, ''); 
        
        if (nombreNorm && nombreNorm !== "chofer asignado" && telNumerico.length >= 8) {
          
          if (!mapaInfo[nombreNorm]) mapaInfo[nombreNorm] = {}; 
          mapaInfo[nombreNorm].telefono = telNumerico; 
          
          // Si el objeto ya tenía un DNI (porque vino de la hoja Legajos), actualizamos también el nodo del DNI
          if (mapaInfo[nombreNorm].dni) {
             mapaInfo[mapaInfo[nombreNorm].dni].telefono = telNumerico;
          }
        }
      }
    }
  } catch(e) { 
    console.error("Error G3 Teléfonos:", e); 
  }  
  // --- 5. ALMACENAMIENTO CACHÉ ---
  if (typeof escribirChunksEnFila === 'function') {
    escribirChunksEnFila(hojaCache, 4, JSON.stringify(mapaDnis));
    escribirChunksEnFila(hojaCache, 6, JSON.stringify(mapaInfo)); 
  } else {
    console.error("No se encontró la función escribirChunksEnFila");
  }
  
  ssMaestro.toast("Datos estáticos, Vacaciones y Legajos actualizados.", "G3 OK");
}

function sincronizarBaseYMovimientos() {
  // --- 1. GESTALT: AGRUPACIÓN DE DEPENDENCIAS (Mapeo de JSON en Col D) ---
  const ssUnidades = SpreadsheetApp.openById("1w86w4I-BMcdtANCBYaMwU03cRL_keMvtY8-fvjYAtF8");
  const hojaUnidades = ssUnidades.getSheetByName("unidades"); // Tab de unidades/ute
  let mapaTD = {};
  
  if (hojaUnidades) {
    const datosUnidades = hojaUnidades.getDataRange().getValues();
    // Columna B (índice 1) = Tractor | Columna D (índice 3) = JSON (td + hex)
    for (let i = 1; i < datosUnidades.length; i++) {
      let tractor = String(datosUnidades[i][1] || "").trim();
      let colDVal = String(datosUnidades[i][3] || "").trim();
      
      if (tractor && colDVal) {
        // Estructura base de seguridad
        let datosUnidad = { td: "-", hex1: "", hex2: "" }; 
        
        // Comprobamos si la celda contiene un JSON (empieza con '{')
        if (colDVal.startsWith("{")) {
          try {
            let parsed = JSON.parse(colDVal);
            // Combinamos los datos. '...parsed' asegura que extraiga td, hex1, hex2
            // y cualquier otra variable que hayas sumado en el JSON de la DB.
            datosUnidad = {
              td: parsed.td || parsed.ute || "-", 
              ...parsed 
            };
          } catch(e) {
            console.warn(`Error parseando JSON para el tractor ${tractor}:`, e);
            datosUnidad.td = colDVal; // Fallback por si hay texto roto
          }
        } else {
          datosUnidad.td = colDVal; // Fallback por si quedan celdas antiguas
        }
        
        mapaTD[tractor] = datosUnidad;
      }
    }
  }

  // --- 2. LECTURA DEL DIAGRAMA ---
  const ssDiag = SpreadsheetApp.openById(ID_SPREADSHEET_DIAGRAMAS);
  const hoy = new Date();
  const mesesAbrev = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];
  const nombreHojaActual = mesesAbrev[hoy.getMonth()] + "-" + String(hoy.getFullYear()).slice(-2);
  
  const hojaDiag = ssDiag.getSheetByName(nombreHojaActual);
  if (!hojaDiag) return;
  const datosDiag = hojaDiag.getDataRange().getValues();
  
  let diccionarioCruce = {};
  for (let i = 5; i < datosDiag.length; i++) {
    let nombre = String(datosDiag[i][1]).trim();   
    let servicio = String(datosDiag[i][2]).trim(); 
    if (nombre !== "" && nombre !== "APELLIDO Y NOMBRE") {
      let key = nombre.toLowerCase();
      // Prägnanz: Estructura visualmente clara, incluyendo los keys para frontend
      diccionarioCruce[key] = { 
        nombre: nombre, 
        tractor: "", 
        semi: "", 
        servicio: servicio, 
        n_ute: "", 
        td: "-",
        hex1: "", 
        hex2: ""  
      };
    }
  }

  // --- 3. LECTURA DE MOVIMIENTOS ---
  const ssMov = SpreadsheetApp.openById(ID_SHEET_MOVIMIENTOS);
  let hojaMov = null;
  for (let sheet of ssMov.getSheets()) {
    if (sheet.getName().includes("Mov.Unidades y Choferes")) { hojaMov = sheet; break; }
  }
  if (!hojaMov) return; 
  
  const datosMov = hojaMov.getDataRange().getValues();
  
  let targetD = hoy.getDate(); 
  let targetM = hoy.getMonth(); 
  let targetY = hoy.getFullYear();
  const mesesLargo = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];
  
  let regexFechaTexto = new RegExp(`\\b0?${targetD}\\s+${mesesLargo[targetM]}\\s+${targetY}\\b`, 'i');
  let colFecha = -1; 

  for (let c = 0; c < datosMov[0].length; c++) {
    let cellVal = datosMov[0][c];
    if (cellVal instanceof Date) {
      if (cellVal.getDate() === targetD && cellVal.getMonth() === targetM && cellVal.getFullYear() === targetY) { colFecha = c; break; }
    } else if (cellVal) {
      let strVal = String(cellVal).toLowerCase().trim();
      if (regexFechaTexto.test(strVal)) { colFecha = c; break; }
    }
  }

  if (colFecha === -1 || colFecha < 3) return; 

  let colNombreActivos = colFecha - 3; 

  for (let i = 2; i < datosMov.length; i++) {
    let nombreMovOriginal = String(datosMov[i][colNombreActivos] || "").trim();
    if (!nombreMovOriginal || nombreMovOriginal === "1") continue;
    if (!/[a-zA-ZáéíóúÁÉÍÓÚñÑ]/.test(nombreMovOriginal)) continue;

    let keyMov = nombreMovOriginal.toLowerCase();
    
    if (diccionarioCruce[keyMov]) {
      let tractorAsignado = String(datosMov[i][4] || "").trim();
      
      diccionarioCruce[keyMov].n_ute = String(datosMov[i][2] || "").trim(); 
      diccionarioCruce[keyMov].tractor = tractorAsignado; 
      diccionarioCruce[keyMov].semi = String(datosMov[i][5] || "").trim(); 
      
      // NUEVO: Transferimos dinámicamente todo el JSON (td, hex1, hex2) al chofer
      if (tractorAsignado && mapaTD[tractorAsignado]) {
        let datosExtraUnidad = mapaTD[tractorAsignado];
        Object.assign(diccionarioCruce[keyMov], datosExtraUnidad);
      }
    }
  }

  // --- 4. EXPORTACIÓN ---
  let flotaArrayJSON = [];
  for (let key in diccionarioCruce) { flotaArrayJSON.push(diccionarioCruce[key]); }
  
  const ssMaestro = SpreadsheetApp.getActiveSpreadsheet();
  let hojaDestino = ssMaestro.getSheetByName('choferes y unidades');
  if (!hojaDestino) hojaDestino = ssMaestro.insertSheet('choferes y unidades');
  
  hojaDestino.clearContents(); 
  // Se renderiza el array final enriquecido en H1[cite: 1]
  hojaDestino.getRange("H1").setValue(JSON.stringify(flotaArrayJSON));
  
  generarJSONBase_Frecuente(); 
}
function guardarNuevaObservacion(datosFormulario) {
  try {
    const sheet = SpreadsheetApp.openById(ID_SHEET_OBSERVACIONES).getSheetByName('Movimientos');
    let nuevaFila = [ datosFormulario.admin, datosFormulario.chofer, datosFormulario.fecha, datosFormulario.unidad, datosFormulario.evento, datosFormulario.obsEvento, datosFormulario.estado, datosFormulario.obsEstado, "","","","","","","","" ];
    sheet.appendRow(nuevaFila);
    
    generarJSONObservacionesGlobal(); 
    generarJSONBase_Frecuente();      
    return { success: true, message: 'Registrado con éxito' };
  } catch (error) { return { success: false, message: error.toString() }; }
}
function actualizarCacheViajesCampo() {
  const ID_SHEET_KILOMETROS = '1Wr-_P4mDvldif_cAx08sp7yT8uTUrajI2HQAJF6tnGM';
  const ssPrincipal = SpreadsheetApp.getActiveSpreadsheet(); 
  
  try {
    const ssViajes = SpreadsheetApp.openById(ID_SHEET_KILOMETROS);
    const sheetViajes = ssViajes.getSheetByName("KM"); 
    
    if (!sheetViajes) {
      ssPrincipal.toast("No se encontró la pestaña 'KM' en el archivo de Kilómetros.", "❌ Error", -1);
      return;
    }

    const data = sheetViajes.getDataRange().getValues();
    if (data.length <= 1) {
      ssPrincipal.toast("La pestaña 'KM' está vacía.", "⚠️ Atención", -1);
      return;
    }

    let viajesAgrupados = {};
    const mesesAbrev = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];
    let contadorLogs = 0;

    // Helper para procesar números de forma segura y limpia
    const parseNumberSafe = (val) => {
        if (typeof val === 'number') return val;
        let limpio = String(val || '').replace(/,/g, '.').replace(/[^0-9.-]/g, '');
        return parseFloat(limpio) || 0;
    };

    for (let i = 1; i < data.length; i++) {
      let fechaRaw = data[i][1];    // Col B (Fecha)
      let nombreRaw = data[i][2];   // Col C (Nombre)
      let livianoRaw = data[i][3];  // Col D (Liviano)
      let euroRaw = data[i][4];     // Col E (Euro)
      let kmRaw = data[i][5];       // Col F (Valor/KM)
      let infiniaDRaw = data[i][7]; // Col H (Infinia D)
      
      if (!nombreRaw || !fechaRaw) continue;

      // Normalización para match con Frontend
      let nombre = String(nombreRaw).trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, ' ');
      
      let mesAnio = "";
      let dateObj = null;

      // Procesamiento de Fecha
      if (fechaRaw instanceof Date) {
        dateObj = fechaRaw;
      } else if (typeof fechaRaw === 'string') {
        let parts = fechaRaw.trim().split(/[\/\-]/);
        if (parts.length >= 3) {
          if(parts[0].length <= 2 && parts[2].length === 4) { 
            dateObj = new Date(parts[2], parseInt(parts[1], 10) - 1, parts[0]);
          } else if (parts[0].length === 4) { 
            dateObj = new Date(parts[0], parseInt(parts[1], 10) - 1, parts[2]);
          }
        }
      }
      
      if (dateObj && !isNaN(dateObj.getTime())) {
          let m = mesesAbrev[dateObj.getMonth()];
          let y = String(dateObj.getFullYear()).slice(-2);
          mesAnio = `${m}-${y}`; // Formato: "Abr-26"
      }
      
      if (!mesAnio) continue;

      // Procesamiento de Valores numéricos usando el Helper
      let kmNum = parseNumberSafe(kmRaw);
      let livianoNum = parseNumberSafe(livianoRaw);
      let euroNum = parseNumberSafe(euroRaw);
      let infiniaDNum = parseNumberSafe(infiniaDRaw);

      // Si al menos uno de los campos tiene un valor mayor a 0, guardamos el registro
      if (kmNum > 0 || livianoNum > 0 || euroNum > 0 || infiniaDNum > 0) {
        
        // Inicializamos el chofer si no existe
        if (!viajesAgrupados[nombre]) {
            viajesAgrupados[nombre] = {};
        }
        
        // Inicializamos el mes como un OBJETO con los campos si no existe
        if (!viajesAgrupados[nombre][mesAnio]) {
            viajesAgrupados[nombre][mesAnio] = {
                km: 0,
                liviano: 0,
                euro: 0,
                infiniaD: 0
            };
        }
        
        // Sumamos a cada categoría
        viajesAgrupados[nombre][mesAnio].km += kmNum;
        viajesAgrupados[nombre][mesAnio].liviano += livianoNum;
        viajesAgrupados[nombre][mesAnio].euro += euroNum;
        viajesAgrupados[nombre][mesAnio].infiniaD += infiniaDNum;
        
        contadorLogs++;
      }
    }

  // Guardado en Cache (Usando Chunks Horizontales en Fila 7)
    let cacheSheet = ssPrincipal.getSheetByName("API_CACHE_BASICO");
    if (cacheSheet) {
      // Usamos tu helper para trocear el JSON y pegarlo a lo largo de la fila 7
      escribirChunksEnFila(cacheSheet, 7, JSON.stringify(viajesAgrupados));
      
      ssPrincipal.toast(`Cache de Viajes Campo actualizada (${contadorLogs} registros con combustibles).`, "✅ Éxito");
    }

  } catch (error) {
    ssPrincipal.toast("Error: " + error.message, "❌ Error", -1);
  }
}
function obtenerDatosTDParaFront() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaTD = ss.getSheetByName('td');
  if (!hojaTD) return { campo:{}, infinia:{}, liviano:{}, euro:{}, estados:{} };

    const leerFila = (numFila) => {
      let values = hojaTD.getRange(numFila + ":" + numFila).getValues()[0];
      let jsonStr = values.filter(String).map(c => String(c).replace(/^'/, "")).join("");
      try { return jsonStr ? JSON.parse(jsonStr) : {}; } catch(e) { return {}; }
    };

  return {
    campo: leerFila(1),
    infinia: leerFila(2),
    liviano: leerFila(3),
    euro: leerFila(4),
    estados: leerFila(12) // Respuestas de los checkboxes
  };
}

function guardarEstadoCheckboxTD(tdId, estado, codigosExtra, usuarioResponsable) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const hoja = ss.getSheetByName('td');
      if (!hoja) throw new Error("No se encontró la pestaña 'td'");

      // 1. Manejar Estados (Fila A5)
      let estados = JSON.parse(hoja.getRange("A5").getValue() || "{}");
      estados[tdId] = estado;
      hoja.getRange("A5").setValue(JSON.stringify(estados));

      // 2. Manejar Códigos (Fila A6)
      let todosLosCodigos = JSON.parse(hoja.getRange("A6").getValue() || "{}");
      
      if (estado === true && codigosExtra && codigosExtra.length > 0) {
        todosLosCodigos[tdId] = codigosExtra;
      } else {
        // Si se apaga el checkbox, limpiamos los códigos para no ocupar espacio
        delete todosLosCodigos[tdId];
      }
      hoja.getRange("A6").setValue(JSON.stringify(todosLosCodigos));

      // 👉 NUEVO LOG: Registra explícitamente el estado del checkbox y los códigos
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

// 1. EL INSTALADOR (Ejecutar manualmente SOLO UNA VEZ)

function instalarTriggerDiagramas() {
  const ID_DIAGRAMAS = ID_SPREADSHEET_DIAGRAMAS; // Usa tu variable global existente
  
  // Borramos triggers anteriores para evitar duplicados si lo corres por error dos veces
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'alEditarDiagramas') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // Creamos el nuevo "escuchador" apuntando al archivo externo
  ScriptApp.newTrigger('alEditarDiagramas')
    .forSpreadsheet(ID_DIAGRAMAS)
    .onEdit()
    .create();
    
  SpreadsheetApp.getActiveSpreadsheet().toast("Trigger de Diagramas instalado OK", "Éxito");
}
// 2. EL ESCUCHADOR (Se dispara automáticamente al editar)
function alEditarDiagramas(e) {
  // Filtro de seguridad
  if (!e || !e.range) return;

  const fila = e.range.getRow();
  const columna = e.range.getColumn();
  const nombreHoja = e.source.getActiveSheet().getName();

  // 👉 LEY DE PRÄGNANZ (Filtro Estricto): 
  // Solo actualizamos si la edición fue en el área útil del diagrama.
  // Fila >= 6 (Donde empiezan los choferes)
  // Columna >= 5 (Donde empiezan los días del mes, Columna E)
  // Verificamos que el nombre de la hoja tenga formato de mes (Ej: "May-26")
  const regexMes = /^[A-Z][a-z]{2}-\d{2}$/; 

  if (fila >= 6 && columna >= 5 && regexMes.test(nombreHoja)) {
    
    // Si la edición fue en un día del mes válido, disparamos la actualización del caché G1
    try {
      // Usamos generarJSONBase_Frecuente porque actualiza el mes pasado, actual y los dos siguientes (la ventana más probable de edición)
      generarJSONBase_Frecuente();
    } catch (error) {
      console.error("Error al sincronizar desde trigger:", error);
    }
  }
}
function vencimientosUnidades() {
  // 1. Referencias y Conexiones (Ley de Proximidad)
  const ID_MASTER = '1eQ9Y5diL5fwxYTxvseNgZJFbX-lSUQ13axbp3cLiqPc';
  
  // Abrimos el archivo de Movimientos usando la constante global existente
  const ssMovimientos = SpreadsheetApp.openById(ID_SHEET_MOVIMIENTOS);
  const hojaVencimientos = ssMovimientos.getSheetByName('Vencimientos.');
  
  if (!hojaVencimientos) {
    console.error("UX/UI Feedback: No se encontró la pestaña 'Vencimientos.' en el archivo origen.");
    return { success: false, error: "Pestaña 'Vencimientos.' no encontrada" };
  }

  // 2. Extracción de Datos
  // CAMBIO CLAVE: Usamos getDisplayValues() para capturar el texto visualizado en la celda, 
  // eliminando la necesidad de parsear fechas largas (ISO strings).
  const data = hojaVencimientos.getDataRange().getDisplayValues();
  let resultados = [];

  // 3. Procesamiento (Iniciamos en row 2 -> índice 1)
  for (let i = 1; i < data.length; i++) {
    let fila = data[i];
    
    // Validación básica: Solo procesar si la columna B (índice 1) tiene datos
    if (String(fila[1]).trim() !== "") {
      
      // Mapeo de columnas: B=1, C=2, G=6, H=7, J=9, K=10, L=11, M=12, N=13
      // Al usar getDisplayValues, estos datos ya son strings limpios (ej: "15/06/2026")
      resultados.push({
        col_b: fila[1], 
        col_c: fila[2], 
        col_g: fila[6], 
        col_h: fila[7], 
        col_j: fila[9], 
        col_k: fila[10],
        col_l: fila[11],
        col_m: fila[12],
        col_n: fila[13]
      });
    }
  }

  // 4. Guardado en Caché Master
  try {
    const ssMaster = SpreadsheetApp.openById(ID_MASTER);
    const hojaCache = ssMaster.getSheetByName('API_CACHE_BASICO');
    
    if (!hojaCache) {
      throw new Error("Pestaña 'API_CACHE_BASICO' no encontrada");
    }

    // Reutilizamos el helper existente para guardar el JSON de forma segura
    escribirChunksEnFila(hojaCache, 11, JSON.stringify(resultados));    
    
    // UX/UI Feedback: Notificación visual en el Google Sheet
    ssMaster.toast("Vencimientos de unidades actualizados con éxito.", "Sincronización OK");
    return { success: true };
    
  } catch (error) {
    console.error("Error al guardar vencimientos:", error);
    return { success: false, error: error.message };
  }
}

/**
 * Busca y mapea los vencimientos para la unidad seleccionada.
 * @param {string} patente - La patente del tractor (ej. AF071WS)
 * @param {Array} jsonVencimientos - El array parseado del API_CACHE_BASICO
 */
function mapearVencimientos(patente, jsonVencimientos) {
    if (!patente || !jsonVencimientos) return null;

    // Sanitización estricta: Eliminar espacios en blanco de la patente buscada
    const patenteLimpia = String(patente).replace(/\s+/g, '').toUpperCase();

    // Buscar la fila correspondiente (Asumiendo que col_b es la patente del Tractor)
    const record = jsonVencimientos.find(item => 
        String(item.col_b).replace(/\s+/g, '').toUpperCase() === patenteLimpia
    );

    if (!record) return null;

    // Mapear las columnas "opacas" a claves legibles para la UI
    // Ajusta las columnas según la estructura real de tu pestaña 'Vencimientos.'
    return {
        vtv_tr: record.col_g,  // VTV Tractor
        vtv_se: record.col_h,  // VTV Semirremolque
        mass_tr: record.col_j, // MASS Tractor
        mass_se: record.col_k, // MASS Semirremolque
        esp_es_tr: record.col_l, // ESP-ES
        vi_tr: record.col_m,   // VI
        ve_tr: record.col_n    // VE
    };
}

/**
 * Motor de Auditoría.
 * @param {string} usuarioResponsable - El ID/Nombre del usuario que envía el frontend.
 * @param {string} accion - Qué hizo (ej. "CAMBIO_ESTADO").
 * @param {object} detalles - El JSON con el payload de la acción.
 */
function registrarLog(usuarioResponsable, accion, detalles) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetLogs = ss.getSheetByName("DB_Logs") || ss.insertSheet("DB_Logs");

  // Si por algún error de red no llega el usuario, evitamos que el log se rompa.
  const usuario = usuarioResponsable || "SESION_PERDIDA";

  const nuevoEvento = { 
    t: new Date().toISOString(), 
    a: accion, 
    d: detalles 
  };
  
  // Buscar la fila correspondiente a este usuario en Columna A
  const data = sheetLogs.getDataRange().getValues();
  let rowIndex = data.findIndex(row => String(row[0]).trim() === String(usuario).trim()) + 1;
  
  // Si no existe en DB_Logs (primera vez que hace una acción), lo creamos
  if (rowIndex === 0) {
    sheetLogs.appendRow([usuario, JSON.stringify([nuevoEvento])]);
    return;
  }

  // Lógica de Chunks: Fragmentación horizontal (Límite 45k)
  let colIndex = 2; 
  let cell = sheetLogs.getRange(rowIndex, colIndex);
  let content = cell.getValue();

  while (content && content.toString().length > 44000) {
    colIndex++;
    cell = sheetLogs.getRange(rowIndex, colIndex);
    content = cell.getValue();
  }

  let logArray = content ? JSON.parse(content) : [];
  logArray.push(nuevoEvento);
  cell.setValue(JSON.stringify(logArray));
}

/**
 * ============================================================================
 * ESCRITURA BIDIRECCIONAL: HOJA DE RUTA EN PLANILLA "KM"
 * ============================================================================
 */
function procesarGuardadoHojaRuta(nombre, fechaIso, hojasArr, usuario, tractor) {
  try {
    const ID_SHEET_KILOMETROS = '1Wr-_P4mDvldif_cAx08sp7yT8uTUrajI2HQAJF6tnGM';
    const ssKm = SpreadsheetApp.openById(ID_SHEET_KILOMETROS); 
    const sheetKm = ssKm.getSheetByName('KM') || ssKm.getSheets()[0];
    
    if (!sheetKm) return { success: false, error: "Pestaña KM no encontrada" };

    const data = sheetKm.getDataRange().getValues();
    
    // Parseo de fecha para buscar
    let fBase = new Date(fechaIso + "T12:00:00");
    let tD = String(fBase.getDate()).padStart(2, '0');
    let tM = String(fBase.getMonth() + 1).padStart(2, '0');
    let tY = String(fBase.getFullYear()).slice(-2);
    let shortDate = `${tD}/${tM}/${tY}`; // Ej: 15/06/26
    
    const normalizar = (n) => String(n).trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, ' ');
    let nBuscado = normalizar(nombre);
    let stringHojas = hojasArr.join(', ');
    let actualizados = 0;

    // Buscamos si la fila ya existe
    for (let i = 1; i < data.length; i++) {
        let nFila = normalizar(data[i][2]); 
        if (nFila === nBuscado) {
            let dFila = data[i][1]; 
            let isMatch = false;
            
            if (dFila instanceof Date) {
                if (dFila.getDate() === fBase.getDate() && dFila.getMonth() === fBase.getMonth() && dFila.getFullYear() === fBase.getFullYear()) {
                    isMatch = true;
                }
            } else {
                let strD = String(dFila).trim();
                if (strD.startsWith(shortDate) || strD.startsWith(fechaIso) || strD.includes(shortDate)) {
                    isMatch = true;
                }
            }

            if (isMatch) {
                sheetKm.getRange(i + 1, 20).setValue(stringHojas); // Columna T = 20
                actualizados++;
            }
        }
    }
    
    // 👉 LÓGICA DE INSERCIÓN (Si el usuario carga HR pero la fila aún no existía en KM)
    if (actualizados === 0 && stringHojas !== "") {
        let numColumnas = Math.max(20, sheetKm.getLastColumn());
        let nuevaFila = new Array(numColumnas).fill("");
        nuevaFila[0] = tractor || "";  // Col A: Dominio
        nuevaFila[1] = shortDate;      // Col B: Fecha
        nuevaFila[2] = nombre;         // Col C: Chofer
        nuevaFila[19] = stringHojas;   // Col T: Hoja de Ruta
        
        sheetKm.appendRow(nuevaFila);
        actualizados++;
    }
    
    // Dejamos un Log del cambio
    registrarLog(usuario, "EDITO_HOJA_RUTA", {
        chofer: nombre,
        fecha_viaje: fechaIso,
        hojas_cargadas: hojasArr
    });

    // Actualizamos el Caché 12 para los próximos que entren al sistema
    if (typeof generarJSONKilometros_Frecuente === 'function') {
        generarJSONKilometros_Frecuente();
    }

    return { success: true, filas_actualizadas: actualizados };
    
  } catch(e) {
      return { success: false, error: e.toString() };
  }
}

// Función que le avisa a Node.js que alguien editó el Excel
function notificarBackendNode(tipoCambio) {
  // 👉 AQUÍ PEGAS TU URL DE RENDER + LA RUTA DEL WEBHOOK
  const URL_NODE_WEBHOOK = "https://diagramasnode.onrender.com/api/webhook/google";
  
  try {
    UrlFetchApp.fetch(URL_NODE_WEBHOOK, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ evento: tipoCambio }),
      muteHttpExceptions: true 
    });
  } catch (e) {
    console.error("Error avisando a Node: " + e);
  }
}
