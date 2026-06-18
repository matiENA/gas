function obtenerNombresMesActual() {
  try {
    const idPlanilla = '1mhfXpFCF6upMlnRnZjDdBVS_wqTx5q8v0qQArNCnNAU';
    const ss = SpreadsheetApp.openById(idPlanilla);
    
    // Calcular el nombre de la pestaña del mes actual (ej: "Feb-25")
    const mesesAbrev = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];
    const hoy = new Date();
    const nombrePestana = mesesAbrev[hoy.getMonth()] + "-" + String(hoy.getFullYear()).slice(-2);
    
    let sheet = ss.getSheetByName(nombrePestana);
    
    // Si la pestaña del mes actual no existe aún, tomamos la primera pestaña como respaldo
    if (!sheet) {
      sheet = ss.getSheets()[0];
    }
    
    const data = sheet.getDataRange().getValues();
    let nombresEncontrados = [];
    
    // Recorremos la planilla buscando los nombres en la Columna B (índice 1)
    for (let i = 1; i < data.length; i++) {
      let nom = String(data[i][1] || "").trim(); // 👉 PRIORIDAD COLUMNA B (Índice 1)
      if (!nom) nom = String(data[i][0] || "").trim(); // Respaldo Columna A (Índice 0)
      
      if (nom && nom.length > 3 && nom.toUpperCase() !== "CHOFER") {
        nombresEncontrados.push(nom);
      }
    }
    
    // Retornamos la lista sin duplicados
    return JSON.stringify([...new Set(nombresEncontrados)]);
  } catch (error) {
    console.error("Error leyendo planilla de mes actual:", error);
    return JSON.stringify([]);
  }
}
