function testEnvioCompleto() {
  Logger.log("=== TEST ENVÍO COMPLETO ===");

  // 1. Verificar columnas del Sheet
  const sheet = getSheet_();
  const colMap = getColumnMap_(sheet);
  Logger.log("Columnas detectadas: " + JSON.stringify(Object.keys(colMap)));

  // 2. Verificar que cada CONFIG.COLUMNS existe en el Sheet
  Object.entries(CONFIG.COLUMNS).forEach(([key, colName]) => {
    const existe = colMap[colName] ? "OK col " + colMap[colName] : "❌ FALTA";
    Logger.log(`  ${key} → "${colName}": ${existe}`);
  });

  // 3. Intentar escribir una fila de prueba
  Logger.log("--- Escribiendo fila de prueba ---");
  try {
    const rowIndex = escribirFila_({
      nombre: "TEST Diagnostico",
      mail: "diag" + new Date().getTime() + "@test.com",
      area: "Producto",
      superpoder: "Detectar bugs",
      actitud: "Persistente",
      consentimientoMural: "Sí",
      idArchivoDrive: "fake123",
      urlArchivoDrive: "https://fake",
      estado: "PROCESANDO",
    });
    Logger.log("Fila escrita en índice: " + rowIndex);

    // 4. Releer esa fila para ver qué quedó
    const data = sheet
      .getRange(rowIndex, 1, 1, sheet.getLastColumn())
      .getValues()[0];
    Logger.log("Contenido de la fila: " + JSON.stringify(data));
  } catch (e) {
    Logger.log("❌ ERROR al escribir: " + e.message);
  }

  Logger.log("=== FIN ===");
}
