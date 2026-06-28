function crearEncabezados() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    CONFIG.SHEET_NAME,
  );

  const headers = [
    "Nombre",
    "Mail",
    "Area",
    "Superpoder",
    "Actitud",
    "Consentimiento mural",
    "id_archivo_drive",
    "url_archivo_drive",
    "id_figurita",
    "id_figurita_generada",
    "url_figurita_generada",
    "estado",
    "detalle_error",
    "timestamp_procesando",
  ];

  // 1. Primero limpiar validaciones de la fila 1 ANTES de escribir
  sheet.getRange(1, 1, 1, headers.length).clearDataValidations();

  // 2. Recién ahora escribir los encabezados
  sheet
    .getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight("bold")
    .setBackground("#E8EAED");

  Logger.log("✅ " + headers.length + " encabezados creados correctamente.");
}
