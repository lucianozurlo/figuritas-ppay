function agregarColumnasFecha() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(h => String(h).trim());

  ['fecha_subida', 'fecha_generacion'].forEach(nombre => {
    if (headers.includes(nombre)) {
      Logger.log('Ya existe: ' + nombre);
      return;
    }
    const nuevaCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, nuevaCol)
      .setValue(nombre)
      .setFontWeight('bold')
      .setBackground('#E8EAED');
    Logger.log('Agregada: ' + nombre + ' en columna ' + nuevaCol);
    headers.push(nombre); // para que la segunda iteración la vea
  });
}