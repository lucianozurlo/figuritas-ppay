function testBlob() {
  const sheet  = getSheet_();
  const colMap = getColumnMap_(sheet);
  const data   = sheet.getDataRange().getValues();
  
  // Tomar la primera figurita con consentimiento de mural
  const colFigId  = colMap[CONFIG.COLUMNS.idFiguraGenerada];
  const colMural  = colMap[CONFIG.COLUMNS.consentimientoMural];
  const colEstado = colMap[CONFIG.COLUMNS.estado];
  
  const fila = data.slice(1).find(row => {
    const mural  = String(row[colMural  - 1] || '').trim();
    const fileId = String(row[colFigId  - 1] || '').trim();
    const estado = String(row[colEstado - 1] || '').trim();
    return (mural === 'Sí') && fileId && (estado === 'EMAIL_ENVIADO' || estado === 'FIGURITA_CREADA');
  });
  
  if (!fila) { Logger.log('No se encontró ninguna figurita elegible.'); return; }
  
  const fileId = String(fila[colFigId - 1]).trim();
  Logger.log('fileId: ' + fileId);
  
  try {
    const file = DriveApp.getFileById(fileId);
    Logger.log('Nombre: ' + file.getName());
    Logger.log('MimeType: ' + file.getMimeType());
    Logger.log('Tamaño: ' + file.getSize() + ' bytes');
    
    const blob = file.getBlob();
    Logger.log('Blob OK: ' + blob.getBytes().length + ' bytes');
    Logger.log('Blob type: ' + blob.getContentType());
  } catch(e) {
    Logger.log('ERROR: ' + e.message);
  }
}